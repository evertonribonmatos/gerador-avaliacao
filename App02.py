import os
import re
import json
import time
import tempfile
from io import BytesIO
from typing import List, Dict, Any, Optional, Tuple

import pdfplumber
import streamlit as st
from dotenv import load_dotenv
from groq import Groq
from PIL import Image
import pytesseract
from docx import Document
from docx.table import _Cell, Table
from docx.shared import Inches
from docx.oxml.ns import qn


# ==========================================
# CONFIGURAÇÕES
# ==========================================
TOTAL_QUESTOES = 10
MODELO_PRINCIPAL = "openai/gpt-oss-120b"
LARGURA_IMAGEM_POLEGADAS = 4.8

TEMPERATURA_GERACAO = 0.45
MAX_TENTATIVAS_MODELO = 3
MAX_TENTATIVAS_POR_QUESTAO = 5

TAMANHO_MINIMO_CONTEXTO = 180
SIMILARIDADE_MAXIMA_PERMITIDA = 0.72

MARCADORES_QUESTOES = {
    1: "primeira",
    2: "segunda",
    3: "terceira",
    4: "quarta",
    5: "quinta",
    6: "sexta",
    7: "sétima",
    8: "oitava",
    9: "nona",
    10: "décima",
}

CAMINHO_TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
if os.path.exists(CAMINHO_TESSERACT):
    pytesseract.pytesseract.tesseract_cmd = CAMINHO_TESSERACT


# ==========================================
# SEGREDOS / CLIENTE IA
# ==========================================
load_dotenv()


def obter_groq_api_key() -> str:
    try:
        if "GROQ_API_KEY" in st.secrets:
            return st.secrets["GROQ_API_KEY"]
    except Exception:
        pass

    chave_env = os.getenv("GROQ_API_KEY")
    if chave_env:
        return chave_env

    raise ValueError("GROQ_API_KEY não encontrada. Configure em st.secrets ou no arquivo .env.")


GROQ_API_KEY = obter_groq_api_key()
client = Groq(api_key=GROQ_API_KEY)


# ==========================================
# UTILITÁRIOS
# ==========================================
def normalizar_texto(texto: Any) -> str:
    if texto is None:
        return ""
    texto = str(texto).replace("\r", "\n")
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


def remover_acentos(texto: str) -> str:
    substituicoes = {
        "á": "a", "à": "a", "ã": "a", "â": "a",
        "é": "e", "ê": "e",
        "í": "i",
        "ó": "o", "ô": "o", "õ": "o",
        "ú": "u",
        "ç": "c",
        "Á": "A", "À": "A", "Ã": "A", "Â": "A",
        "É": "E", "Ê": "E",
        "Í": "I",
        "Ó": "O", "Ô": "O", "Õ": "O",
        "Ú": "U",
        "Ç": "C",
    }
    for origem, destino in substituicoes.items():
        texto = texto.replace(origem, destino)
    return texto


def normalizar_texto_comparacao(texto: Any) -> str:
    texto = normalizar_texto(texto).lower()
    texto = remover_acentos(texto)
    return texto


def extrair_json_de_texto(texto: str) -> Dict[str, Any]:
    texto = texto.strip()
    try:
        return json.loads(texto)
    except json.JSONDecodeError:
        pass

    match = re.search(r"\{.*\}", texto, re.DOTALL)
    if match:
        return json.loads(match.group(0))

    raise ValueError("Não foi possível extrair JSON válido da resposta da IA.")


def chamar_ia_com_retry(**kwargs):
    ultima_excecao = None
    for tentativa in range(MAX_TENTATIVAS_MODELO):
        try:
            return client.chat.completions.create(**kwargs)
        except Exception as e:
            ultima_excecao = e
            espera = 2 ** tentativa
            time.sleep(espera)
    raise ultima_excecao


def uploaded_file_para_bytes(uploaded_file) -> Optional[bytes]:
    if uploaded_file is None:
        return None
    uploaded_file.seek(0)
    return uploaded_file.read()


def tentar_ocr_em_imagem(imagem_bytes: Optional[bytes]) -> str:
    if not imagem_bytes:
        return ""

    try:
        imagem = Image.open(BytesIO(imagem_bytes))
        texto = pytesseract.image_to_string(imagem, lang="por+eng")
        return normalizar_texto(texto)
    except Exception:
        return ""


def tokenizar_simples(texto: str) -> List[str]:
    return re.findall(r"\w+", normalizar_texto(texto).lower())


def similaridade_textual_simples(a: str, b: str) -> float:
    a_tokens = set(tokenizar_simples(a))
    b_tokens = set(tokenizar_simples(b))

    if not a_tokens or not b_tokens:
        return 0.0

    inter = len(a_tokens & b_tokens)
    uniao = len(a_tokens | b_tokens)
    return inter / uniao if uniao else 0.0


def sanitizar_contexto_gerado(contexto: str, possui_imagem: bool) -> str:
    texto = normalizar_texto(contexto)

    substituicoes = [
        (r"\bcom base no relatório\b", "com base nas informações do cenário"),
        (r"\bde acordo com o relatório\b", "de acordo com as informações apresentadas"),
        (r"\bconforme o relatório\b", "conforme as informações do caso"),
        (r"\banalise o relatório\b", "analise a situação apresentada"),
        (r"\bo relatório apresenta\b", "o cenário apresenta"),
        (r"\bsegundo o relatório\b", "segundo as informações apresentadas"),
        (r"\bcom base na tabela\b", "com base nos dados apresentados no enunciado"),
        (r"\bde acordo com a tabela\b", "de acordo com os dados apresentados no enunciado"),
        (r"\bconforme a tabela\b", "conforme os dados apresentados"),
        (r"\bobserve a tabela\b", "analise os dados apresentados"),
        (r"\bobserve o gráfico\b", "analise o comportamento descrito"),
        (r"\bcom base no gráfico\b", "com base no comportamento descrito"),
        (r"\bde acordo com o gráfico\b", "de acordo com o comportamento descrito"),
        (r"\bconforme o gráfico\b", "conforme o comportamento descrito"),
        (r"\bcom base no laudo\b", "com base nas informações técnicas fornecidas"),
        (r"\bde acordo com o laudo\b", "de acordo com as informações técnicas fornecidas"),
        (r"\bconforme o laudo\b", "conforme as informações técnicas fornecidas"),
        (r"\bem anexo\b", ""),
        (r"\bno anexo\b", ""),
        (r"\bprontuário\b", "registro técnico"),
        (r"\bprontuario\b", "registro técnico"),
        (r"\bplanilha\b", "registro de dados"),
    ]

    for padrao, repl in substituicoes:
        texto = re.sub(padrao, repl, texto, flags=re.IGNORECASE)

    if not possui_imagem:
        substituicoes_visuais = [
            (r"\bobserve a figura\b", "considere a situação apresentada"),
            (r"\bcom base na figura\b", "com base na situação apresentada"),
            (r"\bde acordo com a figura\b", "de acordo com a situação apresentada"),
            (r"\bconforme a figura\b", "conforme a situação apresentada"),
            (r"\bobserve a imagem\b", "considere a situação apresentada"),
            (r"\bcom base na imagem\b", "com base na situação apresentada"),
            (r"\bde acordo com a imagem\b", "de acordo com a situação apresentada"),
            (r"\bconforme a imagem\b", "conforme a situação apresentada"),
            (r"\bobserve o diagrama\b", "considere a situação apresentada"),
            (r"\bcom base no diagrama\b", "com base na situação apresentada"),
            (r"\bde acordo com o diagrama\b", "de acordo com a situação apresentada"),
            (r"\bconforme o diagrama\b", "conforme a situação apresentada"),
        ]
        for padrao, repl in substituicoes_visuais:
            texto = re.sub(padrao, repl, texto, flags=re.IGNORECASE)

    texto = re.sub(r"\s{2,}", " ", texto)
    return normalizar_texto(texto)


def detectar_referencia_indevida(contexto: str, possui_imagem: bool) -> Optional[str]:
    texto = normalizar_texto(contexto).lower()

    padroes_indevidos = [
        r"\bcom base no relatório\b",
        r"\bde acordo com o relatório\b",
        r"\bconforme o relatório\b",
        r"\banalise o relatório\b",
        r"\bo relatório apresenta\b",
        r"\bsegundo o relatório\b",
        r"\bcom base na tabela\b",
        r"\bde acordo com a tabela\b",
        r"\bconforme a tabela\b",
        r"\bobserve a tabela\b",
        r"\bobserve o gráfico\b",
        r"\bcom base no gráfico\b",
        r"\bde acordo com o gráfico\b",
        r"\bconforme o gráfico\b",
        r"\bcom base no laudo\b",
        r"\bde acordo<span class="ml-2" /><span data-testid="markdown-streaming-circle" class="inline-block w-3 h-3 rounded-full bg-neutral-a12 align-middle mb-[0.1rem]" />
