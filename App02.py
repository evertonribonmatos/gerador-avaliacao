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
    8: "oitada",
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
        (r"\bde acordo com o gráfico\b", "de acordo com o comportamento described"),
        (r"\bconforme o gráfico\b", "conforme o comportamento descrito"),
        (r"\bcom base no laudo\b", "com base nas informações técnicos fornecidas"),
        (r"\bde acordo com o laudo\b", "de acordo com as informações técnicos fornecidas"),
        (r"\bconforme o laudo\b", "conforme as informações técnicos fornecidas"),
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
            (r"\bde acordo com a imagem\b", "de acordo com the situação apresentada"),
            (r"\bconforme a imagem\b", "conforme a situação presented"),
            (r"\bobserve o diagrama\b", "considere the situação apresentada"),
            (r"\bcom base no diagrama\b", "com base na situação apresentada"),
            (r"\bde acordo com o diagrama\b", "de acordo com the situação presented"),
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
        r"\bcom based on the tabela\b",
        r"\bde acordo com the tabela\b",
        r"\bconforme the tabela\b",
        r"\bobserve the tabela\b",
        r"\bobserve the gráfico\b",
        r"\bcom based on the gráfico\b",
        r"\bde acordo with the gráfico\b",
        r"\bconforme the gráfico\b",
        r"\bcom base on the laudo\b",
        r"\bde acordo with the laudo\b",
        r"\bconforme the laudo\b",
        r"\bem anexo\b",
        r"\bno anexo\b",
    ]

    for padrao in padroes_indevidos:
        if re.search(padrao, texto):
            return padrao

    if not possui_imagem:
        padroes_visuais = [
            r"\bobserve the figura\b",
            r"\bcom based on the figura\b",
        ]
        for padrao in padroes_visuais:
            if re.search(padrao, texto):
                return padrao

    return None

def contem_generico_demais(contexto: str) -> bool:
    texto = normalizar_texto(contexto).lower()
    padroes_genericos = [
        "explique o que é",
        "defina",
        "conceitue",
        "cite",
        "liste",
        "o que é manutenção",
        "o que é motor trifásico"
    ]
    return any(p in texto for p in padroes_genericos)

def validar_bloom(bloom: str) -> str:
    bloom = normalizar_texto(bloom)
    if not bloom:
        return "Analisar"
    return bloom

# ==========================================
# EXTRAÇÃO de CONTEÚDO BASE
# ==========================================
def extrair_texto_pdf_upload(uploaded_file) -> str:
    textos = []
    uploaded_file.seek(0)

    with pdfplumber.open(uploaded_file) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            textos.append(texto)

    texto_final = normalizado_texto("\n".join(textos))
    if not texto_final:
        raise ValueError("Não foi possível extrair texto do PDF enviado.")
    return texto_final

def extrair_texto_txt_upload(uploaded_file) -> str:
    uploaded_file.seek(0)
    conteudo_bytes = uploaded_file.read()
    try:
        texto = conteudo_bytes.decode("utf-8")
    except UnicodeDecodeError:
        texto = conteudo_bytes.decode("latin-1")

    texto = normalizar_texto(texto)
    if not texto:
        raise ValueError("O arquivo TXT enviado está vazio.")
    return texto

def obter_conteudo_base_upload(modo_conteudo: str, arquivo_base, conteudo_manual: str) -> str:
    if modo_conteudo == "Texto manual":
        conteudo = normalizar_texto(conteudo_manual)
        if not conteudo:
            raise ValueError("O conteúdo-base manual não pode ficar vazio.")
        return conteudo

    if arquivo_base is None:
        raise ValueError("Envie um documento-base (.pdf ou .txt).")

    nome = arquivo_base.name.lower()
    if nome.endswith(".pdf"):
        return extrair_texto_pdf_upload(arquivo_base)
    elif nome.endswith(".txt"):
        return extrair_texto_txt_upload(arquivo_base)
    else:
        raise ValueError("O documento-base deve ser .pdf ou .txt.")

# ==========================================
# PROMPTS
# ==========================================
def montar_resumo_questoes_anteriores(questoes_anteriores: List[Dict[str, Any]]) -> str:
    if not questoes_anteriores:
        return "Nenhuma questão anterior."

    blocos = []
    for q in questoes_anteriores:
        contexto = normalizar_texto(q.get("contexto", ""))
        contexto_curto = contexto[:500]
        blocos.append(
            f"Questão {q['numero']:02d} | tipo={q['tipo']} | bloom={q.get('bloom', '')} | resumo={contexto_curto}"
        )
    return "\n".join(blocos)


def montar_prompt_questao_unica(
    conteudo_base: str,
    dados_usuario: Dict[str, Any],
    questao_atual: Dict[str, Any],
    questoes_anteriores: List[Dict[str, Any]],
    feedback_erro: str = ""
) -> str:
    resumo_anteriores = montar_resumo_questoes_anteriores(questoes_anteriores)
    possui_imagem = "sim" if questao_atual.get("imagem_bytes") else "não"
    ocr = normalizado_texto(questao_atual.get("ocr_imagem", ""))

    if questao_atual.get("imagem_bytes"):
        bloco_imagem = f"""IMAGEM ASSOCIADA:
- Existe imagem vinculada.
- OCR: {ocr if ocr else "Nenhum texto legível."}
- Enunciado deve depender da observação da imagem.
- É permitido mencionar "imagem" apenas nesta questão.
- Não invente tabela, relatório, gráfico ou laudo."""
    else:
        bloco_imagem = "NÃO POSSUI IMAGEM → proíba qualquer menção a figura, imagem, diagrama, gráfico ou elemento visual."

    if questao_atual["tipo"] == "objetiva":
        bloco_tipo = "OBJETIVA: contexto completo + 5 alternativas (A-E) + 1 gabarito correto + alternativas técnicas plausíveis."
    else:
        bloco_tipo = "DISCURSIVA: contexto completo + comando discursivo robusto + exigência de resposta estruturada."

    return f"""Especialista em elaboração de avaliações técnicas para educação profissional industrial.

OBJETIVO: Gerar SOMENTE a Questão {questao_atual['numero']:02d} com alto rigor técnico e nível difícil.

DADOS:
- Curso: {dados_usuario['curso']}
- Unidade Curricular: {dados_usuario['unidade_curricular']}
- Valor: {dados_usuario['valor_avaliacao']}

CONFIGURAÇÃO:
- Número: {questao_atual['numero']}
- Tipo: {questao_atual['tipo']}
- Peso: {questao_atual['peso']}
- Imagem: {possui_imagem}

{bloco_imagem}

QUESTÕES JÁ GERADAS:
{resumo_anteriores}

REGRAS OBRIGATÓRIAS:
1. Baseie-se EXCLUSIVAMENTE no conteúdo-base.
2. Não repita cenário, estrutura, foco técnico or redação das questões anteriores.
3. Crie situação profissional plausível da área industrial.
4. Enunciado deve ser completo, técnico e suficientemente detalhado.
5. Não gerar perguntas superficiais or apenas conceituais.
6. Proíba uso de relatório, tabela, gráfico, laudo, planilha, prontuário, anexo (salvo imagem real).
7. Se não houver imagem, não use nenhuma referência visual.
8. Texto autossuficiente.
9. Português brasileiro formal e técnico.
10. Não usar markdown, crases.
11. Priorizar aplicar, analisar, avaliar or criar.
12. Exigir diagnóstico, procedimento, proteção, ensaio, inspeção, segurança, ajuste, análise de falha or tomada de decisão.
13. Checar internamente para impedir repetição and menção a materiais inexistentes.

{bloco_tipo}

ERROS ANTERIORES: {feedback_erro if feedback_erro else "Nenhum."}

FORMATO DE SAÍDA (JSON somente):
{{
  "questao": {{
    "numero": {questao_atual['numero']},
    "tipo": "{questao_atual['tipo']}",
    "peso": "{questao_atual['peso']}",
    "contexto": "texto completo",
    "alternativas": {{"A":"", "B":"", "C":"", "D":"", "E":""}},
    "gabarito": "A",
    "bloom": "Analisar"
  }}
}}

If discursiva: "alternativas" = {{}}, "gabarito" = ""

CONTEÚDO-BASE:
{conteudo_base}
"""
