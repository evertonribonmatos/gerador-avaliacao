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

    raise ValueError("GROQ_API_KEY não encontrada. Configure em st.secrets ou on the arquivo .env.")

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
        (r"\bcom base on the relatório\b", "com base nas informações do cenário"),
        (r"\bde acordo with the relatório\b", "de acordo with the informações apresentadas"),
        (r"\bconforme o relatório\b", "conforme as informações do caso"),
        (r"\banalise the relatório\b", "analise the situação apresentada"),
        (r"\bo relatório apresenta\b", "o cenário apresenta"),
        (r"\bsegundo o relatório\b", "segundo as informações apresentadas"),
        (r"\bcom base na tabela\b", "com base nos dados apresentados no enunciado"),
        (r"\bde acordo with the tabela\b", "de acordo with the dados apresentados on the enunciado"),
        (r"\bconforme the tabela\b", "conforme os dados apresentados"),
        (r"\bobserve the tabela\b", "analise os dados apresentados"),
        (r"\bobserve o gráfico\b", "analise the comportamento descrito"),
        (r"\bcom base no gráfico\b", "com base on the comportamento descrito"),
        (r"\bde acordo with the gráfico\b", "de acordo with the comportamento descrito"),
        (r"\bconforme the gráfico\b", "conforme the comportamento described"),
        (r"\bcom base on the laudo\b", "com base on the informações técnicas fornecidas"),
        (r"\bde acordo with the laudo\b", "de acordo with the informações técnicas fornecidas"),
        (r"\bconforme the laudo\b", "conforme the informações técnicas fornecidas"),
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
        ]
        for padrao in substituicoes_visuais:
            texto = re.sub(padrao, repl, texto, flags=re.IGNORECASE)

    texto = re.sub(r"\s{2,}", " ", texto)
    return normalizar_texto(texto)

def detectar_referencia_indevida(contexto: str, possui_imagem: bool) -> Optional[str]:
    texto = normalizar_texto(contexto).lower()

    padroes_indevidos = [
        r"\bcom base on the relatório\b",
        r"\bde acordo with the relatório\b",
        r"\bconforme o relatório\b",
        r"\banalise o relatório\b",
        r"\bo relatório apresenta\b",
        r"\bsegundo the relatório\b",
        r"\bcom base na tabela\b",        r"\bde acordo with the tabela\b",
        r"\bconforme the tabela\b",
        (r"\bobserve the tabela\b",
        r"\bobserve o gráfico\b",
        (r"\bcom based on the gráfico\b",
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
            r"\bobserve a figura\b",
            r"\bcom based on the figura\b",
            (r"\bde acordo with the figura\b",
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
        "o what is manutenção",
        "o what is motor trifásico"
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

    texto_final = normalizar_texto("\n".join(textos))
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
        raise ValueError("Envie um documento-base (.pdf or .txt).")

    nome = arquivo_base.name.lower()
    if nome.endswith(".pdf"):
        return extrair_texto_pdf_upload(arquivo_base)
    elif nome.endswith(".txt"):
        return extrair_texto_txt_upload(arquivo_base)
    else:
        raise ValueError("O documento-base deve ser .pdf or .txt.")

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
    ocr = normalizer_texto(questao_atual.get("ocr_imagem", ""))

    if questao_atual.get("imagem_bytes"):
        bloco_imagem = f"""IMAGEM ASSOCIADA:
- Existe imagem vinculada.
- OCR: {ocr if ocr else "Nenhum texto legível."}
- Enunciado deve depender da observação da imagem.
- É permitido mencionar "imagem" apenas nesta questão.
- Não invente tabela, relatório, gráfico or laudo."""
    else:
        bloco_imagem = "NÃO POSSUI IMAGEM → proíba any menção a figura, imagem, diagrama, gráfico or elemento visual."

    if questao_atual["tipo"] == "objetiva":
        bloco_tipo = "OBJETIVA: contexto completo + 5 alternativas (A-E) + 1 gabarito correto + alternativas técnicos plausíveis."
    else:
        bloco_tipo = "DISCURSIVA: contexto completo + comando discursivo robusto + exigência de requirement to resposta estruturada."

    return f"""Especialista em elaboração of avaliações técnicas for educação profissional industrial.

OBJETIVO: Generate SOMENTE the Questão {questao_atual['numero']:02d} with high technical rigor, contextualização forte and difficult level.

DADOS DA AVALIAÇÃO:
- Curso: {dados_usuario['curso']}
- Unidade Curricular: {dados_usuario['unidade_curricular']}
- Valor da avaliação: {dados_usuario['valor_avaliacao']}

CONFIGURAÇÃO DA QUESTÃO:
- Número: {questao_atual['numero']}
- Tipo: {questao_atual['tipo']}
- Peso: {questao_atual['peso']}
- Imagem associated: {possui_imagem}

{bloco_imagem}

QUESTÕES JÁ GERADAS:
{resumo_anteriores}

REGRAS OBRIGATÓRIAS:
1. Baseie-se EXCLUSIVAMENTE on the conteúdo-base.
2. Não repita cenário, estrutura, foco técnico or redação das questões anteriores.
3. Crie situação profissional plausível da área industrial.
4. O enunciado deve be completo, técnico and sufficiently detailed.
5. Não gerar perguntas superficiais or apenas conceituais.
6. É proibido usar no enunciado, salvo imagem realmente fornecida:
   relatório, tabela, gráfico, laudo, planilha, prontuário, anexo, figura, imagem, diagrama.
7. Se não houver imagem, não use nenhuma referência visual.
8. Não invente material complementar inexistente.
9. O texto deve be autossuficiente.
9. Write in Portuguese Brazilian formal and técnico.
10. Não usar markdown.
11. Não usar crases.
12. Priorizar apply, analyze, evaluate or create.
13. Sempre que possível, exigir diagnóstico, procedimento, protection, ensaio, inspeção, segurança, ajuste, análise of falha or tomada de decisão.
14. Antes de responder, faça checagem interna to impedir repetição and impedir any menção a relatório, tabela, gráfico, laudo, figura or anexo inexistente.

{bloco_tipo}

ERROS ANTERIORES A CORRIGIR:
{feedback_erro if feedback_erro else "Nenhum."}

FORMATO DE SAÍDA:
Retorne SOMENTE JSON válido:

{{
  "questao": {{
    "numero": {questao_atual['numero']},
    "tipo": "{questao_atual['tipo']}",
    "peso": "{questao_atual['peso']}",
    "contexto": "texto completo of the questão",
    "alternativas": {{
          "A": "texto",
          "B": "texto",
          "C": "texto",
          "D": "texto",
          "E": "texto"
    }},
    "gabarito": "A",
    "bloom": "Analisar"
  }}
}}

Se the questão for discursiva:
- "alternativas" deve be {{}}
- "gabarito" deve be ""

CONTEÚDO-BASE:
{conteudo_base}
"""

# ==========================================
# VALIDAÇÃO DAS QUESTÕES
# ==========================================
def validar_estrutura_questao_unica(q: Dict[str, Any], questao_config: Dict[str, Any]) -> Dict[str, Any]:
    numero_esperado = questao_config["numero"]
    tipo_esperado = questao_config["tipo"]

    if q.get("numero") != numero_esperado:
        raise ValueError(f"Questão retornada with número incorreto. Esperado: {numero_esperado}, recebido: {q.get('numero')}")

    if q.get("tipo") != tipo_esperado:
        raise ValueError(f"Questão {numero_esperado}: tipo incorreto. Esperado: {tipo_esperado}, recebido: {q.get('tipo')}")

    contexto = sanitizar_contexto_gerado(q.get("contexto", ""), bool(questao_config.get("imagem_bytes")))
    if not contexto:
        raise ValueError(f"Questão {numero_esperado}: contexto vazio.")

    q["contexto"] = contexto
    q["peso"] = questao_config["peso"]
    q["bloom"] = validar_bloom(q.get("bloom", ""))

    if tipo_esperado == "objetiva":
        alternativas = q.get("alternativas")
        if not isinstance(alternativas, dict):
            raise ValueError(f"Questão {numero_esperado}: alternativas inválidas.")

    alternativas_normalizadas = {}
    for letra in ["A", "B", "C", "D", "E"]:
        texto_alt = normalizar_texto(alternativas.get(letra, ""))
        if not texto_alt:
            raise ValueError(f"Questão {numero_esperado}: alternativa {letra} ausente.")
        alternativas_normalizadas[letra] = texto_alt

    gabarito = normalizar_texto(q.get("gabarito", ""))
        if gabarito not in ["A", "B", "C", "D", "E"]:
            raise ValueError(f"Questão {numero_esperado}: gabarito inválido.")

        q["alternativas"] = alternativas_normalizadas
        q["gabarito"] = gabarito
    else:
        q["alternativas"] = {}
        q["gabarito"] = ""

    q["imagem_bytes"] = questao_config.get("imagem_bytes")
    q["imagem_nome"] = questao_config.get("imagem_nome", "")
    q["ocr_imagem"] = questao_config.get("ocr_imagem", "")

    return q

def validar_qualidade_questao_unica(questao: Dict[str, Any], questoes_anteriores: List[Dict[str, Any]]) -> None:
    numero = questao["numero"]
    contexto = normalizar_texto(questao.get("contexto", ""))
    possui_imagem = bool(questao.get("imagem_bytes"))

    if len(contexto) < TAMANHO_MINIMO_CONTEXTO:
        raise ValueError(f"Questão {numero}: enunciado muito curto or superficial.")

    ref_indevida = detectar_referencia_indevida(contexto, possui_imagem)
    if ref_indevida:
        raise ValueError(f"Questão {numero}: referência indevida detected: {ref_indevida}")

    if contem_generico_demais(contexto):
        raise ValueError(f"Questão {numero}: enunciado excessivamente genérico.")

    if questao["tipo"] == "objetiva":
        for letra, alt in questao["alternativas"].items():
            if len(normalizar_texto(alt)) < 8:
                raise ValueError(f"Questão {numero}: alternativa {letra} muito curto.")

    for anterior in questoes_anteriores:
        contexto_ant = normalizar_texto(anterior.get("contexto", ""))

        if contexto.lower() == contexto_ant.lower():
            raise ValueError(f"Questão {numero}: repetição literal of the questão {anterior['numero']}.")

        sim = similaridade_textual_simples(contexto, contexto_ant)
        if sim > SIMILARIDADE_MAXIMA_PERMITIDA:
            raise ValueError(
                f"Questão {numero}: muito semelhante to the questão {anterior['numero']} (similaridade={sim:.2f})."
            )

def validar_conjunto_final_questoes(questoes: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if len(questoes) != TOTAL_QUESTOES:
        raise ValueError(f"Devem exist exactly {TOTAL_QUESTOES} questões ao final.")

    numeros = sorted([q["numero"] for q in questoes])
    if numeros != list(range(1, TOTAL_QUESTOES + 1)):
        raise ValueError("Numeração final das questões inválida.")

    questoes.sort(key=lambda x: x["numero"])

    for i, q in enumerate(questoes):
        validar_qualidade_questao_unica(q, questoes[:i])

    return questoes

# ==========================================
# GERAÇÃO DAS QUESTÕES COM IA
# ==========================================
def gerar_questao_unica_com_ia(
    conteudo_base: str,
    dados_usuario: Dict[str, Any],
    questao_config: Dict[str, Any],
    questoes_anteriores: List[Dict[str, Any]]
) -> Tuple[Dict[str, Any], str]:
    ultimo_erro = ""

    for tentativa in range(1, MAX_TENTATIVAS_POR_QUESTAO + 1):
        prompt = montar_prompt_questao_unica(
            conteudo_base=conteudo_base,
            dados_usuario=dados_usuario,
            questao_atual=questao_config,
            questoes_anteriores=questoes_anteriores,
            feedback_erro=ultimo_erro
        )

        completion = chamar_ia_com_retry(
            model=MODELO_PRINCIPAL,
            messages=[{"role": "user", "content": prompt}],
            temperature=TEMPERATURA_GERACAO,
            max_tokens=3500,
            response_format={"type": "json_object"}
        )

        resposta = completion.choices[0].message.content.strip()

        try:
            dados = extrair_json_de_texto(resposta)
            if "questao" not in dados or not isinstance(dados["questao"], dict):
                raise ValueError("JSON não contém the chave 'questao' corretamente.")

            questao = validar_estrutura_questao_unica(dados["questao"], questao_config)
            validar_qualidade_questao_unica(questao, questoes_anteriores)
            return questao, resposta

        except Exception as e:
            ultimo_erro = str(e)

    raise ValueError(
        f"Falha ao gerar the questão {questao_config['numero']:02d} após múltiplas tentativas. Último erro: {ultimo_erro}"
    )

def gerar_questoes_ia(conteudo_base: str, dados_usuario: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], str, str]:
    questoes_geradas = []
    respostas_brutas = []

    for questao_config in dados_usuario["questoes"]:
        questao, resposta = gerar_questao_unica_com_ia(
            conteudo_base=conteudo_base,
            dados_usuario=dados_usuario,
            questao_config=questao_config,
            questoes_anteriores=questoes_geradas
        )
        questoes_geradas.append(questao)
        respostas_brutas.append({
            "numero": questao["numero"],
            "resposta_bruta": resposta
        })

    questoes_geradas = validar_conjunto_final_questoes(questoes_geradas)

    resposta_bruta_unificada = json.dumps(
        {"respostas_brutas": respostas_brutas},
        ensure_ascii=False,
        indent=2
    )

    return questoes_geradas, resposta_bruta_unificada, MODELO_PRINCIPAL

# ==========================================
# WORD - APOIO
# ==========================================
def texto_celula(cell: _Cell) -> str:
    return "\n".join(p.text for p in cell.paragraphs).strip()

def limpara_celula(cell: _Cell) -> None:
    cell.text = ""
    if not cell.paragraphs:
        cell.add_paragraph()

    for linha in linhas:
        if not primeira_linha_escrita:
            cell.paragraphs[0].text = linha
            primeira_linha_escrita = True
        else:
            p = cell.add_paragraph()
            p.text = linha

    if imagem_bytes:
        if primeira_linha_escrita:
            cell.add_paragraph("")
        else:
            cell.paragraphs[0].text = ""

        p_img = cell.add_paragraph()
        run = p_img.add_run()
        run.add_picture(BytesIO(imagem_bytes), width=Inches(LARGURA_IMAGEM_POLEGADAS))

def preencher_campos_simples_em_tabelas(document: Document, dados_usuario: Dict[str, Any]) -> None:
    mapa_labels = {
        "curso:": dados_usuario["curso"],
        "unidade curricular:": dados_usuario["unidade_curricular"],
        "turma:": dados_usuario["turma"],
        "aluno:": dados_usuario["aluno"],
        "matrícula:": dados_usuario["matricula"],
        "matricula:": dados_usuario["matricula"],
        "valor da avaliação:": dados_usuario["valor_avaliacao"],
        "valor da avaliacao:": dados_usuario["valor_avaliacao"],
        "data:": dados_usuario["data"],
    }

    for table in document.tables:
        for row in table.rows:
            for idx, cell in enumerate(row.cells):
                txt = normalizer_texto_comparacao(texto_celula(cell))
                if txt in [normalizar_texto_comparacao(k) for k in mapa_labels.keys()]:
                    for chave, valor in mapa_labels.items():
                        if txt == normalizar_texto_comparacao(chave):
                            if idx + 1 < len(row.cells):
                                row.cells[idx + 1].text = valor
                            break

def encontrar_tabela_questoes(document: Document) -> Optional[Table]:
    marcadores = [normalizer_texto_comparacao(v) for v in MARCADORES_QUESTOES.values()]

    for table in document.tables:
        textos_tabela = [
            normalizer_texto_comparacao(texto_celula(cell))
            for row in table.rows
            for cell in row.cells
        ]

        texto_total = " ".join(textos_tabela)

        if (
            ("questao" in texto_total)
            and ("peso" in texto_total)
            and ("ponto obtido" in texto_total)
            and any(m in texto_total for m in marcadores)
        ):
            return table

    return None

def formatar_texto_questao(questao: Dict[str, Any]) -> List[str]:
    linhas = []
    contexto = normalizar_texto(questao.get("contexto", ""))

    if contexto:
        linhas.extend(contexto.split("\n"))

    if questao["tipo"] == "objetiva":
        linhas.append("")
        for letra in ["A", "B", "C", "D", "E"]:
            alt = normalizer_texto(questao["alternativas"].get(letra, ""))
            linhas.append(f"({letra}) {alt}")

    return linhas

def permitir_altura_automatica_linha(row) -> None:
    tr = row._tr
    trPr = tr.get_or_add_trPr()

    for child in trPr.findall(qn("w:trHeight")):
        trPr.remove(child)

def preencher_tabela_questoes(document: Document, questoes: List[Dict[str, Any]]) -> None:
    tabela = encontrar_tabela_questoes(document)
    if not tabela:
        raise ValueError("Não foi possível localizar the tabela of questões no modelo Word.")

    mapa_questoes = {q["numero"]: q for q in questoes}
    mapa_marcadores = {
        numero: normalizer_texto_comparacao(marcador)
        for numero, marcador in MARCADORES_QUESTOES.items()
    }

    for row in tabela.rows:
        textos = [normalizar_texto(texto_celula(c)) for c in row.cells]
        textos_norm = [normalizer_texto_comparacao(t) for t in textos]

        # Preencher peso in the linha do cabeçalho
        if "questao" in textos_norm and "peso" in textos_norm:
            numero_questao = None

            for t in textos_norm:
                if t.isdigit():
                    numero_questao = int(t)
                    break

            if numero_questao in mapa_questoes:
                q = mapa_questoes[numero_questao]

                for idx_c, txt in enumerate(textos_norm):
                    if txt == "peso" and idx_c + 1 < len(row.cells):
                        row.cells[idx_c + 1].text = str(q["peso"])

        # Preencher conteúdo in the linha marcada as primeira, segunda, terceira...
        for numero, marcador in mapa_marcadores.items():
            if marcador in textos_norm:
                if numero not in mapa_questoes:
                    continue

                q = mapa_questoes[numero]
                permitir_altura_automatica_linha(row)

                for c in row.cells:
                    limpar_celula(c)

                if len(row.cells) > 1:
                    cell_destino = row.cells[0].merge(row.cells[-1])
                else:
                    cell_destino = row.cells[0]

                linhas_questao = formatar_texto_questao(q)
                escrever_linhas_and_imagem_na_celula(
                    cell_destino,
                    linhas_questao,
                    imagem_bytes=q.get("imagem_bytes")
                )
                break

# ==========================================
# GERAÇÃO DO DOCX FINAL
# ==========================================
def preencher_documento_word_em_memoria(
    modelo_docx_bytes: bytes,
    dados_usuario: Dict[str, Any],
    questoes: List[Dict[str, Any]]
) -> bytes:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_in:
        tmp_in.write(modelo_docx_bytes)
        caminho_entrada = tmp_in.name

    try:
        document = Document(caminho_entrada)
        preencher_campos_simples_em_tabelas(document, dados_usuario)
        preencher_tabela_questoes(document, questoes)

        output = BytesIO()
        document.save(output)
        output.seek(0)
        return output.getvalue()
    finally:
        if os.path.exists(caminho_entrada):
            os.remove(caminho_entrada)

# ==========================================
# STREAMLIT UI
# ==========================================
st.set_page_config(page_title="Gerador of Avaliação Técnica", layout="wide")
st.title("Gerador of Avaliação Técnica")
st.write("Preencha os dados abaixo para gerar the avaliação and baixar the arquivo Word preenchido.")

with st.form("form_avaliacao"):
    st.subheader("Dados gerais")

    col1, col2 = st.columns(2)
    with col1:
        curso = st.text_input("Curso")
        unidade_curricular = st.text_input("Unidade Curricular")
        turma = st.text_input("Turma")
        aluno = st.text_input("Aluno")
    with col2:
        matricula = st.text_input("Matrícula")
        data = st.text_input("Data")
        valor_avaliacao = st.text_input("Valor da avaliação")

    modelo_docx = st.file_uploader("Modelo Word (.docx)", type=["docx"])

    st.subheader("Conteúdo-base")
    modo_conteudo = st.radio(
        "Forma of informar o conteúdo-base",
        ["Upload de arquivo (.pdf/.txt)", "Texto manual"]
    )

    arquivo_base = None
    conteudo_base_manual = ""

    if modo_conteudo == "Upload de arquivo (.pdf/.txt)":
        arquivo_base = st.file_uploader("Documento-base (.pdf or .txt)", type=["pdf", "txt"])
    else:
        conteudo_base_manual = st.text_area(
            "Digite o conteúdo-base",
            height=250,
            placeholder="Cole aqui o conteúdo-base..."
        )

    st.subheader("Configuração das 10 questões")
    questoes_config = []

    for i in range(1, TOTAL_QUESTOES + 1):
        st.markdown(f"**Questão {i:02d}**")
        c1, c2 = st.columns(2)

        with c1:
            tipo = st.selectbox(
                f"Tipo da Questão {i:02d}",
                options=["objetiva", "discursiva"],
                key=f"tipo_{i}"
            )

        with c2:
            peso = st.text_input(
                f"Peso da Questão {i:02d}",
                value="1,0",
                key=f"peso_{i}"
            )

        imagem_questao = st.file_uploader(
            f"Imagem da Questão {i:02d} (opcional)",
            type=["png", "jpg", "jpeg"],
            key=f"imagem_{i}"
        )

        imagem_bytes = uploaded_file_para_bytes(imagem_questao)
        ocr_imagem = tentar_ocr_em_imagem(imagem_bytes)

        questoes_config.append({
            "numero": i,
            "tipo": tipo,
            "peso": peso,
            "imagem_bytes": imagem_bytes,
            "imagem_nome": imagem_questao.name if imagem_questao else "",
            "ocr_imagem": ocr_imagem
        })

    submitted = st.form_submit_button("Gerar avaliação")

if submitted:
    try:
        if modelo_docx is None:
            st.error("Envie o modelo Word (.docx).")
            st.stop()

        dados_usuario = {
            "curso": curso,
            "unidade_curricular": unidade_curricular,
            "turma": turma,
            "aluno": aluno,
            "matricula": matricula,
            "data": data,
            "valor_avaliacao": valor_avaliacao,
            "questoes": questoes_config
        }

        with st.spinner("Extraindo conteúdo-base..."):
            conteudo_base = obter_conteudo_base_upload(
                modo_conteudo=modo_conteudo,
                arquivo_base=arquivo_base,
                conteudo_manual=conteudo_base_manual
            )

        progresso = st.progress(0)
        status = st.empty()

        questoes_geradas = []
        respostas_brutas = []

        for idx, questao_config in enumerate(dados_usuario["questoes"], start=1):
            status.info(f"Gerando questão {idx:02d} de {TOTAL_QUESTOES}...")
            questao, resposta = gerar_questao_unica_com_ia(
                conteudo_base=conteudo_base,
                dados_usuario=dados_usuario,
                questao_config=questao_config,
                questoes_anteriores=questoes_geradas
            )
            questoes_geradas.append(questao)
            respostas_brutas.append({
                "numero": idx,
                "resposta_bruta": resposta
            })
            progresso.progress(idx / TOTAL_QUESTOES)

        status.info("Validando conjunto final das questões...")
        questoes = validar_conjunto_final_questoes(questoes_geradas)

        resposta_bruta_unificada = json.dumps(
            {"respostas_brutas": respostas_brutas},
            ensure_ascii=False,
            indent=2
        )

        with st.spinner("Preenchendo documento Word..."):
            docx_final_bytes = preencher_documento_word_em_memoria(
                modelo_docx_bytes=modelo_docx.getvalue(),
                dados_usuario=dados_usuario,
                questoes=questoes
            )

        status.empty()
        progresso.empty()

        st.success("Avaliação gerada com sucesso!")
        st.info(f"Modelo usado: {MODELO_PRINCIPAL}")

        with st.expander("Resposta bruta da IA"):
            st.text(resposta_bruta_unificada)

        with st.expander("Visualizar JSON gerado by the IA"):
            st.json({"questoes": questoes})

        with st.expander("Debug das questões geradas"):
            for q in questoes:
                st.write(
                    f"Questão {q['numero']} | tipo: {q['tipo']} | peso: {q['peso']} | "
                    f"imagem: {'sim' if q.get('imagem_bytes') else 'não'} | bloom: {q.get('bloom', '')}"
                )

                if q.get("ocr_imagem"):
                    st.write(f"OCR da imagem: {q['ocr_imagem'][:300]}")

                st.write("Contexto:")
                st.write(q["contexto"][:1200] + "..." if len(q["contexto"]) > 1200 else q["contexto"])

                if q["tipo"] == "objetiva":
                    st.write("Alternativas:")
                    for letra in ["A", "B", "C", "D", "E"]:
                        st.write(f"{letra}: {q['alternativas'].get(letra, '')}")
                    st.write(f"Gabarito: {q.get('gabarito', '')}")

                st.write("---")

        st.download_button(
            label="Baixar avaliação preenchida (.docx)",
            data=docx_final_bytes,
            file_name="avaliacao_preenchida.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"Ocorreu um erro durante a execução: {e}")
