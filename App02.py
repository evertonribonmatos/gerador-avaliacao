import os
import re
import json
import time
import tempfile
import unicodedata
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

# Caminho do Tesseract OCR no Windows
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

    raise ValueError(
        "GROQ_API_KEY não encontrada. Configure em st.secrets ou no arquivo .env."
    )

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

def normalizar_busca(texto: Any) -> str:
    if texto is None:
        return ""
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(ch for ch in texto if unicodedata.category(ch) != "Mn")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()

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
    for tentativa in range(3):
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

# ==========================================
# EXTRAÇÃO DE CONTEÚDO BASE
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
        raise ValueError("Envie um documento-base (.pdf ou .txt).")

    nome = arquivo_base.name.lower()
    if nome.endswith(".pdf"):
        return extrair_texto_pdf_upload(arquivo_base)
    elif nome.endswith(".txt"):
        return extrair_texto_txt_upload(arquivo_base)
    else:
        raise ValueError("O documento-base deve ser .pdf ou .txt.")

# ==========================================
# GERAÇÃO DAS QUESTÕES COM IA
# ==========================================
def montar_prompt_questoes(conteudo_base: str, dados_usuario: Dict[str, Any]) -> str:
    configuracao_questoes = []

    for q in dados_usuario["questoes"]:
        possui_imagem = "sim" if q.get("imagem_bytes") else "não"
        descricao_imagem = normalizar_texto(q.get("ocr_imagem", ""))

        bloco_imagem = ""
        if possui_imagem == "sim":
            bloco_imagem = f"""
INFORMAÇÕES DA IMAGEM ASSOCIADA À QUESTÃO {q['numero']:02d}:
- Existe uma imagem vinculada a esta questão.
- Texto extraído por OCR da imagem:
{descricao_imagem if descricao_imagem else "Nenhum texto legível foi extraído da imagem."}
"""

        configuracao_questoes.append(
            f"""
Questão {q['numero']:02d}:
- tipo={q['tipo']}
- peso={q['peso']}
- imagem_associada={possui_imagem}
{bloco_imagem}
"""
        )

    configuracao_texto = "\n".join(configuracao_questoes)

    return f"""
Atue como especialista em elaboração de avaliações técnicas para educação profissional industrial.

TAREFA:
Gerar EXATAMENTE {TOTAL_QUESTOES} questões com base EXCLUSIVAMENTE no conteúdo-base fornecido.

DADOS DA AVALIAÇÃO:
- Curso: {dados_usuario['curso']}
- Unidade Curricular: {dados_usuario['unidade_curricular']}
- Valor da avaliação: {dados_usuario['valor_avaliacao']}

CONFIGURAÇÃO DAS QUESTÕES:
{configuracao_texto}

REQUISITOS OBRIGATÓRIOS:
1. Escreva em português brasileiro formal, correto, técnico e rigoroso.
2. O texto deve seguir padrão de avaliação técnica de nível difícil.
3. Todas as questões devem ser contextualizadas.
4. O contexto deve favorecer pensamento crítico, analítico e tomada de decisão.
5. O comando da questão deve usar verbos compatíveis com a Taxonomia de Bloom, priorizando aplicar, analisar, avaliar e criar.
6. O contexto deve estar voltado, sempre que possível, à realidade industrial, produtiva, operacional, de manutenção, segurança, qualidade, diagnóstico, processos ou automação.
7. Não invente conteúdos fora do documento-base.
8. Para questões discursivas:
   - gere contexto + comando discursivo;
   - não inclua alternativas.
9. Para questões objetivas:
   - gere contexto + comando + 5 alternativas obrigatórias;
   - as alternativas devem estar preenchidas nos campos A, B, C, D e E;
   - não deixe nenhuma alternativa vazia;
   - apenas 1 alternativa correta;
   - forneça gabarito.
10. O texto pode ser longo o quanto for necessário para manter qualidade pedagógica e contextualização.
11. Retorne SOMENTE JSON válido.
12. O campo "contexto" deve conter o enunciado completo da questão.
13. Não use markdown.
14. Não use crases.
15. Preserve exatamente o tipo de cada questão solicitado.
16. Se houver imagem associada a uma questão, a elaboração dessa questão deve obrigatoriamente considerar a imagem associada e seu conteúdo textual extraído, vinculando o enunciado à análise, interpretação, identificação, diagnóstico ou aplicação relacionada à imagem.
17. Quando houver imagem associada à questão, o enunciado deve deixar claro que a resposta depende da observação da figura/imagem apresentada.
18. Não diga que você não consegue ver a imagem. Use apenas as informações disponibilizadas ao elaborar a questão.
19. Em cada questão objetiva, o campo "alternativas" deve obrigatoriamente conter exatamente as chaves A, B, C, D e E, todas com texto preenchido.

FORMATO DE SAÍDA:
{{
  "questoes": [
    {{
      "numero": 1,
      "tipo": "objetiva",
      "peso": "1,0",
      "contexto": "texto completo da questão",
      "alternativas": {{
        "A": "texto",
        "B": "texto",
        "C": "texto",
        "D": "texto",
        "E": "texto"
      }},
      "gabarito": "A",
      "bloom": "Analisar"
    }},
    {{
      "numero": 2,
      "tipo": "discursiva",
      "peso": "1,0",
      "contexto": "texto completo da questão",
      "alternativas": {{}},
      "gabarito": "",
      "bloom": "Avaliar"
    }}
  ]
}}

CONTEÚDO-BASE:
{conteudo_base}
"""

def validar_questoes_ia(dados: Dict[str, Any], dados_usuario: Dict[str, Any]) -> List[Dict[str, Any]]:
    if "questoes" not in dados or not isinstance(dados["questoes"], list):
        raise ValueError("A IA não retornou uma lista válida de questões.")

    questoes = dados["questoes"]
    if len(questoes) != TOTAL_QUESTOES:
        raise ValueError(f"A IA deve retornar exatamente {TOTAL_QUESTOES} questões.")

    config_por_numero = {q["numero"]: q for q in dados_usuario["questoes"]}

    for q in questoes:
        numero = q.get("numero")
        if numero not in config_por_numero:
            raise ValueError(f"Questão inválida retornada pela IA: {numero}")

        tipo_esperado = config_por_numero[numero]["tipo"]
        if q.get("tipo") != tipo_esperado:
            raise ValueError(
                f"Questão {numero}: tipo retornado '{q.get('tipo')}' difere do tipo esperado '{tipo_esperado}'."
            )

        if not normalizar_texto(q.get("contexto", "")):
            raise ValueError(f"Questão {numero} sem contexto/enunciado.")

        if tipo_esperado == "objetiva":
            alternativas = q.get("alternativas")
            if not isinstance(alternativas, dict):
                raise ValueError(f"Questão {numero}: alternativas inválidas.")

            alternativas_normalizadas = {}
            for letra in ["A", "B", "C", "D", "E"]:
                texto_alt = normalizar_texto(alternativas.get(letra, ""))
                if not texto_alt:
                    raise ValueError(f"Questão {numero}: alternativa {letra} ausente.")
                alternativas_normalizadas[letra] = texto_alt

            q["alternativas"] = alternativas_normalizadas

            if q.get("gabarito") not in ["A", "B", "C", "D", "E"]:
                raise ValueError(f"Questão {numero}: gabarito inválido.")
        else:
            q["alternativas"] = {}
            q["gabarito"] = ""

        q["peso"] = config_por_numero[numero]["peso"]
        q["imagem_bytes"] = config_por_numero[numero].get("imagem_bytes")
        q["imagem_nome"] = config_por_numero[numero].get("imagem_nome", "")
        q["ocr_imagem"] = config_por_numero[numero].get("ocr_imagem", "")

    questoes.sort(key=lambda x: x["numero"])
    return questoes

def tentar_geracao_com_modelo(modelo: str, prompt: str, dados_usuario: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], str, str]:
    completion = chamar_ia_com_retry(
        model=modelo,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=8000,
        response_format={"type": "json_object"}
    )

    resposta = completion.choices[0].message.content.strip()
    dados = extrair_json_de_texto(resposta)
    questoes = validar_questoes_ia(dados, dados_usuario)
    return questoes, resposta, modelo

def gerar_questoes_ia(conteudo_base: str, dados_usuario: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], str, str]:
    prompt = montar_prompt_questoes(conteudo_base, dados_usuario)

    try:
        return tentar_geracao_com_modelo(MODELO_PRINCIPAL, prompt, dados_usuario)
    except Exception as e:
        raise ValueError(f"Falha ao gerar questões com o modelo {MODELO_PRINCIPAL}: {e}")

# ==========================================
# WORD - APOIO
# ==========================================
def texto_celula(cell: _Cell) -> str:
    return "\n".join(p.text for p in cell.paragraphs).strip()

def limpar_celula(cell: _Cell) -> None:
    cell.text = ""

def escrever_linhas_e_imagem_na_celula(cell: _Cell, linhas: List[str], imagem_bytes: Optional[bytes] = None) -> None:
    cell.text = ""

    primeira = True
    for linha in linhas:
        if primeira:
            p = cell.paragraphs[0]
            p.text = linha
            primeira = False
        else:
            cell.add_paragraph(linha)

    if imagem_bytes:
        if not primeira:
            cell.add_paragraph("")
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
                txt = normalizar_busca(texto_celula(cell))
                if txt in mapa_labels:
                    valor = mapa_labels[txt]
                    if idx + 1 < len(row.cells):
                        row.cells[idx + 1].text = valor

def formatar_texto_questao(questao: Dict[str, Any]) -> List[str]:
    linhas = []

    bloom = normalizar_texto(questao.get("bloom", ""))
    contexto = normalizar_texto(questao.get("contexto", ""))

    if bloom:
        linhas.append(f"Nível cognitivo predominante: {bloom}.")
        linhas.append("")

    if contexto:
        linhas.extend(contexto.split("\n"))

    if questao["tipo"] == "objetiva":
        linhas.append("")
        for letra in ["A", "B", "C", "D", "E"]:
            alt = normalizar_texto(questao["alternativas"].get(letra, ""))
            linhas.append(f"({letra}) {alt}")

        if questao.get("gabarito"):
            linhas.append("")
            linhas.append(f"Gabarito: {questao['gabarito']}")

    return linhas

def permitir_altura_automatica_linha(row) -> None:
    tr = row._tr
    trPr = tr.get_or_add_trPr()

    for child in trPr.findall(qn("w:trHeight")):
        trPr.remove(child)

def obter_marcadores_questoes() -> Dict[str, int]:
    return {
        "primeira": 1,
        "segunda": 2,
        "terceira": 3,
        "quarta": 4,
        "quinta": 5,
        "sexta": 6,
        "setima": 7,
        "sétima": 7,
        "oitava": 8,
        "nona": 9,
        "decima": 10,
        "décima": 10,
    }

def preencher_tabela_questoes(document: Document, questoes: List[Dict[str, Any]]) -> None:
    mapa_questoes = {q["numero"]: q for q in questoes}
    marcadores = obter_marcadores_questoes()
    questoes_preenchidas = set()

    for table in document.tables:
        for i, row in enumerate(table.rows):
            permitir_altura_automatica_linha(row)

            for cell in row.cells:
                texto_original = texto_celula(cell)
                texto_norm = normalizar_busca(texto_original)

                numero_questao = None
                marcador_encontrado = None

                for marcador, numero in marcadores.items():
                    if texto_norm == marcador or marcador in texto_norm:
                        numero_questao = numero
                        marcador_encontrado = marcador
                        break

                if numero_questao is not None and numero_questao in mapa_questoes:
                    q = mapa_questoes[numero_questao]

                    # Preencher peso na linha acima
                    if i - 1 >= 0:
                        row_top = table.rows[i - 1]
                        permitir_altura_automatica_linha(row_top)

                        for idx_c, cell_top in enumerate(row_top.cells):
                            txt_top = normalizar_busca(texto_celula(cell_top))
                            if txt_top == "peso" and idx_c + 1 < len(row_top.cells):
                                row_top.cells[idx_c + 1].text = str(q["peso"])
                                break

                    # Apaga marcador e escreve questão
                    limpar_celula(cell)
                    linhas_questao = formatar_texto_questao(q)

                    escrever_linhas_e_imagem_na_celula(
                        cell,
                        linhas_questao,
                        imagem_bytes=q.get("imagem_bytes")
                    )

                    questoes_preenchidas.add(numero_questao)

    faltantes = [n for n in range(1, TOTAL_QUESTOES + 1) if n not in questoes_preenchidas]
    if faltantes:
        raise ValueError(
            f"Não foi possível localizar no Word os marcadores das questões: {faltantes}. "
            f"Verifique se o modelo contém exatamente: primeira, segunda, terceira, quarta, quinta, sexta, sétima, oitava, nona e decima."
        )

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
st.set_page_config(page_title="Gerador de Avaliação Técnica", layout="wide")
st.title("Gerador de Avaliação Técnica")
st.write("Preencha os dados abaixo para gerar a avaliação e baixar o arquivo Word preenchido.")

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
        "Forma de informar o conteúdo-base",
        ["Upload de arquivo (.pdf/.txt)", "Texto manual"]
    )

    arquivo_base = None
    conteudo_base_manual = ""

    if modo_conteudo == "Upload de arquivo (.pdf/.txt)":
        arquivo_base = st.file_uploader("Documento-base (.pdf ou .txt)", type=["pdf", "txt"])
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

        with st.spinner(f"Gerando questões com IA ({MODELO_PRINCIPAL})..."):
            questoes, resposta_bruta, modelo_usado = gerar_questoes_ia(conteudo_base, dados_usuario)

        with st.spinner("Preenchendo documento Word..."):
            docx_final_bytes = preencher_documento_word_em_memoria(
                modelo_docx_bytes=modelo_docx.getvalue(),
                dados_usuario=dados_usuario,
                questoes=questoes
            )

        st.success("Avaliação gerada com sucesso!")
        st.info(f"Modelo usado: {modelo_usado}")

        with st.expander("Resposta bruta da IA"):
            st.text(resposta_bruta)

        with st.expander("Visualizar JSON gerado pela IA"):
            st.json({"questoes": questoes})

        with st.expander("Debug das questões geradas"):
            for q in questoes:
                st.write(
                    f"Questão {q['numero']} | tipo: {q['tipo']} | peso: {q['peso']} | "
                    f"imagem: {'sim' if q.get('imagem_bytes') else 'não'}"
                )

                if q.get("ocr_imagem"):
                    st.write(f"OCR da imagem: {q['ocr_imagem'][:300]}")

                st.write("Contexto:")
                st.write(q["contexto"][:1000] + "..." if len(q["contexto"]) > 1000 else q["contexto"])

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