# app.py
# Leitor de Manifestos Jadlog — OCR Final (v5.1)
# Melhorias: extração robusta de valor + detecção de destinatário (prioriza bloco do destinatário)

import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
import pdfplumber
import pytesseract
from pdf2image import convert_from_bytes
from PIL import ImageOps, ImageFilter
import streamlit as st

# ========== Config Streamlit ==========
st.set_page_config(page_title="Leitor de Manifestos Jadlog", page_icon="🚛", layout="centered")
st.markdown("""
<style>
.stDownloadButton > button {
    background-color: #d62828;
    color: White;
    border: none;
    border-radius: 8px;
    padding: .6rem 1.2rem;
    font-weight: 700;
}
.stDownloadButton > button:hover {
    background-color: #6c757d;
}
</style>
""", unsafe_allow_html=True)

# ========== Mapas ==========
ROTA_CO_MAP = {
    "S10": "CO GUAPIMIRIM",
    "S12": "CO RIO DE JANEIRO 13",
    "S16": "CO QUEIMADOS",
    "S18": "CO SAO JOAO DE MERITI 01",
    "S20": "ML BELFORD ROXO 01",
    "S21": "CO JUIZ DE FORA",
    "S22": "DL DUQUE DE CAXIAS",
    "S24": "CO ITABORAI",
    "S25": "CO VOLTA REDONDA",
    "S27": "CO RIO DE JANEIRO 05",
    "S28": "CO RIO DE JANEIRO 04",
    "S30": "CO RIO DE JANEIRO 01",
    "S31": "CO JUIZ DE FORA",
    "S32": "CO RIO DE JANEIRO 03",
    "S37": "CO TRES RIOS",
    "S38": "CO RIO DE JANEIRO 06",
    "S41": "CO RIO DE JANEIRO 08",
    "S43": "CO ANGRA DOS REIS",
    "S45": "CO PARATY",
    "S48": "CO PETROPOLIS",
    "S49": "CO RIO DE JANEIRO 07",
    "S54": "CO NOVA FRIBURGO",
    "S56": "CO TERESOPOLIS",
    "S58": "CO CAMPOS D. GOYTCAZES",
    "S59": "CO SANT. ANT DE PADUA",
    "S60": "CO DUQUE DE CAXAIS",
    "S61": "CO CABO FRIO",
    "S63": "CO ARARUAMA",
    "S64": "CO SAO GONCALO",
    "S65": "CO RIO DAS OSTRAS",
    "S67": "CO NITEROI",
    "S68": "CO MARICÁ",
    "S70": "FL RIO DE JANEIRO",
}
UF_VALIDAS = {
    "AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT",
    "PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO"
}

# ========== Utilitárias ==========
def normalize(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in s if not unicodedata.combining(c))

def clean_int(s: str) -> str:
    return re.sub(r"[^\d]", "", s or "")

def pil_preprocess(img):
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    g = g.filter(ImageFilter.SHARPEN)
    g = g.point(lambda x: 255 if x > 180 else 0)
    return g

def ocr_image(img, psm=6):
    cfg = f"--psm {psm} --oem 3"
    return pytesseract.image_to_string(img, lang="por+eng", config=cfg).upper()

# ========== Ler PDF ==========
def read_pdf_text(file_bytes: bytes):
    pages_txt = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for p in pdf.pages:
            pages_txt.append(p.extract_text() or "")
    return pages_txt

# ========== Data/hora ==========
def extract_data_hora_from_head(first_page_text: str):
    cab = normalize(first_page_text or "")
    m = re.search(
        r"\b(\d{1,2})\s+(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\s+(\d{4})\s+(\d{2}:\d{2}:\d{2})",
        cab, re.I
    )
    if not m:
        return "", ""
    meses = {"jan":"01","fev":"02","mar":"03","abr":"04","mai":"05","jun":"06","jul":"07","ago":"08","set":"09","out":"10","nov":"11","dez":"12"}
    dia, mes_txt, ano, hora = m.groups()
    return f"{int(dia):02d}/{meses.get(mes_txt.lower(),'')}/{ano}", hora

# ========== Encontrar DESTINATÁRIO via bloco ==========
def find_destino_from_city_block(text: str) -> str:
    """
    Procura ocorrências CIDADE - UF e escolhe a que tem contexto de destinatário.
    Se não achar por contexto, retorna a última ocorrência.
    """
    lines = [ln.strip() for ln in re.split(r"\r\n|\r|\n", text) if ln.strip()]
    city_matches = []
    for i, ln in enumerate(lines):
        m = re.search(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ\.\s]+?)\s*[-–]\s*([A-Z]{2})\b", ln)
        if m:
            city_matches.append((i, m))

    if not city_matches:
        return ""

    # procurar ocorrência com contexto de destinatário nas 3 linhas anteriores
    for lineno, m in city_matches:
        context = " ".join(lines[max(0, lineno-3):lineno+1])
        if re.search(r"(DESTINATÁRIO|DESTINATARIO|REMESSA|CNPJ|RUA|AVENIDA|AV\.|GALPAO|GALPÃO|RODOVIA|ENDEREÇO|ENDERECO)", context, re.I):
            cidade = m.group(1).strip()
            uf = m.group(2).strip().upper()
            if uf in UF_VALIDAS:
                return f"{normalize(cidade).upper()} - {uf}"

    # fallback: retornar a última ocorrência
    lineno, m = city_matches[-1]
    cidade = m.group(1).strip()
    uf = m.group(2).strip().upper()
    if uf in UF_VALIDAS:
        return f"{normalize(cidade).upper()} - {uf}"
    return ""

# ========== Extração de valores (robusta) ==========
def parse_money_to_float(mtext: str) -> float | None:
    """
    Normaliza uma string numérica (varios formatos) para float (valor em reais).
    Ex: "1.234,56" -> 1234.56 ; "1,234.56" -> 1234.56 ; "1234.56" -> 1234.56
    """
    if not mtext:
        return None
    # remove espaços
    t = mtext.strip()
    # retirar tudo que não é digito, ponto, vírgula
    t = re.sub(r"[^\d\.,]", "", t)
    # se houver tanto . quanto , decidir formato
    if "." in t and "," in t:
        # se último separador é vírgula, então vírgula = decimal
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            # caso raro: ponto decimal (ex: 1234.56) e vírgula miles? tratar removendo vírgulas
            t = t.replace(",", "")
    else:
        # se só vírgula, é decimal em pt-BR
        if "," in t and "." not in t:
            t = t.replace(".", "").replace(",", ".")
        else:
            # t contém só pontos ou só dígitos (ex: 1.234 -> 1234)
            t = t.replace(",", "")
    try:
        return float(t)
    except:
        return None

def extract_valor_from_text(full_text: str) -> str:
    """
    Tenta extrair o valor total do manifesto a partir do texto nativo.
    Estratégia:
      1) procurar por linhas com 'VALOR TOTAL' e extrair valor próximo;
      2) senão, coletar todos padrões monetários e escolher o último (ou o que estiver junto a 'TOTAL').
    Retorna string formatada em pt-BR (ex: '1.234,56') ou ''.
    """
    if not full_text:
        return ""

    txt = full_text.upper()
    lines = [ln.strip() for ln in re.split(r"\r\n|\r|\n", txt) if ln.strip()]

    # 1) procurar por palavras-chave
    for ln in lines:
        if "VALOR TOTAL" in ln or "VALOR TOTAL DO MANIFESTO" in ln or re.search(r"\bVALOR.*MANIFESTO\b", ln):
            m = re.search(r"([\d\.,\s]+)", ln)
            if m:
                valf = parse_money_to_float(m.group(1))
                if valf is not None:
                    return f"{valf:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    # 2) procurar por padrões de dinheiro em todo o texto e escolher o mais provável:
    money_patterns = re.findall(r"(\d{1,3}(?:[.,\s]\d{3})*(?:[.,]\d{2}))", txt)
    if money_patterns:
        # preferir valores que aparecem em linhas contendo TOTAL
        for ln in reversed(lines):
            if "TOTAL" in ln:
                m = re.search(r"(\d{1,3}(?:[.,\s]\d{3})*(?:[.,]\d{2}))", ln)
                if m:
                    valf = parse_money_to_float(m.group(1))
                    if valf is not None:
                        return f"{valf:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        # senão, pegar o último padrão geral
        last = money_patterns[-1]
        valf = parse_money_to_float(last)
        if valf is not None:
            return f"{valf:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    return ""

# ========== Extração manifesto / destino (regra final) ==========
def extract_manifesto_destino_from_text(full_text: str):
    norm = normalize(full_text).upper()

    # Manifesto
    m = re.search(r"\b(?:NUM[EÉ]RO|N[ºO])\s*[:\-]?\s*(\d{8,15})\b", norm)
    manifesto = m.group(1) if m else ""
    if not manifesto:
        m2 = re.search(r"\b(\d{10,15})\b", norm)
        manifesto = m2.group(1) if m2 else ""

    # procurar Sxx
    rot = re.search(r"\bS[-\s]*([0-9]{1,2})\b", norm)
    if rot:
        rota_code = "S" + str(int(rot.group(1)))
        if rota_code in ROTA_CO_MAP:
            return manifesto, ROTA_CO_MAP[rota_code].upper()
        # se Sxx não mapeado, prossegue para identificar o destinatário

    # tentar identificar o destinatário pelo bloco (com contexto)
    destino = find_destino_from_city_block(norm)
    if destino:
        return manifesto, destino

    # fallback: última ocorrência CIDADE - UF
    matches = list(re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ\s\.]+?)\s*[-–]\s*([A-Z]{2})\b", norm))
    if matches:
        cidade, uf = matches[-1].groups()
        cidade = cidade.strip()
        uf = uf.strip().upper()
        if uf in UF_VALIDAS:
            return manifesto, f"{cidade} - {uf}"

    return manifesto, ""

# ========== OCR helpers ==========
def ocr_page_bytes(file_bytes: bytes, page_index="first", dpi=500):
    images = convert_from_bytes(file_bytes, dpi=dpi, fmt="jpeg")
    if not images:
        return None, ""
    img = images[0] if page_index == "first" else images[-1]
    img_p = pil_preprocess(img)
    txt = ocr_image(img_p, psm=6)
    if len(txt.strip()) < 10:
        txt = ocr_image(img_p, psm=3)
    return img_p, txt

# ========== Processamento do PDF ==========
def process_pdf(file_bytes: bytes, want_debug: bool = False):
    out = {"manifesto": "", "data": "", "hora": "", "destino": "", "valor": "", "volumes": "", "debug": {}}

    # texto nativo
    pages_txt = read_pdf_text(file_bytes)
    full_text = "\n".join(pages_txt) if pages_txt else ""

    if pages_txt:
        out["data"], out["hora"] = extract_data_hora_from_head(pages_txt[0])
        manifesto, destino = extract_manifesto_destino_from_text(full_text)
        out["manifesto"], out["destino"] = manifesto, destino

        # tentar extrair valor do texto nativo (robusto)
        valor_txt = extract_valor_from_text(full_text)
        if valor_txt:
            out["valor"] = valor_txt

    # fallback manifesto via OCR 1ª página
    if not out["manifesto"]:
        _img1, ocr1 = ocr_page_bytes(file_bytes, page_index="first", dpi=500)
        out["debug"]["OCR_1a_PAG"] = ocr1
        m = re.search(r"\b(?:NUM[EÉ]RO|N[ºO])\s*[:\-]?\s*(\d{8,15})\b", ocr1)
        if m:
            out["manifesto"] = m.group(1)

    # OCR última página
    _imgL, ocrL = ocr_page_bytes(file_bytes, page_index="last", dpi=500)
    out["debug"]["OCR_ULTIMA_PAG"] = ocrL

    # se ainda não extraiu valor do texto nativo, tentar no OCR
    if not out["valor"]:
        # primeiro, procurar por "VALOR TOTAL" no OCR bruto
        m_val = re.search(r"VALOR\s+TOTAL(?:\s+DO\s+MANIFESTO)?\s*[:\-]?\s*([\d\.,\s]+)", ocrL or "", re.I)
        if m_val:
            valf = parse_money_to_float(m_val.group(1))
            if valf is not None:
                out["valor"] = f"{valf:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            # se não, tentar extrair último padrão monetário do OCR
            money_patterns = re.findall(r"(\d{1,3}(?:[.,\s]\d{3})*(?:[.,]\d{2}))", ocrL or "")
            if money_patterns:
                last = money_patterns[-1]
                valf = parse_money_to_float(last)
                if valf is not None:
                    out["valor"] = f"{valf:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    # volumes (OCR)
    if not out["volumes"]:
        m_vol = re.search(r"\bVOLUMES?\b\s*[:\-]?\s*([0-9]{1,6})\b", (full_text or "") + "\n" + (ocrL or ""), re.I)
        if m_vol:
            out["volumes"] = clean_int(m_vol.group(1))

    # fallback destino via OCR (se ainda vazio)
    if not out["destino"]:
        mds = list(re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ\s\.]+?)\s*[-–]\s*([A-Z]{2})", ocrL or ""))
        if mds:
            cidade, uf = mds[-1].groups()
            out["destino"] = f"{cidade.strip().upper()} - {uf.upper()}"

    # debug
    if want_debug:
        with st.expander("🔍 Mostrar Debug Completo do OCR", expanded=False):
            st.markdown("### 🧠 Texto Bruto Extraído (para diagnóstico)")
            with st.expander("🗂️ OCR — 1ª Página", expanded=False):
                st.code(out["debug"].get("OCR_1a_PAG", "(sem texto)"))
            with st.expander("📄 OCR — Última Página", expanded=False):
                st.code(out["debug"].get("OCR_ULTIMA_PAG", "(sem texto)"))

    return out

# ========== Interface ==========
st.title("📦 Leitor de Manifestos Jadlog")
st.caption("Extrai Manifesto, Data, Destino, Valor Total e Volumes — com OCR inteligente.")

responsavel = st.text_input("Responsável", placeholder="Digite o nome completo")
want_debug = st.checkbox("Mostrar debug do OCR (texto bruto)", value=False)

files = st.file_uploader("Envie PDFs do manifesto", type=["pdf"], accept_multiple_files=True)

if files:
    linhas = []
    for f in files:
        try:
            pdf_bytes = f.read()
            result = process_pdf(pdf_bytes, want_debug=want_debug)

            linhas.append({
                "Data": result["data"],
                "Manifesto": result["manifesto"],
                "Destino": result["destino"],
                "Referência": "",
                "Responsável": (responsavel or "").upper(),
                "Valor total": result["valor"],
                "Quantidade": result["volumes"],
            })

            ok = all([result["manifesto"], result["data"], result["destino"]])
            if ok:
                st.success(f"✅ {f.name} | {result['destino']} | {result['valor'] or 'valor indefinido'}")
            else:
                st.warning(f"⚠️ {f.name} — faltou algum campo (Manifesto/Data/Destino)")

        except Exception as e:
            st.error(f"Erro ao processar {f.name}: {e}")

    df = pd.DataFrame(linhas, columns=["Data","Manifesto","Destino","Referência","Responsável","Valor total","Quantidade"])
    st.subheader("Prévia — MANIFESTOS")
    st.dataframe(df, hide_index=True, use_container_width=True)

    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    st.download_button("📥 Baixar Planilha Operacional", data=buf,
                       file_name=f"OPERACIONAL_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Envie 1 ou mais PDFs para extrair informações automaticamente.")
