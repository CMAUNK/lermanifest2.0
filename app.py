# app.py
# -----------------------------------------------
# Leitor de Manifestos Jadlog — OCR Final (v5.0)
# - Extrai Manifesto, Data, Destino (CO via Sxx ou Destinatário), Valor e Volumes
# - Usa texto nativo do PDF quando possível (pdfplumber)
# - Fallback via OCR com Tesseract + pdf2image
# -----------------------------------------------

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

# ==========================
#  CONFIGURAÇÃO STREAMLIT
# ==========================
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

# ==========================
#  MAPA DE ROTAS Sxx → CO
# ==========================
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

# ==========================
#  FUNÇÕES AUXILIARES
# ==========================
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

# ==========================
#  EXTRAÇÃO DE TEXTO
# ==========================
def read_pdf_text(file_bytes: bytes):
    pages_txt = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for p in pdf.pages:
            pages_txt.append(p.extract_text() or "")
    return pages_txt

def extract_data_hora_from_head(first_page_text: str):
    cab = normalize(first_page_text or "")
    m = re.search(
        r"\b(\d{1,2})\s+(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\s+(\d{4})\s+(\d{2}:\d{2}:\d{2})",
        cab, re.I
    )
    if not m:
        return "", ""
    meses = {
        "jan":"01","fev":"02","mar":"03","abr":"04","mai":"05","jun":"06",
        "jul":"07","ago":"08","set":"09","out":"10","nov":"11","dez":"12"
    }
    dia, mes_txt, ano, hora = m.groups()
    return f"{int(dia):02d}/{meses.get(mes_txt.lower(),'')}/{ano}", hora

# ==========================
#  REGRA FINAL DO DESTINO
# ==========================
def extract_manifesto_destino_from_text(full_text: str):
    """
    Regra FINAL:
      1. Se houver Sxx mapeado → destino = CO da rota.
      2. Se houver Sxx não mapeado → ignorar e usar o DESTINATÁRIO.
      3. Se não houver Sxx → destino = DESTINATÁRIO.
    """
    norm = normalize(full_text).upper()

    # --- Manifesto ---
    m = re.search(r"\b(?:NUM[EÉ]RO|N[ºO])\s*[:\-]?\s*(\d{8,15})\b", norm)
    manifesto = m.group(1) if m else ""
    if not manifesto:
        m2 = re.search(r"\b(\d{10,15})\b", norm)
        manifesto = m2.group(1) if m2 else ""

    # --- Busca Sxx ---
    rot = re.search(r"\bS[-\s]*([0-9]{1,2})\b", norm)

    if rot:
        rota_code = "S" + str(int(rot.group(1)))
        # Caso Sxx esteja mapeado → usar CO correspondente
        if rota_code in ROTA_CO_MAP:
            return manifesto, ROTA_CO_MAP[rota_code].upper()
        # Se Sxx não estiver mapeado → seguir para DESTINATÁRIO

    # --- DESTINATÁRIO: última ocorrência de CIDADE - UF ---
    matches = list(
        re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ\s\.]+?)\s*[-–]\s*([A-Z]{2})\b", norm)
    )

    if matches:
        cidade, uf = matches[-1].groups()
        cidade = cidade.strip()
        uf = uf.strip().upper()
        if uf in UF_VALIDAS:
            return manifesto, f"{cidade} - {uf}"

    return manifesto, ""

# ==========================
#  OCR
# ==========================
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

# ==========================
#  PROCESSAMENTO DE PDF
# ==========================
def process_pdf(file_bytes: bytes, want_debug: bool = False):
    out = {"manifesto": "", "data": "", "hora": "", "destino": "", "valor": "", "volumes": "", "debug": {}}

    pages_txt = read_pdf_text(file_bytes)

    if pages_txt:
        out["data"], out["hora"] = extract_data_hora_from_head(pages_txt[0])
        manifesto, destino = extract_manifesto_destino_from_text("\n".join(pages_txt))
        out["manifesto"] = manifesto
        out["destino"] = destino

    # Fallback OCR manifesto
    if not out["manifesto"]:
        _, ocr1 = ocr_page_bytes(file_bytes, page_index="first", dpi=500)
        out["debug"]["OCR_1a_PAG"] = ocr1
        m = re.search(r"\b(?:NUM[EÉ]RO|N[ºO])\s*[:\-]?\s*(\d{8,15})\b", ocr1)
        if m:
            out["manifesto"] = m.group(1)

    # OCR última página
    _, ocrL = ocr_page_bytes(file_bytes, page_index="last", dpi=500)
    out["debug"]["OCR_ULTIMA_PAG"] = ocrL

    # Valor total
    if not out["valor"]:
        m_val = re.search(r"VALOR TOTAL DO MANIFESTO\s*[:\-]?\s*([\d\.,]+)", ocrL or "")
        if m_val:
            val = m_val.group(1)
            val = val.replace(".", "").replace(",", ".")
            try:
                valor = float(val)
                out["valor"] = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except:
                out["valor"] = val

    # Volumes
    if not out["volumes"]:
        m_vol = re.search(r"\bVOLUMES?\b\s*[:\-]?\s*([0-9]{1,6})\b", ocrL or "")
        if m_vol:
            out["volumes"] = clean_int(m_vol.group(1))

    # Fallback destino pelo OCR
    if not out["destino"]:
        mds = list(re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ\s]+?)\s*-\s*([A-Z]{2})", ocrL or ""))
        if mds:
            cidade, uf = mds[-1].groups()
            out["destino"] = f"{cidade.strip().upper()} - {uf.upper()}"

    return out

# ==========================
#  INTERFACE STREAMLIT
# ==========================
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
                st.success(f"✅ {f.name} | {result['destino']}")
            else:
                st.warning(f"⚠️ {f.name} — faltou algum campo")

        except Exception as e:
            st.error(f"Erro ao processar {f.name}: {e}")

    df = pd.DataFrame(linhas)
    st.subheader("Prévia — MANIFESTOS")
    st.dataframe(df, hide_index=True, use_container_width=True)

    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)

    st.download_button(
        "📥 Baixar Planilha Operacional",
        data=buf,
        file_name=f"OPERACIONAL_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Envie 1 ou mais PDFs para extrair informações automaticamente.")
