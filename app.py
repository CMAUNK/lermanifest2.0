# app.py
# ------------------------------------------------------------
# Leitor de Manifestos Jadlog — OCR Final (v4.7 otimizado)
# ------------------------------------------------------------
# Extrai:
#   - Manifesto
#   - Data
#   - Destino
#   - Valor total
#   - Quantidade total (somando todas as páginas)
#
# Usa:
#   - pdfplumber (texto nativo)
#   - pytesseract + pdf2image (OCR fallback)
#
# Interface:
#   - Streamlit (limpa, rápida, sem debug indesejado)
# ------------------------------------------------------------

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

# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
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

# ------------------------------------------------------------
# MAPAS
# ------------------------------------------------------------
ROTA_CO_MAP = {
    "S27": "CO RIO DE JANEIRO 05",
    "S28": "CO RIO DE JANEIRO 04",
    "S30": "CO RIO DE JANEIRO 01",
    "S32": "CO RIO DE JANEIRO 03",
    "S38": "CO RIO DE JANEIRO 06",
    "S41": "CO RIO DE JANEIRO 08",
    "S49": "CO RIO DE JANEIRO 07",
}

UF_VALIDAS = {
    "AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI",
    "PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO"
}

# ------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ------------------------------------------------------------
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
    return pytesseract.image_to_string(
        img, lang="por+eng", config=f"--psm {psm} --oem 3"
    ).upper()

# ------------------------------------------------------------
# TEXTO DO PDF
# ------------------------------------------------------------
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
    dia, mes_txt, ano, _hora = m.groups()
    return f"{int(dia):02d}/{meses.get(mes_txt.lower(),'')}/{ano}", _hora

# ------------------------------------------------------------
# MANIFESTO + DESTINO
# ------------------------------------------------------------
def extract_manifesto_destino_from_text(full_text: str):
    norm = normalize(full_text or "").upper()

    # Manifesto
    m = re.search(r"\b(?:NUM[EÉ]RO|N[ºO])\s*[:\-]?\s*(\d{8,})", norm)
    manifesto = m.group(1) if m else ""

    if not manifesto:
        m2 = re.search(r"\b(\d{10,15})\b", norm)
        manifesto = m2.group(1) if m2 else ""

    destino = ""

    # Sxx
    rota = re.search(r"\bS\s*[-\/]?\s*0*([0-9]{1,2})\b", norm)
    if rota:
        code = "S" + str(int(rota.group(1)))
        destino = ROTA_CO_MAP.get(code, "")

    if not destino:
        for k, v in ROTA_CO_MAP.items():
            if re.search(fr"S\s*[-\/]?\s*0*{k[1:]}", norm):
                destino = v
                break

    if not destino:
        matches = list(re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ \.\-]+?)\s*-\s*([A-Z]{2})", norm))
        for m in reversed(matches):
            cidade, uf = m.groups()
            if uf in UF_VALIDAS:
                destino = f"{cidade.strip().upper()} - {uf.upper()}"
                break

    return manifesto, destino

# ------------------------------------------------------------
# OCR POR PÁGINA
# ------------------------------------------------------------
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

# ------------------------------------------------------------
# PROCESSAMENTO PRINCIPAL
# ------------------------------------------------------------
def process_pdf(file_bytes: bytes, want_debug=False):

    out = {"manifesto":"", "data":"", "hora":"", "destino":"",
           "valor":"", "volumes":"", "debug":{}}

    # --------- Texto nativo ----------
    pages_txt = read_pdf_text(file_bytes)
    if pages_txt:
        out["data"], out["hora"] = extract_data_hora_from_head(pages_txt[0])
        manifesto, destino = extract_manifesto_destino_from_text("\n".join(pages_txt))
        out["manifesto"], out["destino"] = manifesto, destino

    # --------- OCR 1ª página ----------
    if not out["manifesto"]:
        _img1, ocr1 = ocr_page_bytes(file_bytes, page_index="first", dpi=300)
        out["debug"]["OCR_1a_PAG"] = ocr1

        m = re.search(r"\b(?:NUM[EÉ]RO|N[ºO])[:\-]?\s*(\d{8,})", ocr1)
        if not m:
            m = re.search(r"\b(\d{10,15})\b", ocr1)
        if m:
            out["manifesto"] = m.group(1)

    # --------- OCR última página ----------
    _imgL, ocrL = ocr_page_bytes(file_bytes, page_index="last", dpi=300)
    out["debug"]["OCR_ULTIMA_PAG"] = ocrL

    # --------- Valor total ----------
    linha_valor = ""
    for linha in (ocrL or "").splitlines():
        if "VALOR TOTAL DO MANIFESTO" in linha:
            linha_valor = linha.strip()
            break

    if linha_valor:
        m_val = re.search(r"VALOR\s+TOTAL\s+DO\s+MANIFESTO\s*[:\-]?\s*([\d\s\.,]+)", linha_valor)
        if m_val:
            raw = re.sub(r"[^0-9\.,]", "", m_val.group(1)).replace(" ", "")

            if "," in raw and "." in raw:
                if raw.rfind(".") > raw.rfind(","):
                    raw = raw.replace(",", "")
                else:
                    raw = raw.replace(".", "").replace(",", ".")
            elif "," in raw:
                raw = raw.replace(".", "").replace(",", ".")
            raw = raw.replace(",", "")

            try:
                val = float(raw)
                out["valor"] = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except:
                out["valor"] = raw

    # --------- Soma de volumes ----------
    total_vol = 0
    images = convert_from_bytes(file_bytes, dpi=300, fmt="jpeg")

    for img in images:
        img_p = pil_preprocess(img)
        txt = ocr_image(img_p, psm=6)
        if len(txt.strip()) < 10:
            txt = ocr_image(img_p, psm=3)

        m1 = re.search(r"\bVOLUMES?\b[:\-]?\s*([0-9]{1,6})\b", txt)
        if m1:
            total_vol += int(m1.group(1))
            continue

        m2 = re.search(r"\b([0-9]{1,5})\s*\.\s*[0-9]{1,3}\b", txt)
        if m2:
            try:
                total_vol += int(m2.group(1))
                continue
            except:
                pass

    if total_vol > 0:
        out["volumes"] = str(total_vol)

    return out

# ------------------------------------------------------------
# INTERFACE STREAMLIT
# ------------------------------------------------------------
st.title("📦 Leitor de Manifestos Jadlog")

responsavel = st.text_input("Responsável", placeholder="Digite o nome completo")
want_debug = st.checkbox("Mostrar debug OCR", value=False)

files = st.file_uploader("Envie PDFs", type=["pdf"], accept_multiple_files=True)

if files:

    linhas = []

    for f in files:
        f_bytes = f.read()
        result = process_pdf(f_bytes, want_debug)

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

        # --------- CORRIGIDO: sem inline ternário ---------
        if ok:
            st.success(f"✔️ {f.name} processado")
        else:
            st.warning(f"⚠️ {f.name} incompleto")

    df = pd.DataFrame(linhas)

    st.subheader("Prévia — MANIFESTOS")
    st.dataframe(df, use_container_width=True, hide_index=True)

    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)

    st.download_button(
        "📥 Baixar Planilha",
        data=buf,
        file_name=f"OPERACIONAL_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Envie 1 ou mais PDFs de manifesto para extrair automaticamente.")
