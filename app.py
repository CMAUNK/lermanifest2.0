# app.py
# -----------------------------------------------
# Leitor de Manifestos Jadlog — versão otimizada, fiel ao código original
# - Igual ao seu código atual, apenas MUITO mais rápido
# - OCR rápido: 300 DPI primeiro, 500 DPI apenas se necessário
# - Todo o resto mantido exatamente igual
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
#  CONFIGURAÇÃO DA PÁGINA
# ==========================
st.set_page_config(page_title="Leitor de Manifestos Jadlog", page_icon="🚛", layout="centered")
st.markdown(
    """
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
""",
    unsafe_allow_html=True,
)

# ==========================
#  CONSTANTES / MAPAS
# ==========================
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
    "AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT", "PA",
    "PB", "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO"
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

def ocr_image_fast(img, psm=6):
    cfg = f"--psm {psm} --oem 3"
    return pytesseract.image_to_string(img, lang="por+eng", config=cfg).upper()

# ==========================
#  OCR RÁPIDO
# ==========================
def ocr_page_bytes(file_bytes: bytes, page_index="first"):
    """
    OCR rápido: tenta 300 DPI primeiro e só usa 500 DPI se necessário.
    """
    # 1) tentativa rápida (300 DPI)
    images = convert_from_bytes(file_bytes, dpi=300, fmt="jpeg")
    img = images[0] if page_index == "first" else images[-1]
    img_p = pil_preprocess(img)
    txt = ocr_image_fast(img_p, psm=6)

    # se veio fraco → tenta 500 DPI
    if len(txt.strip()) < 15:
        images = convert_from_bytes(file_bytes, dpi=500, fmt="jpeg")
        img = images[0] if page_index == "first" else images[-1]
        img_p = pil_preprocess(img)
        txt = ocr_image_fast(img_p, psm=6)
        if len(txt.strip()) < 10:
            txt = ocr_image_fast(img_p, psm=3)

    return img_p, txt


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
        "jan": "01", "fev": "02", "mar": "03", "abr": "04", "mai": "05", "jun": "06",
        "jul": "07", "ago": "08", "set": "09", "out": "10", "nov": "11", "dez": "12"
    }
    dia, mes_txt, ano, _hora = m.groups()
    return f"{int(dia):02d}/{meses.get(mes_txt.lower(),'')}/{ano}", _hora

# ==========================
#  DESTINO
# ==========================
def extract_manifesto_destino_from_text(full_text: str):
    norm = normalize(full_text or "").upper()

    # Manifesto
    m = re.search(r"\b(?:NUM[EÉ]RO|N[ºO])\s*[:\-]?\s*(\d{8,})", norm)
    manifesto = m.group(1) if m else ""
    if not manifesto:
        m2 = re.search(r"\b(\d{10,15})\b", norm)
        manifesto = m2.group(1) if m2 else ""

    destino = ""

    # 1) via Sxx
    rota = re.search(r"\bS\s*[-\/]?\s*0*([0-9]{1,2})\b", norm)
    if rota:
        code = "S" + str(int(rota.group(1)))
        destino = ROTA_CO_MAP.get(code, "")

    # 2) fallback cidade - UF
    if not destino:
        mds = list(re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ .-]+)\s*-\s*([A-Z]{2})", norm))
        for m in reversed(mds):
            cidade, uf = m.groups()
            if uf.upper() in UF_VALIDAS:
                destino = f"{cidade.strip().upper()} - {uf.upper()}"
                break

    return manifesto, destino


# ==========================
#  PROCESSAMENTO
# ==========================
def process_pdf(file_bytes: bytes, want_debug: bool = False):
    out = {"manifesto": "", "data": "", "hora": "", "destino": "", "valor": "", "volumes": "", "debug": {}}

    pages_txt = read_pdf_text(file_bytes)
    if pages_txt:
        out["data"], out["hora"] = extract_data_hora_from_head(pages_txt[0])
        manifesto, destino = extract_manifesto_destino_from_text("\n".join(pages_txt))
        out["manifesto"], out["destino"] = manifesto, destino

    # OCR 1ª página
    if not out["manifesto"]:
        _img1, ocr1 = ocr_page_bytes(file_bytes, page_index="first")
        out["debug"]["OCR_1a_PAG"] = ocr1
        m = re.search(r"\b(?:NUM[EÉ]RO|N[ºO])\s*[:\-]?\s*(\d{8,})", ocr1)
        if not m:
            m = re.search(r"\b(\d{10,15})\b", ocr1)
        if m:
            out["manifesto"] = m.group(1)

    # OCR última página
    _imgL, ocrL = ocr_page_bytes(file_bytes, page_index="last")
    out["debug"]["OCR_ULTIMA_PAG"] = ocrL

    # Valor total (igual ao original)
    if not out["valor"]:
        for linha in (ocrL or "").splitlines():
            if "VALOR TOTAL DO MANIFESTO" in linha:
                m_val = re.search(r"([\d\.,]+)", linha)
                if m_val:
                    val = m_val.group(1)
                    val = val.replace(" ", "").strip()

                    if "." in val and "," in val:
                        if val.rfind(".") > val.rfind(","):
                            val = val.replace(",", "")
                        else:
                            val = val.replace(".", "").replace(",", ".")
                    elif "," in val:
                        val = val.replace(".", "").replace(",", ".")
                    else:
                        val = val.replace(",", "")

                    try:
                        v = float(val)
                        out["valor"] = (
                            f"{v:,.2f}"
                            .replace(",", "X")
                            .replace(".", ",")
                            .replace("X", ".")
                        )
                    except:
                        out["valor"] = val
                break

    # ==========================
    #  VOLUMES (IGUAL AO ORIGINAL, APENAS COM OCR RÁPIDO)
    # ==========================
    total_volumes = 0
    images = convert_from_bytes(file_bytes, dpi=300, fmt="jpeg")

    for i, img in enumerate(images):
        img_p = pil_preprocess(img)
        txt = ocr_image_fast(img_p, psm=6)

        if len(txt.strip()) < 15:
            pages500 = convert_from_bytes(file_bytes, dpi=500, fmt="jpeg")
            img_p = pil_preprocess(pages500[i])
            txt = ocr_image_fast(img_p, psm=6)

        # padrão oficial
        m1 = re.search(r"VOLUMES?\s*[:\-]?\s*([0-9]+)", txt)
        if m1:
            total_volumes += int(m1.group(1))
            continue

        # padrão final "31.13 74.00"
        m2 = re.search(r"\b([0-9]{1,5})\s*\.\s*[0-9]{1,3}\b", txt)
        if m2:
            total_volumes += int(m2.group(1))

    if total_volumes > 0:
        out["volumes"] = str(total_volumes)

    return out


# ==========================
#  INTERFACE
# ==========================
st.title("📦 Leitor de Manifestos Jadlog")
st.caption("Extrai Manifesto, Data, Destino, Valor Total e Volumes — versão otimizada.")

responsavel = st.text_input("Responsável", placeholder="Digite o nome completo")
want_debug = st.checkbox("Mostrar debug do OCR (texto bruto)", value=False)

files = st.file_uploader("Envie PDFs de manifesto", type=["pdf"], accept_multiple_files=True)

if files:
    linhas = []
    for f in files:
        f_bytes = f.read()
        result = process_pdf(f_bytes, want_debug=want_debug)

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
        st.success(f"✔️ {f.name} processado") if ok else st.warning(f"⚠️ {f.name} incompleto")

    df = pd.DataFrame(linhas)

    st.subheader("Prévia — MANIFESTOS")
    st.dataframe(df, use_container_width=True, hide_index=True)

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
    st.info("Envie PDFs para iniciar o processamento.")
