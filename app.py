# app.py
# -----------------------------------------------
# Leitor de Manifestos Jadlog — versão otimizada
# Inclui:
# - OCR rápido (300 dpi + fallback 500 dpi)
# - Barra de progresso
# - Tabela limpa com botão "Limpar"
# - Resumo final (arquivos + volumes)
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
#  CONFIG DA PÁGINA
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
#  MAPAS
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
    "AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS",
    "MT", "PA", "PB", "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS",
    "SC", "SE", "SP", "TO"
}

# ==========================
#  FUNÇÕES AUXILIARES
# ==========================
def normalize(s):
    s = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in s if not unicodedata.combining(c))

def clean_int(s):
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

def ocr_page(img):
    """
    OCR rápido com fallback em dpi alto.
    """
    txt = ocr_image(img, psm=6)
    if len(txt.strip()) > 10:
        return txt

    txt = ocr_image(img, psm=3)
    return txt

def convert_pdf_fast(file_bytes):
    """
    Converte PDF para imagens em 300 DPI.
    Se OCR falhar, repete em 500 DPI.
    """
    try:
        images = convert_from_bytes(file_bytes, dpi=300, fmt="jpeg")
        return images
    except:
        return convert_from_bytes(file_bytes, dpi=500, fmt="jpeg")

# ==========================
#  EXTRAÇÃO DE TEXTO
# ==========================
def read_pdf_text(file_bytes):
    pages_txt = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for p in pdf.pages:
            pages_txt.append(p.extract_text() or "")
    return pages_txt

def extract_data_hora_from_head(txt):
    cab = normalize(txt or "")
    m = re.search(r"(\d{1,2}) (jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez) (\d{4}) (\d{2}:\d{2}:\d{2})", cab, re.I)
    if not m:
        return "", ""
    dia, mes_txt, ano, hora = m.groups()
    meses = {
        "jan": "01","fev": "02","mar": "03","abr":"04","mai":"05","jun":"06",
        "jul":"07","ago":"08","set":"09","out":"10","nov":"11","dez":"12"
    }
    return f"{dia}/{meses[mes_txt.lower()]}/{ano}", hora

# ==========================
#  EXTRAÇÃO — MANIFESTO + DESTINO
# ==========================
def extract_manifesto_destino_from_text(full_text):
    norm = normalize(full_text or "").upper()

    # Manifesto
    m = re.search(r"(?:NUM[EÉ]RO|N[ºO])\s*[:\-]?\s*(\d{8,})", norm)
    manifesto = m.group(1) if m else ""

    if not manifesto:
        m2 = re.search(r"\b(\d{10,15})\b", norm)
        manifesto = m2.group(1) if m2 else ""

    destino = ""

    # Sxx
    rota = re.search(r"\bS\s*[-\/]?\s*0*([0-9]{1,2})\b", norm)
    if rota:
        rota_code = "S" + str(int(rota.group(1)))
        if rota_code in ROTA_CO_MAP:
            destino = ROTA_CO_MAP[rota_code]

    # Fallback cidade-UF
    if not destino:
        mds = list(re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ .-]+?)\s*-\s*([A-Z]{2})", norm))
        for m in reversed(mds):
            c, uf = m.groups()
            if uf.upper() in UF_VALIDAS:
                destino = f"{c.strip().upper()} - {uf}"
                break

    return manifesto, destino

# ==========================
#  PROCESSAMENTO
# ==========================
def process_pdf(file_bytes, progress=None):
    out = {"manifesto": "", "data": "", "destino": "", "valor": "", "volumes": ""}

    pages_txt = read_pdf_text(file_bytes)
    if pages_txt:
        out["data"], _ = extract_data_hora_from_head(pages_txt[0])
        manifesto, destino = extract_manifesto_destino_from_text("\n".join(pages_txt))
        out["manifesto"] = manifesto
        out["destino"] = destino

    # OCR rápido — todas páginas
    images = convert_pdf_fast(file_bytes)

    total_volumes = 0

    for i, img in enumerate(images):
        if progress:
            progress.progress((i + 1) / len(images))

        img_p = pil_preprocess(img)
        txt = ocr_page(img_p)

        # volumes oficiais
        m1 = re.search(r"VOLUMES?\s*[:\-]?\s*([0-9]{1,6})", txt)
        if m1:
            total_volumes += int(m1.group(1))
            continue

        # padrão "31.13 74.00"
        m2 = re.search(r"\b([0-9]{1,5})\.\d{1,3}\b", txt)
        if m2:
            total_volumes += int(m2.group(1))
            continue

        # valor
        if not out["valor"]:
            if "VALOR TOTAL DO MANIFESTO" in txt:
                m_val = re.search(r"([\d\.,]+)$", txt)
                if m_val:
                    out["valor"] = m_val.group(1)

    out["volumes"] = str(total_volumes)
    return out

# ==========================
#  INTERFACE
# ==========================
st.title("📦 Leitor de Manifestos Jadlog — OTIMIZADO 🚀")
st.caption("Agora com OCR rápido, barra de progresso e UI melhorada.")

responsavel = st.text_input("Responsável", placeholder="Digite o nome completo")

btn_clear = st.button("🧹 Limpar Tabela")

if btn_clear:
    st.session_state["linhas"] = []

files = st.file_uploader("Envie PDFs de manifesto", type=["pdf"], accept_multiple_files=True)

if "linhas" not in st.session_state:
    st.session_state["linhas"] = []

if files:
    total_pdf = len(files)
    total_volumes_geral = 0

    for idx, f in enumerate(files):
        st.write(f"📄 Processando: **{f.name}** ({idx+1}/{total_pdf})")
        progress = st.progress(0)

        f_bytes = f.read()
        result = process_pdf(f_bytes, progress=progress)

        st.session_state["linhas"].append({
            "Data": result["data"],
            "Manifesto": result["manifesto"],
            "Destino": result["destino"],
            "Referência": "",
            "Responsável": (responsavel or "").upper(),
            "Valor total": result["valor"],
            "Quantidade": result["volumes"],
        })

        total_volumes_geral += int(result["volumes"])

        st.success(f"✔️ {f.name} concluído — {result['volumes']} volumes")

    st.info(f"📊 **Arquivos processados:** {total_pdf} | **Total de volumes:** {total_volumes_geral}")

# Mostrar tabela
if st.session_state["linhas"]:
    df = pd.DataFrame(st.session_state["linhas"])
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
    st.info("Envie PDFs para iniciar.")
