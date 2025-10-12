import re
import io
import fitz
import streamlit as st
import pandas as pd
from pdf2image import convert_from_bytes
import pytesseract

# ------------------ DICIONÁRIO DE CÓDIGOS OPERACIONAIS ------------------
CODIGOS_CO = {
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
    "S58": "CO CAMPOS D. GOYTACAZES",
    "S59": "CO SANT. ANT DE PADUA",
    "S60": "CO DUQUE DE CAXAIS",
    "S61": "CO CABO FRIO",
    "S63": "CO ARARUAMA",
    "S64": "CO SAO GONCALO",
    "S65": "CO RIO DAS OSTRAS",
    "S67": "CO NITEROI",
    "S68": "CO MARICÁ",
    "S70": "FL RIO DE JANEIRO"
}

UF_VALIDAS = {"AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO"}

# ------------------ FUNÇÕES AUXILIARES ------------------
def ocr_image(image):
    return pytesseract.image_to_string(image, lang="por")

def pdf_to_text_ocr(file_bytes):
    images = convert_from_bytes(file_bytes, dpi=200)
    if not images:
        return "", ""
    text_first = ocr_image(images[0])
    text_last = ocr_image(images[-1])
    return text_first, text_last

def clean_int(value):
    try:
        return int(re.sub(r"\D", "", value))
    except:
        return None

# ------------------ EXTRAÇÃO ------------------
def extrair_dados(pdf_bytes):
    out = {"manifesto": "", "data": "", "hora": "", "destino": "", "valor": "", "volumes": ""}

    # Extrai OCR das páginas
    ocrF, ocrL = pdf_to_text_ocr(pdf_bytes)
    ocrAll = (ocrF or "") + "\n" + (ocrL or "")

    # ---- Manifesto e Data/Hora ----
    m_manifesto = re.search(r"N[ÚU]MERO[:\s]*([0-9]{8,15})", ocrAll, re.I)
    if m_manifesto:
        out["manifesto"] = m_manifesto.group(1).strip()

    m_data = re.search(r"(\d{1,2}[/\-\. ]\d{1,2}[/\-\. ]\d{2,4})", ocrAll)
    if m_data:
        out["data"] = m_data.group(1).strip().replace(".", "/").replace("-", "/")

    m_hora = re.search(r"(\d{2}[:h]\d{2}[:h]?\d{0,2})", ocrAll)
    if m_hora:
        out["hora"] = m_hora.group(1).replace("h", ":")

    # ---- Valor Total ----
    m_val = re.search(r"VALOR\s+TOTAL\s+DO\S*\s+MANIFESTO\S*[:\-]?\s*([\d\., ]+)", ocrL, re.I)
    if m_val:
        val_raw = m_val.group(1)
        val_text = val_raw.replace(" ", "").replace(",", ".")
        val_match = re.search(r"\d+\.\d{2}$", val_text)
        if not val_match:
            val_match = re.search(r"\d+[.,]\d{2}", val_text)
        if val_match:
            val_final = val_match.group(0).replace(",", ".")
            try:
                out["valor"] = f"{float(val_final):,.2f}".replace(",", "x").replace(".", ",").replace("x", ".")
            except:
                out["valor"] = val_final

    # ---- Volumes ----
    m_vol = re.search(r"\bVOLUMES[:\s]*([0-9]{1,5})", ocrL, re.I)
    if m_vol:
        out["volumes"] = clean_int(m_vol.group(1))

    # ---- Destino (cidade e UF) ----
    mds = list(re.finditer(r"([A-ZÇÃÉÓÍÚÂÊÔÀÈÙ ]+)\s*-\s*([A-Z]{2})", ocrL))
    for m in reversed(mds):
        cidade, uf = m.groups()
        if uf.upper() in UF_VALIDAS:
            out["destino"] = f"{cidade.strip().upper()} - {uf.upper()}"
            break

    # ---- Detecta CO (Sxx) ----
    m_co = re.search(r"\bS[\s\-]?\d{2}\b", ocrAll, flags=re.I)
    if m_co:
        co_code = m_co.group(0).upper().replace(" ", "").replace("-", "")
        if co_code in CODIGOS_CO:
            out["destino"] = f"{CODIGOS_CO[co_code]} ({co_code})"

    return out, ocrF, ocrL

# ------------------ INTERFACE STREAMLIT ------------------
st.set_page_config(page_title="Leitor de Manifestos Jadlog", layout="wide")

st.title("📦 Leitor de Manifestos Jadlog (OCR + CO Mapeado)")
st.caption("Versão completa com extração de Manifesto, Data/Hora, Valor Total, Volumes e COs mapeados")

responsavel = st.text_input("Responsável", value="Wandeilson Barros")
uploaded_files = st.file_uploader("Envie um ou mais manifestos (PDF)", type=["pdf"], accept_multiple_files=True)

dados = []
if uploaded_files:
    for f in uploaded_files:
        pdf_bytes = f.read()
        result, ocrF, ocrL = extrair_dados(pdf_bytes)
        result["responsavel"] = responsavel
        ok = all(result[k] for k in ["manifesto", "data", "destino"])
        if ok:
            st.success(f"✅ {f.name} | {result['destino']}")
        else:
            st.warning(f"⚠️ {f.name} - Dados incompletos")
        dados.append(result)

    if dados:
        st.subheader("📊 Prévia — MANIFESTOS")
        df = pd.DataFrame(dados)
        df.rename(columns={
            "manifesto": "MANIFESTO",
            "data": "DATA",
            "hora": "HORA",
            "destino": "DESTINO",
            "valor": "VALOR TOTAL (R$)",
            "volumes": "VOLUMES",
            "responsavel": "RESPONSÁVEL"
        }, inplace=True)
        st.dataframe(df, use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Manifestos")
        st.download_button("📥 Baixar Planilha Operacional", buffer.getvalue(), "planilha_manifestos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
