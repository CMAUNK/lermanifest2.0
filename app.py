import io
import re
import unicodedata
import pdfplumber
import pytesseract
from pdf2image import convert_from_bytes
import pandas as pd
import streamlit as st
from datetime import datetime

# ---------------------- CONFIGURAÇÃO STREAMLIT ----------------------
st.set_page_config(page_title="Leitor de Manifestos Jadlog", page_icon="🚛", layout="centered")

st.markdown("""
<style>
.stDownloadButton > button {
    background-color: #d62828;
    color: white;
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

# ---------------------- MAPA DE ROTAS ----------------------
ROTA_CO_MAP = {
    "S10":"CO GUAPIMIRIM","S12":"CO RIO DE JANEIRO 13","S16":"CO QUEIMADOS","S18":"CO SAO JOAO DE MERITI 01",
    "S20":"ML BELFORD ROXO 01","S21":"CO JUIZ DE FORA","S22":"DL DUQUE DE CAXIAS","S24":"CO ITABORAI",
    "S25":"CO VOLTA REDONDA","S27":"CO RIO DE JANEIRO 05","S28":"CO RIO DE JANEIRO 04","S30":"CO RIO DE JANEIRO 01",
    "S31":"CO JUIZ DE FORA","S32":"CO RIO DE JANEIRO 03","S37":"CO TRES RIOS","S38":"CO RIO DE JANEIRO 06",
    "S41":"CO RIO DE JANEIRO 08","S43":"CO ANGRA DOS REIS","S45":"CO PARATY","S48":"CO PETROPOLIS",
    "S49":"CO RIO DE JANEIRO 07","S54":"CO NOVA FRIBURGO","S56":"CO TERESOPOLIS","S58":"CO CAMPOS D. GOYTCAZES",
    "S59":"CO SANT. ANT DE PADUA","S60":"CO DUQUE DE CAXAIS","S61":"CO CABO FRIO","S63":"CO ARARUAMA",
    "S64":"CO SAO GONCALO","S65":"CO RIO DAS OSTRAS","S67":"CO NITEROI","S68":"CO MARICÁ","S70":"FL RIO DE JANEIRO",
}
UF_VALIDAS = {"AC","AL","AM","AP","BA","CE","DF","ES","GO","MA","MG","MS","MT","PA","PB","PE","PI","PR","RJ","RN","RO","RR","RS","SC","SE","SP","TO"}

# ---------------------- FUNÇÕES AUXILIARES ----------------------
def normalize(s: str) -> str:
    s = unicodedata.normalize('NFKD', s)
    return ''.join(c for c in s if not unicodedata.combining(c))

def clean_money(s: str) -> str:
    if not s: return ""
    s2 = s.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
    try:
        v = float(re.findall(r"-?\d+(?:\.\d+)?", s2)[0])
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return ""

def clean_int(s: str) -> str:
    return re.sub(r"[^\d]", "", s or "")

def read_pdf_text(file_like):
    pages = []
    file_like.seek(0)
    with pdfplumber.open(file_like) as pdf:
        for p in pdf.pages:
            pages.append(p.extract_text() or "")
    return pages

# ---------------------- EXTRAÇÕES ----------------------
def extract_data_hora(pages_text):
    cab = normalize(pages_text[0]) if pages_text else ""
    m = re.search(r"\b(\d{1,2})\s+(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\s+(\d{4})\s+(\d{2}:\d{2}:\d{2})", cab, re.I)
    if not m: return "", ""
    meses = {"jan":"01","fev":"02","mar":"03","abr":"04","mai":"05","jun":"06","jul":"07","ago":"08","set":"09","out":"10","nov":"11","dez":"12"}
    dia, mes_txt, ano, hora = m.groups()
    return f"{int(dia):02d}/{meses[mes_txt.lower()]}/{ano}", hora

def extract_manifesto_destino(texto):
    norm = normalize(texto)
    m = re.search(r"\bnumero\s*:\s*(\d{8,})", norm, re.I)
    manifesto = m.group(1) if m else ""
    rota = re.search(r"\bS\d{2}\b", norm)
    if rota and rota.group(0) in ROTA_CO_MAP:
        destino = ROTA_CO_MAP[rota.group(0)]
    else:
        destino = ""
        matches = list(re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ ]+)\s*-\s*([A-Z]{2})", norm, re.I))
        for m in reversed(matches):
            cidade, uf = m.groups()
            if uf.upper() in UF_VALIDAS:
                destino = f"{cidade.strip().upper()} - {uf.upper()}"
                break
    return manifesto, destino

def extract_valor_volume_ocr(file_like):
    """OCR apenas no rodapé para pegar Valor e Volume"""
    out = {"valor": "", "volumes": ""}
    pages = convert_from_bytes(file_like.read(), dpi=500, fmt='jpeg')
    if not pages: return out
    img = pages[-1]
    width, height = img.size
    crop_box = (0, height - 400, width, height)  # parte inferior
    cropped = img.crop(crop_box)
    ocr_text = pytesseract.image_to_string(cropped, lang="por+eng").upper()

    # Buscar padrões
    m_val = re.search(r"VALOR\s+TOTAL\s+DO\s+MANIFESTO[:\s]*([\d\.,]+)", ocr_text, re.I)
    if m_val: out["valor"] = clean_money(m_val.group(1))

    m_vol = re.search(r"VOLUMES[:\s]*([0-9]+)", ocr_text, re.I)
    if m_vol: out["volumes"] = clean_int(m_vol.group(1))

    return out

# ---------------------- INTERFACE ----------------------
st.title("📦 Leitor de Manifestos Jadlog (OCR Híbrido)")
st.caption("Extrai Manifesto, Data, Hora, Destino, Valor e Volumes (OCR localizado no rodapé).")

responsavel = st.text_input("Responsável", placeholder="Digite o nome completo")
arquivos = st.file_uploader("Envie um ou mais manifestos (PDF)", type=["pdf"], accept_multiple_files=True)

if arquivos:
    linhas = []
    for arq in arquivos:
        try:
            paginas = read_pdf_text(arq)
            data, hora = extract_data_hora(paginas)
            manifesto, destino = extract_manifesto_destino("\n".join(paginas))
            arq.seek(0)
            valor_vol = extract_valor_volume_ocr(arq)

            linhas.append({
                "MANIFESTO": manifesto,
                "DATA": data,
                "HORA": hora,
                "DESTINO": destino,
                "VALOR TOTAL (R$)": valor_vol["valor"],
                "VOLUMES": valor_vol["volumes"],
                "RESPONSÁVEL": (responsavel or "").upper()
            })
            st.success(f"✅ {arq.name} | {destino or 'Destino não identificado'}")
        except Exception as e:
            st.error(f"Erro ao processar {arq.name}: {e}")

    df = pd.DataFrame(linhas, columns=["MANIFESTO","DATA","HORA","DESTINO","VALOR TOTAL (R$)","VOLUMES","RESPONSÁVEL"])
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
    st.info("Envie um ou mais PDFs de manifesto para processar.")
