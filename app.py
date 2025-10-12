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
st.set_page_config(page_title="Leitor de Manifestos Jadlog — OCR Final", page_icon="🚛", layout="centered")
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

# ==========================
#  CONSTANTES / MAPAS
# ==========================
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

# ==========================
#  FUNÇÕES AUXILIARES
# ==========================
def normalize(s: str) -> str:
    s = unicodedata.normalize('NFKD', s or "")
    return ''.join(c for c in s if not unicodedata.combining(c))

def clean_money(s: str) -> str:
    if not s: return ""
    s2 = s.replace("R$", "").replace(" ", "")
    s2 = s2.replace(".", "").replace(",", ".")
    try:
        v = float(re.findall(r"-?\d+(?:\.\d+)?", s2)[0])
        return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return ""

def clean_int(s: str) -> str:
    return re.sub(r"[^\d]", "", s or "")

def pil_preprocess(img):
    """Cinza + autocontraste + nitidez leve + binarização simples."""
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    g = g.filter(ImageFilter.SHARPEN)
    g = g.point(lambda x: 255 if x > 180 else 0)
    return g

def ocr_image(img, psm=6):
    """Executa OCR e retorna texto em UPPER."""
    cfg = f"--psm {psm} --oem 3"
    return pytesseract.image_to_string(img, lang="por+eng", config=cfg).upper()

# ==========================
#  EXTRAÇÃO DE TEXTO
# ==========================
def read_pdf_text(file_bytes):
    pages_txt = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for p in pdf.pages:
            pages_txt.append(p.extract_text() or "")
    return pages_txt

def extract_data_hora_from_head(first_page_text):
    cab = normalize(first_page_text or "")
    m = re.search(r"\b(\d{1,2})\s+(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\s+(\d{4})\s+(\d{2}:\d{2}:\d{2})", cab, re.I)
    if not m: 
        return "",""
    meses = {"jan":"01","fev":"02","mar":"03","abr":"04","mai":"05","jun":"06","jul":"07","ago":"08","set":"09","out":"10","nov":"11","dez":"12"}
    dia, mes_txt, ano, hora = m.groups()
    return f"{int(dia):02d}/{meses.get(mes_txt.lower(),'')}/{ano}", hora

def extract_manifesto_destino_from_text(full_text):
    norm = normalize(full_text)
    m = re.search(r"\b(?:NUM[EÉ]RO|N[ºO])\s*:\s*(\d{8,})", norm, re.I)
    manifesto = m.group(1) if m else ""
    if not manifesto:
        m2 = re.search(r"\b(\d{10,15})\b", norm)
        manifesto = m2.group(1) if m2 else ""

    rota = re.search(r"\bS\d{1,2}\b", norm)
    if rota and rota.group(0).upper() in ROTA_CO_MAP:
        destino = ROTA_CO_MAP[rota.group(0).upper()]
    else:
        destino = ""
        matches = list(re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ ]+)\s*-\s*([A-Z]{2})", norm))
        for m in reversed(matches):
            cidade, uf = m.groups()
            if uf.upper() in UF_VALIDAS:
                destino = f"{cidade.strip().upper()} - {uf.upper()}"
                break
    return manifesto, destino

def ocr_page_bytes(file_bytes, page_index="first", dpi=500):
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
#  PROCESSAMENTO
# ==========================
def process_pdf(file_bytes, want_debug=False):
    out = {"manifesto":"","data":"","hora":"","destino":"","valor":"","volumes":"","debug":{}}

    pages_txt = read_pdf_text(file_bytes)
    if pages_txt:
        out["data"], out["hora"] = extract_data_hora_from_head(pages_txt[0])
        manifesto, destino = extract_manifesto_destino_from_text("\n".join(pages_txt))
        out["manifesto"], out["destino"] = manifesto, destino

    if not out["manifesto"]:
        img1, ocr1 = ocr_page_bytes(file_bytes, page_index="first", dpi=500)
        out["debug"]["OCR_1a_PAG"] = ocr1
        m = re.search(r"\b(?:NUM[EÉ]RO|N[ºO])\s*[:\-]?\s*(\d{8,})", ocr1, re.I)
        if not m: m = re.search(r"\b(\d{10,15})\b", ocr1)
        if m: out["manifesto"] = m.group(1)

    imgL, ocrL = ocr_page_bytes(file_bytes, page_index="last", dpi=500)
    out["debug"]["OCR_ULTIMA_PAG"] = ocrL

    # ---- CAPTURA INTELIGENTE DO VALOR TOTAL ----
    if not out["valor"]:
        # tenta pegar a linha inteira após "VALOR TOTAL DO MANIFESTO"
        linha_valor = ""
        for linha in ocrL.splitlines():
            if "VALOR TOTAL DO MANIFESTO" in linha:
                linha_valor = linha.strip()
                break
        if linha_valor:
            # captura o número com tudo (pontuação, espaço, vírgula)
            m_val = re.search(r"VALOR\s+TOTAL\s+DO\s+MANIFESTO\s*[:\-]?\s*([\d\s\.,]+)", linha_valor, re.I)
            if m_val:
                val_text = m_val.group(1)
                val_text = val_text.replace(" ", "").strip()

                # remove símbolos errados, mantém só números e separadores
                val_text = re.sub(r"[^0-9\.,]", "", val_text)

                # determina qual é o separador decimal (último símbolo)
                if "." in val_text and "," in val_text:
                    # define o último símbolo como decimal
                    if val_text.rfind(".") > val_text.rfind(","):
                        val_text = val_text.replace(",", "")
                    else:
                        val_text = val_text.replace(".", "").replace(",", ".")
                elif "," in val_text:
                    val_text = val_text.replace(".", "").replace(",", ".")
                else:
                    val_text = val_text.replace(",", "")

                try:
                    v = float(val_text)
                    out["valor"] = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                except:
                    out["valor"] = val_text


    if not out["volumes"]:
        m_vol = re.search(r"\bVOLUMES?\b\s*[:\-]?\s*([0-9]{1,6})\b", ocrL, re.I)
        if m_vol:
            out["volumes"] = clean_int(m_vol.group(1))

    if not out["destino"]:
        mds = list(re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ ]+)\s*-\s*([A-Z]{2})", ocrL))
        for m in reversed(mds):
            cidade, uf = m.groups()
            if uf.upper() in UF_VALIDAS:
                out["destino"] = f"{cidade.strip().upper()} - {uf.upper()}"
                break

    if want_debug:
        st.subheader("🔬 Debug do OCR (texto bruto)")
        with st.expander("OCR — 1ª Página (cabeçalho/manifesto)"):
            st.code(out["debug"].get("OCR_1a_PAG","(sem texto)"))
        with st.expander("OCR — Última Página (valor/volumes)"):
            st.code(out["debug"].get("OCR_ULTIMA_PAG","(sem texto)"))

    return out

# ==========================
#  INTERFACE
# ==========================
st.title("📦 Leitor de Manifestos Jadlog — OCR Final (v4.4)")
st.caption("Lê Manifesto, Data, Hora, Destino, Valor Total e Volumes — com OCR automático e correção de separadores.")

responsavel = st.text_input("Responsável", placeholder="Digite o nome completo")
want_debug = st.checkbox("Mostrar debug do OCR (texto bruto)")

files = st.file_uploader("Envie um ou mais PDFs de manifesto", type=["pdf"], accept_multiple_files=True)

if files:
    linhas = []
    for f in files:
        try:
            f_bytes = f.read()
            result = process_pdf(f_bytes, want_debug=want_debug)
            linhas.append({
                "MANIFESTO": result["manifesto"],
                "DATA": result["data"],
                "HORA": result["hora"],
                "DESTINO": result["destino"],
                "VALOR TOTAL (R$)": result["valor"],
                "VOLUMES": result["volumes"],
                "RESPONSÁVEL": (responsavel or "").upper()
            })
            ok = all([result["manifesto"], result["data"], result["hora"], result["destino"]])
            st.success(f"✅ {f.name} | {result['destino'] or 'Destino indefinido'}") if ok else st.warning(f"⚠️ {f.name} — faltou algum campo")
        except Exception as e:
            st.error(f"Erro ao processar {f.name}: {e}")

    df = pd.DataFrame(linhas, columns=["MANIFESTO","DATA","HORA","DESTINO","VALOR TOTAL (R$)","VOLUMES","RESPONSÁVEL"])
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
    st.info("Envie 1 ou mais PDFs de manifesto para extrair automaticamente.")

