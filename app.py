# app.py
# -----------------------------------------------
# Leitor de Manifestos Jadlog — versão otimizada e corrigida
# - Extrai Manifesto, Data, Destino, Valor Total e Quantidade (Volumes)
# - Usa texto nativo do PDF quando possível (pdfplumber)
# - Faz OCR rápido (300 dpi) e fallback (500 dpi) quando necessário
# - Soma volumes de todas as páginas
# - Interface Streamlit com barra de progresso e botão limpar
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
def normalize(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in s if not unicodedata.combining(c))

def clean_int(s: str) -> str:
    return re.sub(r"[^\d]", "", s or "")

def pil_preprocess(img):
    """Pré-processamento para OCR: grayscale, autocontrast, sharpen, binarize leve."""
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    g = g.filter(ImageFilter.SHARPEN)
    # binarize leve: mantém contraste, evita perder pequenos detalhes
    g = g.point(lambda x: 255 if x > 180 else 0)
    return g

def ocr_image(img, psm=6):
    cfg = f"--psm {psm} --oem 3"
    return pytesseract.image_to_string(img, lang="por+eng", config=cfg).upper()

def ocr_page_quick(img):
    """Tenta PSM=6, se texto curto tenta PSM=3."""
    txt = ocr_image(img, psm=6)
    if len(txt.strip()) >= 10:
        return txt
    return ocr_image(img, psm=3)

def convert_pdf_fast(file_bytes: bytes, dpi=300):
    """Converte PDF para imagens em dpi especificado. Fallback 500 dpi se falhar."""
    try:
        images = convert_from_bytes(file_bytes, dpi=dpi, fmt="jpeg")
        if images:
            return images
    except Exception:
        pass
    # fallback
    return convert_from_bytes(file_bytes, dpi=500, fmt="jpeg")

# ==========================
#  EXTRAÇÃO DE TEXTO (texto nativo)
# ==========================
def read_pdf_text(file_bytes: bytes):
    pages_txt = []
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for p in pdf.pages:
                pages_txt.append(p.extract_text() or "")
    except Exception:
        # se pdfplumber falhar, retorna lista vazia
        return []
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
        "jan": "01","fev":"02","mar":"03","abr":"04","mai":"05","jun":"06",
        "jul":"07","ago":"08","set":"09","out":"10","nov":"11","dez":"12"
    }
    dia, mes_txt, ano, hora = m.groups()
    return f"{int(dia):02d}/{meses.get(mes_txt.lower(),'')}/{ano}", hora

# ==========================
#  EXTRAÇÃO — MANIFESTO + DESTINO
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

    # 1) Detecção robusta de Sxx (aceita S27, S 27, S-27, S/27, S027)
    rota_match = re.search(r"\bS\s*[-\/]?\s*0*([0-9]{1,2})\b", norm)
    if rota_match:
        rota_code = "S" + str(int(rota_match.group(1)))
        if rota_code in ROTA_CO_MAP:
            destino = ROTA_CO_MAP[rota_code]

    # 2) fallback: buscar diretamente cada chave do mapa
    if not destino:
        for k, v in ROTA_CO_MAP.items():
            pat = fr"S\s*[-\/]?\s*0*{k[1:]}"
            if re.search(pat, norm):
                destino = v
                break

    # 3) último recurso: "CIDADE - UF"
    if not destino:
        matches = list(re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ0-9 \.\-]+?)\s*-\s*([A-Z]{2})", norm))
        for m in reversed(matches):
            cidade, uf = m.groups()
            if uf.upper() in UF_VALIDAS:
                destino = f"{cidade.strip().upper()} - {uf.upper()}"
                break

    return manifesto, destino

# ==========================
#  HELPERS PARA VALORES
# ==========================
def find_monetary_candidates(text: str):
    """
    Retorna lista de strings que parecem valores monetários ou números com separadores.
    Ex: '26,475.41', '26.475,41', '26475.41', '26 475,41'
    """
    if not text:
        return []
    # busca padrões com grupos de milhares e decimais
    pattern = r"(?:(?:\d{1,3}(?:[.\s]\d{3})*(?:[,\.\s]\d{1,2})?)|\d+[\.,]\d+)"
    matches = re.findall(pattern, text)
    # dedupe e strip
    seen = []
    for m in matches:
        mm = m.strip()
        if mm and mm not in seen:
            seen.append(mm)
    return seen

def try_parse_amount(s: str):
    """
    Tenta converter string numérica em float (valor em unidades),
    tentando interpretar corretamente separadores.
    Retorna float ou None.
    """
    if not s:
        return None
    cur = s.strip().replace(" ", "")
    # remove non-digit, dot, comma
    cur = re.sub(r"[^\d\.,]", "", cur)
    # if both '.' and ',' present, decide by last occurrence
    if "." in cur and "," in cur:
        if cur.rfind(",") > cur.rfind("."):
            # comma likely decimal separator: remove dots (thousands), replace comma by dot
            try:
                norm = cur.replace(".", "").replace(",", ".")
                return float(norm)
            except:
                return None
        else:
            # dot likely decimal separator: remove commas
            try:
                norm = cur.replace(",", "")
                return float(norm)
            except:
                return None
    # only comma
    if "," in cur and "." not in cur:
        try:
            return float(cur.replace(".", "").replace(",", "."))
        except:
            return None
    # only dot
    if "." in cur and "," not in cur:
        try:
            return float(cur.replace(",", ""))
        except:
            return None
    # only digits
    try:
        return float(cur)
    except:
        return None

# ==========================
#  PROCESSAMENTO DE CADA PDF
# ==========================
def process_pdf(file_bytes: bytes, progress_bar=None, want_debug: bool = False):
    """
    Processa um PDF inteiro e retorna dicionário com campos:
    manifesto, data, destino, valor, volumes
    """
    out = {"manifesto": "", "data": "", "destino": "", "valor": "", "volumes": ""}

    # 1) Texto nativo (se disponível)
    pages_txt = read_pdf_text(file_bytes)
    if pages_txt:
        out["data"], _hora = extract_data_hora_from_head(pages_txt[0])
        manifesto, destino = extract_manifesto_destino_from_text("\n".join(pages_txt))
        out["manifesto"], out["destino"] = manifesto, destino

    # 2) Converter p/ imagens (300dpi rápido, fallback 500dpi)
    images = convert_pdf_fast(file_bytes, dpi=300)

    # 3) Extrair volumes somando todas páginas
    total_volumes = 0
    ocr_texts_per_page = []  # para debug/fallback do valor
    for idx, img in enumerate(images):
        if progress_bar:
            try:
                progress_bar.progress((idx + 1) / len(images))
            except Exception:
                pass

        img_p = pil_preprocess(img)
        txt = ocr_page_quick(img_p)
        ocr_texts_per_page.append(txt)

        # Primeiro busca explicitamente "VOLUMES: N" ou "Volumes N"
        m_vol = re.search(r"\bVOLUMES?\b\s*[:\-]?\s*([0-9]{1,6})\b", txt, re.I)
        if m_vol:
            try:
                total_volumes += int(m_vol.group(1))
                continue
            except Exception:
                pass

        # Em alguns modelos o final da página tem "31.13 74.00" -> pegar 74
        m_total_like = re.search(r"\b([0-9]{1,5})\s*\.\s*[0-9]{1,3}\b", txt)
        if m_total_like:
            try:
                total_volumes += int(m_total_like.group(1))
                continue
            except Exception:
                pass

        # Algumas linhas mostram "44 19.16 1" -> capturar primeiro número isolado
        m_first_num = re.search(r"^\s*([0-9]{1,4})\b", txt)
        if m_first_num:
            try:
                num = int(m_first_num.group(1))
                # heurística: se num plausível (<= 5000) e não estiver repetido com peso, adicionar
                if num > 0 and num < 5000:
                    total_volumes += num
                    continue
            except Exception:
                pass

    if total_volumes > 0:
        out["volumes"] = str(total_volumes)

    # 4) VALOR TOTAL — TENTAR PDFNATIVE (última página) PRIMEIRO, DEPOIS OCR
    valor_encontrado = ""

    # 4.1 tentar com pdfplumber (última página)
    last_text = ""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            try:
                last_text = pdf.pages[-1].extract_text() or ""
            except Exception:
                last_text = ""
    except Exception:
        last_text = ""

    # coleta candidatos do texto nativo
    candidates = []
    if last_text:
        # procurar explicitamente próximas ocorrências à expressão VALOR
        # pega linha inteira onde apareça VALOR
        for line in last_text.splitlines():
            if re.search(r"VALOR\s+TOTAL", line, re.I) or re.search(r"VALOR\s+TOTAL\s+DO", line, re.I) or re.search(r"VALOR\s*:", line, re.I):
                candidates += find_monetary_candidates(line)
        # se não achar via VALOR, coletar todos candidatos na última página
        if not candidates:
            candidates = find_monetary_candidates(last_text)

    # 4.2 fallback OCR última página (se não achou candidatos ou quiser reforçar)
    if not candidates:
        try:
            last_img = pil_preprocess(images[-1])
            ocr_last = ocr_page_quick(last_img)
            # prefer lines near "VALOR"
            for line in ocr_last.splitlines():
                if re.search(r"VALOR\s+TOTAL", line, re.I) or re.search(r"VALOR\s*:", line, re.I):
                    candidates += find_monetary_candidates(line)
            if not candidates:
                candidates = find_monetary_candidates(ocr_last)
        except Exception:
            candidates = []

    # 4.3 interpretar candidatos e escolher o melhor
    parsed = []
    for s in candidates:
        v = try_parse_amount(s)
        if v is not None:
            parsed.append((s, v))
    chosen_value = None
    if parsed:
        # priorizar candidatos extraídos de linhas que contenham "VALOR" (se existirem)
        valor_pref = []
        if last_text:
            for s, v in parsed:
                # if s appears in a line with 'VALOR' prefer it
                for line in last_text.splitlines():
                    if s in line and re.search(r"VALOR", line, re.I):
                        valor_pref.append((s, v))
            if valor_pref:
                # escolhe o maior entre preferenciais
                chosen_value = max(valor_pref, key=lambda x: x[1])[1]
        if chosen_value is None:
            # escolhe o maior valor plausível (normalmente o total é o maior número da página)
            chosen_value = max(parsed, key=lambda x: x[1])[1]

    # 4.4 formatar valor no padrão BR
    if chosen_value is not None:
        try:
            val_float = float(chosen_value)
            out["valor"] = f"{val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            # fallback: armazena o texto bruto do candidato escolhido
            out["valor"] = str(chosen_value)

    # 5) DESTINO fallback por OCR da última página (se ainda não detectado)
    if not out["destino"]:
        ocr_last_all = ""
        try:
            ocr_last_all = ocr_page_quick(pil_preprocess(images[-1]))
        except Exception:
            ocr_last_all = ""
        if ocr_last_all:
            mds = list(re.finditer(r"([A-ZÇÃÕÉÍÓÚÂÊÔÜ0-9 \.\-]+?)\s*-\s*([A-Z]{2})", ocr_last_all))
            for m in reversed(mds):
                cidade, uf = m.groups()
                if uf.upper() in UF_VALIDAS:
                    out["destino"] = f"{cidade.strip().upper()} - {uf.upper()}"
                    break

    # 6) debug (opcional)
    if want_debug:
        out["debug"] = {
            "pages_txt_count": len(pages_txt),
            "last_text_preview": last_text[:500] if last_text else "",
            "ocr_last_preview": ocr_texts_per_page[-1][:500] if ocr_texts_per_page else "",
            "candidates": candidates,
            "parsed": parsed,
        }

    return out

# ==========================
#  INTERFACE STREAMLIT
# ==========================
st.title("📦 Leitor de Manifestos Jadlog — OTIMIZADO")
st.caption("OCR rápido, soma de volumes em todas páginas, extração robusta do valor e UI melhorada.")

responsavel = st.text_input("Responsável", placeholder="Digite o nome completo")

if "linhas" not in st.session_state:
    st.session_state["linhas"] = []

if st.button("🧹 Limpar Tabela"):
    st.session_state["linhas"] = []

files = st.file_uploader("Envie PDFs de manifesto", type=["pdf"], accept_multiple_files=True)

if files:
    total_pdf = len(files)
    total_volumes_geral = 0

    for idx, f in enumerate(files):
        st.write(f"📄 Processando: **{f.name}** ({idx+1}/{total_pdf})")
        progress = st.progress(0)

        f_bytes = f.read()
        result = process_pdf(f_bytes, progress_bar=progress)

        st.session_state["linhas"].append({
            "Data": result.get("data", ""),
            "Manifesto": result.get("manifesto", ""),
            "Destino": result.get("destino", ""),
            "Referência": "",
            "Responsável": (responsavel or "").upper(),
            "Valor total": result.get("valor", ""),
            "Quantidade": result.get("volumes", ""),
        })

        try:
            total_volumes_geral += int(result.get("volumes") or 0)
        except Exception:
            pass

        st.success(f"✔️ {f.name} concluído — {result.get('volumes','0')} volumes")

    st.info(f"📊 Arquivos processados: {total_pdf} | Total de volumes: {total_volumes_geral}")

# Mostrar tabela
if st.session_state["linhas"]:
    df = pd.DataFrame(st.session_state["linhas"], columns=["Data", "Manifesto", "Destino", "Referência", "Responsável", "Valor total", "Quantidade"])
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
