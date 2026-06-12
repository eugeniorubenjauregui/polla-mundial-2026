"""
Gran Polla Familiar & Amigos — FIFA World Cup 2026 🏆
Dashboard interactivo — Created by Souza Weich

INSTALACIÓN:   pip install streamlit pandas openpyxl plotly requests
EJECUCIÓN:     streamlit run app.py

GOOGLE SHEETS (auto-sync):
  1. Drive → Compartir → Cualquier persona con el enlace → Lector
  2. streamlit run app.py  (Google Sheets en vivo está activo por defecto)
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import numpy as np
from io import BytesIO
from datetime import datetime, timedelta

# ══════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Polla Mundial 2026 ⚽",
    page_icon="⚽",
    layout="wide",
    initial_sidebar_state="expanded",
)

GOOGLE_SHEETS_ID     = "1eqkt7qN7vthA8SMVWMYSKcsbDizKPiHIwsfGKVi3M0g"
GOOGLE_SHEETS_URL    = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEETS_ID}/export?format=xlsx"
POLL_SHEET_NAME      = "Respuestas de formulario 1"
LOCAL_FILE           = "Gran Polla Familiar y Amigoss - Mundial 2026 🏆 (respuestas).xlsx"
PRECIO_PP            = 20_000
REFRESH_SECONDS      = 300

# ── Banderas ──────────────────────────────────────────────
FLAGS = {
    "México":"🇲🇽","Sudáfrica":"🇿🇦","Corea del sur":"🇰🇷","Corea del Sur":"🇰🇷",
    "República Checa":"🇨🇿","Canadá":"🇨🇦","Bosnia Herzegovina":"🇧🇦",
    "Bosnia y Herzegovina":"🇧🇦","Estados Unidos":"🇺🇸","Paraguay":"🇵🇾",
    "Catar":"🇶🇦","Suiza":"🇨🇭","Brasil":"🇧🇷","Marruecos":"🇲🇦",
    "Haití":"🇭🇹","Escocia":"🏴󠁧󠁢󠁳󠁣󠁴󠁿","Australia":"🇦🇺","Turquía":"🇹🇷",
    "Alemania":"🇩🇪","Curazao":"🇨🇼","Países Bajos":"🇳🇱","Japón":"🇯🇵",
    "Costa de marfil":"🇨🇮","Costa de Marfil":"🇨🇮","Ecuador":"🇪🇨","Suecia":"🇸🇪",
    "Túnez":"🇹🇳","España":"🇪🇸","Cabo Verde":"🇨🇻","Bélgica":"🇧🇪",
    "Egipto":"🇪🇬","Arabia Saudita":"🇸🇦","Uruguay":"🇺🇾","Irán":"🇮🇷",
    "Iran":"🇮🇷","Nueva Zelanda":"🇳🇿","Francia":"🇫🇷","Senegal":"🇸🇳",
    "Irak":"🇮🇶","Noruega":"🇳🇴","Argentina":"🇦🇷","Argelia":"🇩🇿",
    "Austria":"🇦🇹","Jordania":"🇯🇴","Portugal":"🇵🇹","RD Congo":"🇨🇩",
    "Inglaterra":"🏴󠁧󠁢󠁥󠁮󠁧󠁿","Croacia":"🇭🇷","Ghana":"🇬🇭","Panamá":"🇵🇦",
    "Uzbekistán":"🇺🇿","Uzbekistan":"🇺🇿","Colombia":"🇨🇴","Haiti":"🇭🇹",
}

ALL_NATIONS = [
    ("🇲🇽","México"),("🇿🇦","Sudáfrica"),("🇰🇷","Corea"),("🇨🇿","Rep. Checa"),
    ("🇨🇦","Canadá"),("🇧🇦","Bosnia"),("🇺🇸","EE.UU."),("🇵🇾","Paraguay"),
    ("🇶🇦","Catar"),("🇨🇭","Suiza"),("🇧🇷","Brasil"),("🇲🇦","Marruecos"),
    ("🇭🇹","Haití"),("🏴󠁧󠁢󠁳󠁣󠁴󠁿","Escocia"),("🇦🇺","Australia"),("🇹🇷","Turquía"),
    ("🇩🇪","Alemania"),("🇨🇼","Curazao"),("🇳🇱","P. Bajos"),("🇯🇵","Japón"),
    ("🇨🇮","C. Marfil"),("🇪🇨","Ecuador"),("🇸🇪","Suecia"),("🇹🇳","Túnez"),
    ("🇪🇸","España"),("🇨🇻","Cabo Verde"),("🇧🇪","Bélgica"),("🇪🇬","Egipto"),
    ("🇸🇦","Arabia S."),("🇺🇾","Uruguay"),("🇮🇷","Irán"),("🇳🇿","N. Zelanda"),
    ("🇫🇷","Francia"),("🇸🇳","Senegal"),("🇮🇶","Irak"),("🇳🇴","Noruega"),
    ("🇦🇷","Argentina"),("🇩🇿","Argelia"),("🇦🇹","Austria"),("🇯🇴","Jordania"),
    ("🇵🇹","Portugal"),("🇨🇩","RD Congo"),("🏴󠁧󠁢󠁥󠁮󠁧󠁿","Inglaterra"),("🇭🇷","Croacia"),
    ("🇬🇭","Ghana"),("🇵🇦","Panamá"),("🇺🇿","Uzbekistán"),("🇨🇴","Colombia"),
]

def flag(name):
    if not name or (isinstance(name, float) and np.isnan(name)):
        return "🌍"
    return FLAGS.get(str(name).strip(), "🌍")

def safe_int(val, default=None):
    try:
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return default
        return int(float(val))
    except:
        return default

def classify_winner(pred_w, team_a, team_b):
    if not pred_w or str(pred_w) == "nan":
        return "unknown"
    pl = str(pred_w).lower()
    if "empate" in pl:
        return "draw"
    for w in [t for t in str(team_a).split() if len(t) > 3]:
        if w.lower() in pl:
            return "a"
    for w in [t for t in str(team_b).split() if len(t) > 3]:
        if w.lower() in pl:
            return "b"
    return "unknown"

# ══════════════════════════════════════════════════════════
# CSS ESPECTACULAR — FIFA 2026 × Panini
# ══════════════════════════════════════════════════════════
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;900&display=swap');

.stApp {
    background: #060d1a;
    font-family: 'Inter', sans-serif;
}

/* ── Geo shapes background ── */
.stApp::before {
    content: '';
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background:
        radial-gradient(circle at 15% 20%, rgba(230,57,70,.12) 0%, transparent 40%),
        radial-gradient(circle at 85% 15%, rgba(69,123,157,.15) 0%, transparent 40%),
        radial-gradient(circle at 50% 80%, rgba(42,157,143,.10) 0%, transparent 40%),
        radial-gradient(circle at 90% 70%, rgba(244,162,97,.08) 0%, transparent 35%),
        radial-gradient(circle at 10% 80%, rgba(123,47,190,.08) 0%, transparent 35%);
    pointer-events: none;
    z-index: 0;
}

header[data-testid="stHeader"] { background: transparent !important; }
.block-container { padding-top: 0 !important; position: relative; z-index: 1; }
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0d1929 0%, #060d1a 100%) !important;
    border-right: 1px solid rgba(255,255,255,0.06);
}
section[data-testid="stSidebar"] * { color: rgba(255,255,255,0.85) !important; }

/* ── Hero banner ── */
.hero {
    background: linear-gradient(135deg,
        #C1121F 0%, #E63946 15%,
        #1D3557 30%, #457B9D 45%,
        #2A9D8F 60%, #0F6E56 75%,
        #7B2FBE 90%, #553098 100%);
    border-radius: 20px;
    padding: 0;
    overflow: hidden;
    position: relative;
    margin-bottom: 20px;
}
.hero-inner {
    position: relative;
    z-index: 2;
    padding: 30px 24px 24px;
    text-align: center;
}
.hero::before {
    content: '';
    position: absolute; inset: 0;
    background:
        radial-gradient(ellipse at 20% 50%, rgba(255,255,255,.08) 0%, transparent 55%),
        radial-gradient(ellipse at 80% 50%, rgba(255,255,255,.06) 0%, transparent 55%),
        repeating-linear-gradient(45deg, transparent, transparent 20px,
            rgba(255,255,255,.02) 20px, rgba(255,255,255,.02) 21px);
    z-index: 1;
}

/* Geometric color blobs in header */
.hero-blobs {
    position: absolute; inset: 0; z-index: 1; overflow: hidden;
}
.blob {
    position: absolute;
    border-radius: 50%;
    opacity: .25;
    filter: blur(1px);
}
.b1 { width:90px;height:90px; background:#E63946; top:-20px; left:5%; }
.b2 { width:70px;height:70px; background:#F4A261; top:10px; left:15%; border-radius:30% 70% 70% 30%/30% 30% 70% 70%; }
.b3 { width:110px;height:110px; background:#2A9D8F; top:-30px; right:10%; }
.b4 { width:80px;height:80px; background:#E9C46A; bottom:-10px; right:20%; border-radius:60% 40% 40% 60%/60% 60% 40% 40%; }
.b5 { width:60px;height:60px; background:#7B2FBE; bottom:0; left:40%; }
.b6 { width:100px;height:100px; background:#457B9D; top:20px; left:50%; border-radius:40% 60% 60% 40%/40% 40% 60% 60%; }

.hero-trophy {
    font-size: 3.5rem;
    filter: drop-shadow(0 4px 12px rgba(0,0,0,.5));
    display: block;
    margin-bottom: 6px;
}
.hero-title {
    color: white !important;
    font-size: 2rem;
    font-weight: 900;
    letter-spacing: 1.5px;
    margin: 0;
    text-shadow: 0 2px 20px rgba(0,0,0,.6);
}
.hero-subtitle {
    color: rgba(255,255,255,.7);
    font-size: .85rem;
    letter-spacing: 4px;
    text-transform: uppercase;
    margin: 4px 0;
}
.hero-badge {
    display: inline-block;
    background: rgba(255,255,255,.15);
    border: 1px solid rgba(255,255,255,.3);
    border-radius: 20px;
    padding: 5px 18px;
    color: white;
    font-size: .85rem;
    font-weight: 700;
    letter-spacing: 2px;
    margin: 10px 4px 0;
}
.created-by {
    margin-top: 12px;
    font-size: .78rem;
    color: rgba(255,255,255,.5);
    letter-spacing: 3px;
    text-transform: uppercase;
}
.created-by span {
    color: rgba(255,255,255,.9);
    font-weight: 700;
    letter-spacing: 2px;
}

/* Nations grid */
.nations-wrap {
    display: flex;
    flex-wrap: wrap;
    gap: 5px;
    justify-content: center;
    margin: 14px 0 0;
}
.nc {
    background: rgba(255,255,255,.1);
    border: 1px solid rgba(255,255,255,.15);
    border-radius: 8px;
    padding: 4px 9px;
    font-size: .78rem;
    color: rgba(255,255,255,.85);
    cursor: default;
    transition: background .15s, transform .15s;
}
.nc:hover {
    background: rgba(255,255,255,.2);
    transform: translateY(-2px);
}

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(255,255,255,.04);
    border-radius: 14px;
    padding: 5px;
    gap: 5px;
    border: 1px solid rgba(255,255,255,.07);
}
.stTabs [data-baseweb="tab"] {
    border-radius: 10px;
    color: rgba(255,255,255,.5) !important;
    font-weight: 600;
    font-size: .9rem;
    padding: 9px 18px;
    background: transparent !important;
    border: none !important;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, rgba(230,57,70,.35), rgba(200,30,50,.2)) !important;
    color: white !important;
    box-shadow: 0 2px 14px rgba(230,57,70,.25);
}

/* Metrics */
div[data-testid="metric-container"] {
    background: rgba(255,255,255,.05) !important;
    border: 1px solid rgba(255,255,255,.08) !important;
    border-radius: 14px !important;
    padding: 14px !important;
    transition: transform .2s, box-shadow .2s;
}
div[data-testid="metric-container"]:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 28px rgba(0,0,0,.4);
}
div[data-testid="metric-container"] label { color: rgba(255,255,255,.5) !important; font-size: .75rem !important; }
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: white !important; font-size: 2rem !important; font-weight: 900 !important;
}

/* Selectbox */
.stSelectbox > div > div {
    background: rgba(255,255,255,.06) !important;
    border: 1px solid rgba(255,255,255,.12) !important;
    border-radius: 10px !important;
    color: white !important;
}

/* Headings */
h1,h2,h3,h4 { color: white !important; }
p, label { color: rgba(255,255,255,.8) !important; }

/* Tables */
.stDataFrame { border-radius: 12px !important; overflow: hidden; }

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #E63946, #C1121F) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
}
.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(230,57,70,.4) !important;
}

/* Divider */
hr { border-color: rgba(255,255,255,.08) !important; margin: 10px 0 !important; }

/* Scrollbar */
::-webkit-scrollbar { width: 5px; }
::-webkit-scrollbar-track { background: rgba(255,255,255,.02); }
::-webkit-scrollbar-thumb { background: rgba(255,255,255,.15); border-radius: 3px; }

/* Card glass */
.glass {
    background: rgba(255,255,255,.04);
    border: 1px solid rgba(255,255,255,.08);
    border-radius: 18px;
    padding: 20px;
    transition: transform .2s, box-shadow .2s;
}
.glass:hover { transform: translateY(-3px); box-shadow: 0 8px 32px rgba(0,0,0,.4); }

/* Match scoreboard */
.match-board {
    background: rgba(255,255,255,.04);
    border: 1px solid rgba(255,255,255,.1);
    border-radius: 18px;
    padding: 24px 20px;
    text-align: center;
    margin: 14px 0;
}

/* Pozo card in sidebar */
.pozo {
    background: linear-gradient(135deg, rgba(42,157,143,.2), rgba(15,110,86,.1));
    border: 1px solid rgba(42,157,143,.3);
    border-radius: 14px;
    padding: 16px;
    text-align: center;
    margin-bottom: 12px;
}
.pozo-val {
    font-size: 1.6rem;
    font-weight: 900;
    background: linear-gradient(135deg, #2A9D8F, #43e8d8);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}

/* Badge status */
.badge-j { background: rgba(42,157,143,.2); border: 1px solid #2A9D8F; color: #2A9D8F;
           border-radius: 20px; padding: 3px 12px; font-size: .75rem; font-weight: 700; display: inline-block; }
.badge-p { background: rgba(233,196,106,.2); border: 1px solid #E9C46A; color: #E9C46A;
           border-radius: 20px; padding: 3px 12px; font-size: .75rem; font-weight: 700; display: inline-block; }

/* Rank row */
.rrow {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 10px 14px;
    border-radius: 12px;
    margin-bottom: 4px;
    border: 1px solid transparent;
    transition: all .15s;
}
.rrow:hover { background: rgba(255,255,255,.04); border-color: rgba(255,255,255,.07); }

/* Pts pill */
.pill {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: .75rem;
    font-weight: 700;
}

/* Toggle */
.stToggle label { color: rgba(255,255,255,.7) !important; }

/* Checkbox */
.stCheckbox label { color: rgba(255,255,255,.7) !important; }

/* Source info */
.src { text-align: right; color: rgba(255,255,255,.25); font-size: .72rem; padding-bottom: 6px; }

.stAlert { border-radius: 10px !important; }
</style>
"""

# ══════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════

def read_polla_sheet(source):
    """Read the poll spreadsheet from bytes or a local path."""
    workbook = pd.ExcelFile(source)
    sheet_name = (
        POLL_SHEET_NAME
        if POLL_SHEET_NAME in workbook.sheet_names
        else workbook.sheet_names[0]
    )
    return pd.read_excel(workbook, sheet_name=sheet_name, header=None)


@st.cache_data(ttl=REFRESH_SECONDS, show_spinner="Sincronizando con Google Sheets…")
def load_data(use_gsheets=True):
    if use_gsheets:
        try:
            resp = requests.get(GOOGLE_SHEETS_URL, timeout=20)
            resp.raise_for_status()
            df = read_polla_sheet(BytesIO(resp.content))
            return df, "📡 Google Sheets (en vivo)", datetime.now()
        except Exception:
            pass
    try:
        df = read_polla_sheet(LOCAL_FILE)
        source_label = "📁 Archivo local (respaldo)"
        if use_gsheets:
            source_label = "📁 Archivo local (Google Sheets no disponible)"
        return df, source_label, datetime.now()
    except FileNotFoundError:
        return None, None, None


@st.fragment(run_every=timedelta(seconds=REFRESH_SECONDS))
def auto_refresh_from_gsheets():
    if not st.session_state.get("use_gsheets"):
        return
    load_data.clear()
    parse_data.clear()
    st.rerun()

@st.cache_data
def parse_data(df_raw):
    row0,row1,row2,row3,row4,row5 = [df_raw.iloc[i] for i in range(6)]
    matches = []
    for n, val in enumerate(row0):
        if pd.isna(val) or not str(val).strip():
            continue
        n3 = n + 3
        status_raw = str(row1.iloc[n3]).strip() if n3 < len(row1) and pd.notna(row1.iloc[n3]) else ""
        played = "jugado" in status_raw.lower()
        parts = str(val).split(" Vr ")
        ta = parts[0].strip()
        tb = parts[1].strip() if len(parts) > 1 else ""
        if played:
            ta_off = str(row2.iloc[n3]).strip() if pd.notna(row2.iloc[n3]) else ta
            tb_off = str(row4.iloc[n3]).strip() if pd.notna(row4.iloc[n3]) else tb
            ga = safe_int(row3.iloc[n3] if pd.notna(row3.iloc[n3]) else None)
            gb = safe_int(row5.iloc[n3] if pd.notna(row5.iloc[n3]) else None)
        else:
            ta_off, tb_off, ga, gb = ta, tb, None, None
        matches.append({"name": str(val).strip(), "col": n, "status": "Jugado" if played else "Sin Jugar",
                        "team_a": ta_off or ta, "team_b": tb_off or tb, "goals_a": ga, "goals_b": gb})
    participants = []
    for idx in range(7, len(df_raw)):
        row = df_raw.iloc[idx]
        name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if not name or name.lower() in ("nan","none",""): continue
        acc = 0.0
        try: acc = float(row.iloc[11]) if pd.notna(row.iloc[11]) else 0
        except: pass
        p = {"nombre": name, "acumulado": acc,
             "campeon": str(row.iloc[7]).strip() if pd.notna(row.iloc[7]) else "",
             "subcampeon": str(row.iloc[8]).strip() if pd.notna(row.iloc[8]) else "",
             "tercero": str(row.iloc[9]).strip() if pd.notna(row.iloc[9]) else "",
             "cuarto": str(row.iloc[10]).strip() if pd.notna(row.iloc[10]) else ""}
        for m in matches:
            n = m["col"]
            pa = safe_int(row.iloc[n] if n < len(row) and pd.notna(row.iloc[n]) else None)
            pb = safe_int(row.iloc[n+1] if n+1 < len(row) and pd.notna(row.iloc[n+1]) else None)
            raw_w = row.iloc[n+2] if n+2 < len(row) else None
            pw = str(raw_w).strip() if pd.notna(raw_w) else ""
            pts = safe_int(row.iloc[n+3] if n+3 < len(row) and pd.notna(row.iloc[n+3]) else None, default=0)
            p[f"pa_{n}"] = pa; p[f"pb_{n}"] = pb; p[f"pw_{n}"] = pw; p[f"pts_{n}"] = pts or 0
        participants.append(p)
    return matches, pd.DataFrame(participants)

# ══════════════════════════════════════════════════════════
# COMPONENTS
# ══════════════════════════════════════════════════════════

def render_hero():
    nations_html = "".join(f'<span class="nc">{f} {n}</span>' for f, n in ALL_NATIONS)
    st.markdown(f"""
    <div class="hero">
        <div class="hero-blobs">
            <div class="blob b1"></div><div class="blob b2"></div>
            <div class="blob b3"></div><div class="blob b4"></div>
            <div class="blob b5"></div><div class="blob b6"></div>
        </div>
        <div class="hero-inner">
            <span class="hero-trophy">⚽</span>
            <div class="hero-subtitle">FIFA WORLD CUP™</div>
            <h1 class="hero-title">GRAN POLLA FAMILIAR & AMIGOS</h1>
            <div style="display:inline-flex;gap:8px;flex-wrap:wrap;justify-content:center;margin-top:8px">
                <span class="hero-badge">2026</span>
                <span class="hero-badge">🇲🇽 · 🇺🇸 · 🇨🇦</span>
                <span class="hero-badge">11 Jun – 19 Jul</span>
            </div>
            <div class="nations-wrap">{nations_html}</div>
            <div class="created-by">Created by <span>Souza Weich</span> &nbsp;·&nbsp; Con ❤️ para la familia</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_sidebar(use_gsheets, n_participants, last_updated, source=None):
    with st.sidebar:
        st.markdown("""
        <div style="text-align:center;padding:18px 0 10px">
            <div style="font-size:3rem">⚽</div>
            <div style="font-weight:700;font-size:1.1rem;color:white">Polla 2026</div>
            <div style="color:rgba(255,255,255,.4);font-size:.75rem">Familiar & Amigos</div>
            <div style="color:rgba(255,255,255,.3);font-size:.65rem;margin-top:4px;letter-spacing:2px">Created by Souza Weich</div>
        </div>""", unsafe_allow_html=True)
        st.divider()
        pozo = n_participants * PRECIO_PP
        st.markdown(f"""
        <div class="pozo">
            <div style="color:rgba(255,255,255,.5);font-size:.65rem;letter-spacing:2px;text-transform:uppercase">💰 Pozo Estimado</div>
            <div class="pozo-val">${pozo:,}</div>
            <div style="font-size:.8rem;color:rgba(255,255,255,.5)">COP</div>
            <div style="font-size:.68rem;color:rgba(255,255,255,.3);margin-top:3px">{n_participants} participantes × $20.000</div>
        </div>""", unsafe_allow_html=True)
        st.divider()
        st.markdown("<div style='font-size:.68rem;letter-spacing:2px;text-transform:uppercase;color:rgba(255,255,255,.35)'>⚙️ Fuente de datos</div>", unsafe_allow_html=True)
        if use_gsheets and source and "local" in source.lower():
            st.warning("Google Sheets no disponible. Usando respaldo local.")
        new_use = st.toggle("📡 Google Sheets en vivo", value=use_gsheets)
        if new_use:
            st.markdown(f"""<div style="background:rgba(42,157,143,.1);border:1px solid rgba(42,157,143,.25);border-radius:8px;padding:10px;font-size:.75rem;color:rgba(255,255,255,.6);margin-top:6px">
            ✅ Sincronizado con Google Sheets<br>🔄 Auto-refresh: {REFRESH_SECONDS // 60} min<br>
            ⚠️ La hoja debe ser <b>pública</b></div>""", unsafe_allow_html=True)
        else:
            st.markdown("""<div style="background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);border-radius:8px;padding:10px;font-size:.75rem;color:rgba(255,255,255,.45);margin-top:6px">
            Modo offline: usa el Excel local si existe.</div>""", unsafe_allow_html=True)
        if st.button("🔄 Actualizar ahora", use_container_width=True):
            st.cache_data.clear(); st.rerun()
        if last_updated:
            st.markdown(f"<div style='font-size:.65rem;color:rgba(255,255,255,.2);margin-top:4px'>Actualizado: {last_updated.strftime('%d/%m/%Y %H:%M:%S')}</div>", unsafe_allow_html=True)
        st.divider()
        st.markdown("<div style='font-size:.68rem;letter-spacing:2px;text-transform:uppercase;color:rgba(255,255,255,.35)'>📋 Sistema de puntos</div>", unsafe_allow_html=True)
        st.markdown("""<div style="font-size:.82rem;line-height:2.1">
        🥇 <b style="color:#FFD700">3 pts</b> — Marcador exacto<br>
        🥈 <b style="color:#C0C0C0">1 pt</b> — Ganador correcto<br>
        ❌ <b style="color:rgba(255,255,255,.3)">0 pts</b> — Pronóstico fallido</div>""", unsafe_allow_html=True)
    return new_use


# ══════════════════════════════════════════════════════════
# TAB 1 — TABLA DE POSICIONES
# ══════════════════════════════════════════════════════════
def tab_leaderboard(df, matches):
    if df is None or df.empty:
        st.info("No hay datos de participantes.")
        return
    df_s = df.sort_values("acumulado", ascending=False).reset_index(drop=True)
    played  = [m for m in matches if m["status"] == "Jugado"]
    pending = [m for m in matches if m["status"] != "Jugado"]
    max_pts = int(df_s["acumulado"].max()) if not df_s.empty else 0
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("👥 Participantes", len(df))
    c2.metric("✅ Partidos jugados", f"{len(played)} / 72")
    c3.metric("⏳ Pendientes", len(pending))
    c4.metric("🏅 Puntaje más alto", f"{max_pts} pts")
    st.markdown("<br>", unsafe_allow_html=True)

    # Podio
    if len(df_s) >= 3:
        st.markdown("### 🏅 Podio")
        pc = st.columns(3)
        PODIUM = [(1,"🥈","#C0C0C0","1.8rem"),(0,"🥇","#FFD700","2.2rem"),(2,"🥉","#CD7F32","1.5rem")]
        for col,(ri,medal,color,fz) in zip(pc, PODIUM):
            row = df_s.iloc[ri]
            with col:
                st.markdown(f"""
                <div class="glass" style="border-color:{color}30;text-align:center">
                    <div style="font-size:{fz};font-weight:500;color:{color}">{medal}</div>
                    <div style="color:white;font-weight:700;font-size:.95rem;margin:7px 0">{row["nombre"]}</div>
                    <div style="font-size:1.9rem;font-weight:900;color:{color}">{int(row["acumulado"])}</div>
                    <div style="color:rgba(255,255,255,.35);font-size:.72rem">puntos</div>
                    <div style="color:rgba(255,255,255,.5);font-size:.78rem;margin-top:7px">{flag(row.get("campeon",""))} {row.get("campeon","—")}</div>
                </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns([1.1, 1])
    medals = {0:"🥇",1:"🥈",2:"🥉"}
    with left:
        st.markdown("### 📋 Clasificación completa")
        rows = []
        for i, row in df_s.iterrows():
            rows.append({"Pos": medals.get(i, str(i+1)),
                         "Participante": row["nombre"],
                         "Pts": int(row["acumulado"]),
                         "Campeón apostado": f"{flag(row.get('campeon',''))} {row.get('campeon','—')}"})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True,
                     column_config={"Pos": st.column_config.TextColumn("Pos", width="small"),
                                    "Pts": st.column_config.NumberColumn("Puntos ⭐", format="%d", width="small")})
    with right:
        st.markdown("### 📊 Top 10")
        top = df_s.head(10).copy()
        top["short"] = top["nombre"].apply(lambda x: x.split()[0])
        pts_vals = top["acumulado"].tolist()
        colors = ["#FFD700" if i==0 else "#C0C0C0" if i==1 else "#CD7F32" if i==2 else
                  "#E63946" if p==3 else "#2A9D8F" if p==2 else "#457B9D" if p==1 else "#888780"
                  for i,(p) in enumerate(pts_vals)]
        fig = go.Figure(go.Bar(x=top["short"], y=top["acumulado"],
                               marker_color=colors, marker_line_color="rgba(255,255,255,.1)",
                               marker_line_width=1, borderradius=5,
                               text=top["acumulado"].astype(int), textposition="outside",
                               textfont={"color":"white","size":11}))
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                          font={"color":"white"}, height=340,
                          xaxis={"gridcolor":"rgba(255,255,255,.04)","tickfont":{"size":10}},
                          yaxis={"gridcolor":"rgba(255,255,255,.08)"},
                          margin={"t":30,"b":10,"l":10,"r":10}, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.markdown("### 🌟 ¿A quién le apuesta la familia como campeón?")
    champ = df["campeon"].value_counts().reset_index()
    champ.columns = ["equipo","votos"]
    champ = champ[champ["equipo"].str.strip() != ""]
    champ["label"] = champ["equipo"].apply(lambda t: f"{flag(t)} {t}")
    if not champ.empty:
        c_pie, c_bar = st.columns(2)
        COLORS = ["#534AB7","#1D9E75","#D85A30","#7F77DD","#BA7517","#378ADD","#993556","#888780","#639922"]
        with c_pie:
            fig_pie = px.pie(champ, values="votos", names="label", hole=.42,
                             color_discrete_sequence=COLORS)
            fig_pie.update_traces(textposition="inside", textinfo="percent+label", textfont_size=11)
            fig_pie.update_layout(paper_bgcolor="rgba(0,0,0,0)", font={"color":"white"},
                                  legend={"bgcolor":"rgba(0,0,0,0)","font":{"size":11}},
                                  height=280, margin={"t":10,"b":10})
            st.plotly_chart(fig_pie, use_container_width=True)
        with c_bar:
            fig_cb = go.Figure(go.Bar(y=champ["label"], x=champ["votos"], orientation="h",
                                      marker_color=COLORS[:len(champ)],
                                      text=champ["votos"], textposition="inside",
                                      textfont={"color":"white","size":12}))
            fig_cb.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                                  font={"color":"white"},
                                  xaxis={"gridcolor":"rgba(255,255,255,.08)","title":"Votos"},
                                  yaxis={"gridcolor":"rgba(255,255,255,.04)"},
                                  height=280, margin={"t":10,"b":10,"l":10,"r":20}, showlegend=False)
            st.plotly_chart(fig_cb, use_container_width=True)


# ══════════════════════════════════════════════════════════
# TAB 2 — ANÁLISIS POR PARTIDO
# ══════════════════════════════════════════════════════════
def tab_match_analysis(df, matches):
    labels = [f"{'✅' if m['status']=='Jugado' else '⏳'} {flag(m['team_a'])} {m['team_a']} vs {flag(m['team_b'])} {m['team_b']}"
              for m in matches]
    sel = st.selectbox("Selecciona un partido:", labels, label_visibility="visible")
    idx = labels.index(sel)
    m = matches[idx]; n = m["col"]
    fa, fb = flag(m["team_a"]), flag(m["team_b"])

    st.markdown("<br>", unsafe_allow_html=True)
    if m["status"] == "Jugado":
        ga, gb = m["goals_a"], m["goals_b"]
        badge_html = '<span class="badge-j">✅ PARTIDO JUGADO</span>'
        score_str = f"{ga} – {gb}"
        w_text = f"Ganó {fa} {m['team_a']}" if ga > gb else (f"Ganó {fb} {m['team_b']}" if gb > ga else "🤝 Empate")
        w_col  = "#2A9D8F" if ga > gb else ("#E63946" if gb > ga else "#E9C46A")
    else:
        badge_html = '<span class="badge-p">⏳ PENDIENTE</span>'
        score_str = "? – ?"; w_text = "Partido aún no disputado"; w_col = "rgba(255,255,255,.35)"

    st.markdown(f"""
    <div class="match-board">
        <div style="margin-bottom:14px">{badge_html}</div>
        <div style="display:flex;justify-content:center;align-items:center;gap:28px;flex-wrap:wrap">
            <div style="text-align:center">
                <div style="font-size:2.5rem">{fa}</div>
                <div style="color:white;font-weight:700;font-size:1.05rem;margin-top:4px">{m["team_a"]}</div>
            </div>
            <div style="text-align:center">
                <div style="font-size:3rem;font-weight:900;color:white;letter-spacing:8px;line-height:1">{score_str}</div>
                <div style="color:{w_col};font-weight:600;font-size:.9rem;margin-top:7px">{w_text}</div>
            </div>
            <div style="text-align:center">
                <div style="font-size:2.5rem">{fb}</div>
                <div style="color:white;font-weight:700;font-size:1.05rem;margin-top:4px">{m["team_b"]}</div>
            </div>
        </div>
    </div>""", unsafe_allow_html=True)

    st.markdown("### 📊 ¿Cómo apostó la familia?")
    count_a = count_draw = count_b = 0
    for _, row in df.iterrows():
        cls = classify_winner(row.get(f"pw_{n}",""), m["team_a"], m["team_b"])
        if cls=="a": count_a+=1
        elif cls=="draw": count_draw+=1
        elif cls=="b": count_b+=1
    total = len(df); pct = lambda c: round(c/total*100) if total else 0

    mc1, mc2, mc3 = st.columns(3)
    for col, (label, count, color, emoji) in zip(
        [mc1,mc2,mc3],
        [(f"Gana {m['team_a']}",count_a,"#2A9D8F",fa),
         ("Empate",count_draw,"#E9C46A","🤝"),
         (f"Gana {m['team_b']}",count_b,"#E63946",fb)]):
        with col:
            st.markdown(f"""
            <div class="glass" style="text-align:center;border-color:{color}30">
                <div style="font-size:2rem">{emoji}</div>
                <div style="font-size:2rem;font-weight:900;color:{color}">{count}</div>
                <div style="color:rgba(255,255,255,.6);font-size:.82rem">{label}</div>
                <div style="color:rgba(255,255,255,.35);font-size:.75rem">{pct(count)}%</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    if count_a + count_draw + count_b > 0:
        fig = go.Figure(go.Bar(
            x=[f"{fa} {m['team_a']}", "🤝 Empate", f"{fb} {m['team_b']}"],
            y=[count_a, count_draw, count_b],
            marker_color=["#2A9D8F","#E9C46A","#E63946"],
            text=[count_a,count_draw,count_b], textposition="outside",
            textfont={"color":"white","size":14}))
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                          font={"color":"white"},
                          xaxis={"gridcolor":"rgba(255,255,255,.04)"},
                          yaxis={"gridcolor":"rgba(255,255,255,.08)","title":"Apuestas"},
                          height=260, margin={"t":30,"b":10}, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 🔍 Pronósticos individuales")
    rows = []
    for _, row in df.iterrows():
        pa = row.get(f"pa_{n}"); pb = row.get(f"pb_{n}"); pw = row.get(f"pw_{n}",""); pts = row.get(f"pts_{n}",0) or 0
        pronos = f"{int(pa) if pa is not None else '?'} – {int(pb) if pb is not None else '?'}"
        if m["status"] == "Jugado":
            res = "🥇 3 pts" if pts==3 else ("🥈 1 pt" if pts==1 else "❌ 0 pts")
        else:
            res = "⏳ Pendiente"
        rows.append({"Participante": row["nombre"],"Pronóstico": pronos,"Ganador apostado": str(pw),"Resultado": res})
    df_det = pd.DataFrame(rows)
    if m["status"] == "Jugado":
        order = {"🥇 3 pts":0,"🥈 1 pt":1,"❌ 0 pts":2,"⏳ Pendiente":3}
        df_det["_s"] = df_det["Resultado"].map(order).fillna(9)
        df_det = df_det.sort_values("_s").drop("_s",axis=1)
    st.dataframe(df_det, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════
# TAB 3 — CONSULTA INDIVIDUAL
# ══════════════════════════════════════════════════════════
def tab_individual(df, matches):
    if df is None or df.empty:
        st.info("No hay datos de participantes.")
        return
    nombres = sorted(df["nombre"].tolist())
    sel_name = st.selectbox("👤 Selecciona un participante:", nombres)
    if not sel_name: return
    p = df[df["nombre"] == sel_name].iloc[0]
    df_s = df.sort_values("acumulado", ascending=False).reset_index(drop=True)
    rank_pos = int(df_s[df_s["nombre"] == sel_name].index[0]) + 1
    medals = {1:"🥇",2:"🥈",3:"🥉"}
    rank_icon = medals.get(rank_pos, f"#{rank_pos}")
    played = [m for m in matches if m["status"] == "Jugado"]
    max_pts = len(played) * 3
    eff = round(int(p["acumulado"]) / max_pts * 100) if max_pts > 0 else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🏆 Puntos totales", int(p["acumulado"]))
    c2.metric("📊 Posición", rank_icon)
    c3.metric("🎯 Efectividad", f"{eff}%")
    c4.metric("✅ Partidos jugados", len(played))

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 🏆 Pronóstico de campeones")
    cc = st.columns(4)
    for col, (label, team) in zip(cc, [("🥇 Campeón",p.get("campeon","—")),("🥈 Sub-campeón",p.get("subcampeon","—")),("🥉 3er puesto",p.get("tercero","—")),("4️⃣ 4to puesto",p.get("cuarto","—"))]):
        f_icon = flag(team) if team and team not in ("—","") else "❓"
        with col:
            st.markdown(f"""
            <div class="glass" style="text-align:center;min-height:100px">
                <div style="font-size:1.8rem">{f_icon}</div>
                <div style="color:white;font-weight:700;font-size:.88rem;margin:5px 0">{team if team and team!='—' else '—'}</div>
                <div style="color:rgba(255,255,255,.4);font-size:.72rem">{label}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── NUEVA SECCIÓN: VER PRONÓSTICO POR PARTIDO ──────────────
    st.markdown("### 🔎 ¿Qué aposté en este partido?")
    st.caption("Selecciona cualquier partido de los 72 para ver tu pronóstico")

    match_opts = [f"{'✅' if m['status']=='Jugado' else '⏳'} {flag(m['team_a'])} {m['team_a']} vs {flag(m['team_b'])} {m['team_b']}" for m in matches]
    sel_match_label = st.selectbox("Partido:", match_opts, label_visibility="collapsed")
    mi = match_opts.index(sel_match_label)
    m = matches[mi]; n_col = m["col"]
    pa = p.get(f"pa_{n_col}"); pb = p.get(f"pb_{n_col}"); pw = p.get(f"pw_{n_col}",""); pts = p.get(f"pts_{n_col}",0) or 0

    fa2, fb2 = flag(m["team_a"]), flag(m["team_b"])
    pred_score = f"{int(pa) if pa is not None else '?'} – {int(pb) if pb is not None else '?'}"

    # Badge de puntos
    if m["status"] == "Jugado":
        pts_badge = f"🥇 <b style='color:#FFD700'>3 pts — marcador exacto</b>" if pts==3 else (
                    f"🥈 <b style='color:#C0C0C0'>1 pt — ganador correcto</b>" if pts==1 else
                    f"❌ <b style='color:rgba(255,255,255,.4)'>0 pts</b>")
        official_score = f"{m['goals_a']} – {m['goals_b']}"
        winner_off = (f"Ganó {fa2} {m['team_a']}" if m['goals_a'] > m['goals_b'] else
                      (f"Ganó {fb2} {m['team_b']}" if m['goals_b'] > m['goals_a'] else "🤝 Empate"))
        result_section = f"""
        <div style="margin-top:14px;padding-top:14px;border-top:.5px solid rgba(255,255,255,.1)">
            <div style="display:flex;align-items:center;justify-content:center;gap:20px;flex-wrap:wrap">
                <div style="font-size:.8rem;color:rgba(255,255,255,.5)">Resultado real:</div>
                <div style="font-size:1.5rem;font-weight:700;color:white">{official_score}</div>
                <div style="font-size:.85rem;color:rgba(255,255,255,.6)">{winner_off}</div>
                <div>{pts_badge}</div>
            </div>
        </div>"""
    else:
        result_section = '<div style="margin-top:12px;font-size:.8rem;color:rgba(255,255,255,.4)"><span class="badge-p">Partido aún no disputado</span></div>'

    st.markdown(f"""
    <div class="match-board">
        <div style="margin-bottom:12px">{'<span class="badge-j">✅ Jugado</span>' if m["status"]=="Jugado" else '<span class="badge-p">⏳ Pendiente</span>'}</div>
        <div style="font-size:.85rem;color:rgba(255,255,255,.6);margin-bottom:8px">{m['team_a']} vs {m['team_b']}</div>
        <div style="display:flex;justify-content:center;align-items:center;gap:24px;flex-wrap:wrap">
            <div style="text-align:center">
                <div style="font-size:2.2rem">{fa2}</div>
                <div style="font-size:.9rem;font-weight:600;color:white;margin-top:4px">{m['team_a']}</div>
            </div>
            <div style="text-align:center">
                <div style="font-size:.7rem;color:rgba(255,255,255,.4);letter-spacing:2px;text-transform:uppercase">Mi pronóstico</div>
                <div style="font-size:2.8rem;font-weight:900;color:white;letter-spacing:8px;line-height:1.1">{pred_score}</div>
                <div style="font-size:.82rem;color:rgba(255,255,255,.6);margin-top:4px">Aposté: <b style="color:white">{pw or '—'}</b></div>
            </div>
            <div style="text-align:center">
                <div style="font-size:2.2rem">{fb2}</div>
                <div style="font-size:.9rem;font-weight:600;color:white;margin-top:4px">{m['team_b']}</div>
            </div>
        </div>
        {result_section}
    </div>""", unsafe_allow_html=True)

    # ── Partidos jugados ──────────────────────────────────────
    if played:
        st.markdown("### 📋 Partidos jugados — pronóstico vs resultado real")
        fc1, fc2 = st.columns(2)
        with fc1: only_played = st.checkbox("Ver solo con puntos (≥ 1 pt)")
        with fc2: show_evo = st.checkbox("Mostrar evolución de puntos", value=True)

        bet_rows = []
        for m_p in played:
            np = m_p["col"]
            pa2 = p.get(f"pa_{np}"); pb2 = p.get(f"pb_{np}"); pw2 = p.get(f"pw_{np}",""); pts2 = p.get(f"pts_{np}",0) or 0
            fa3, fb3 = flag(m_p["team_a"]), flag(m_p["team_b"])
            pronos = f"{int(pa2) if pa2 is not None else '?'} – {int(pb2) if pb2 is not None else '?'}"
            oficial = f"{m_p['goals_a']} – {m_p['goals_b']}"
            pts_d = "🥇 3" if pts2==3 else ("🥈 1" if pts2==1 else "❌ 0")
            bet_rows.append({
                "Partido": f"{fa3} {m_p['team_a']} vs {fb3} {m_p['team_b']}",
                "Mi pronóstico": pronos,
                "Ganador apostado": str(pw2),
                "Resultado real": oficial,
                "Pts": pts_d})

        df_bets = pd.DataFrame(bet_rows)
        if only_played: df_bets = df_bets[df_bets["Pts"].isin(["🥇 3","🥈 1"])]
        st.dataframe(df_bets, use_container_width=True, hide_index=True)

        if show_evo:
            st.markdown("### 📈 Evolución de puntos")
            evo, cum = [], 0
            for m_p in played:
                pts_v = p.get(f"pts_{m_p['col']}", 0) or 0
                cum += pts_v
                evo.append({"Partido": m_p["team_a"].split()[0], "Pts ganados": pts_v, "Acumulado": cum})
            df_evo = pd.DataFrame(evo)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=df_evo["Partido"], y=df_evo["Acumulado"],
                                     mode="lines+markers+text", text=df_evo["Acumulado"],
                                     textposition="top center", textfont={"color":"white","size":11},
                                     line={"color":"#E63946","width":3},
                                     marker={"size":10,"color":"#FFD700","line":{"color":"#E63946","width":2}},
                                     fill="tozeroy", fillcolor="rgba(230,57,70,.08)"))
            fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                              font={"color":"white"},
                              xaxis={"gridcolor":"rgba(255,255,255,.04)","tickangle":-35},
                              yaxis={"gridcolor":"rgba(255,255,255,.08)","title":"Puntos acumulados"},
                              height=300, margin={"t":20,"b":60}, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════
def main():
    st.markdown(CSS, unsafe_allow_html=True)

    if "use_gsheets" not in st.session_state:
        st.session_state.use_gsheets = True

    df_raw, source, last_updated = load_data(use_gsheets=st.session_state.use_gsheets)

    if df_raw is None:
        st.error(
            "❌ No se pudo cargar la hoja de Google Sheets ni el archivo local. "
            f"Comparte la hoja como **Lector** o coloca `{LOCAL_FILE}` en esta carpeta."
        )
        return

    matches, df_p = parse_data(df_raw)

    new_use = render_sidebar(st.session_state.use_gsheets, len(df_p), last_updated, source)
    if new_use != st.session_state.use_gsheets:
        st.session_state.use_gsheets = new_use
        st.cache_data.clear(); st.rerun()

    render_hero()

    if source and last_updated:
        st.markdown(f"<div class='src'>{source} — {last_updated.strftime('%d/%m/%Y %H:%M:%S')}</div>", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs([
        "🏆   Tabla de posiciones",
        "⚽   Análisis por partido",
        "🔍   Consulta individual",
    ])
    with tab1: tab_leaderboard(df_p, matches)
    with tab2: tab_match_analysis(df_p, matches)
    with tab3: tab_individual(df_p, matches)

    auto_refresh_from_gsheets()

    st.markdown("""
    <div style="text-align:center;padding:28px 0 12px;color:rgba(255,255,255,.15);font-size:.7rem;letter-spacing:1.5px">
        ⚽ FIFA WORLD CUP 2026™ &nbsp;·&nbsp; Created by <b style="color:rgba(255,255,255,.3)">Souza Weich</b>
        &nbsp;·&nbsp; Resultados oficiales FIFA rigen
    </div>""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
