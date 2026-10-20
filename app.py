import io
import os
import json
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image

st.set_page_config(
    page_title="Ball Committee — Seating Studio & Floor Plan",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@600;700&family=Inter:wght@400;500;600&display=swap');
    
    .stApp {
        background-color: #F8F9FA;
        font-family: 'Inter', sans-serif;
    }
    
    .gala-title {
        font-family: 'Cinzel', serif !important;
        letter-spacing: 4px;
        color: #111111;
        font-size: 1.6rem;
        margin: 0;
        text-transform: uppercase;
        text-align: center;
    }

    .stButton>button {
        background-color: #111111 !important;
        color: #FFFFFF !important;
        border: 1px solid #111111 !important;
        border-radius: 4px !important;
        font-size: 0.85rem !important;
        font-weight: 500 !important;
        padding: 4px 12px !important;
    }
</style>
""", unsafe_allow_html=True)

def render_header():
    for logo in ["ball_logo.png", "ball logo high res.jpg", "input_file_0.png"]:
        if os.path.exists(logo):
            st.image(logo, width=240)
            return
    st.markdown("<h1 class='gala-title'>BALL COMMITTEE</h1>", unsafe_allow_html=True)

col_l, col_m, col_r = st.columns([1.5, 1, 1.5])
with col_m:
    render_header()

st.markdown("<p style='text-align: center; letter-spacing: 2px; color: #6B7280; text-transform: uppercase; font-size: 0.75rem; margin-top: -8px; margin-bottom: 14px;'>Seating Studio & Hall Floor Plan Designer</p>", unsafe_allow_html=True)

# ==========================================
# PARSER LOGICA
# ==========================================
COMMITTEE_ALIASES = {
    "SexK": ["sexk", "sexkreation", "sex kreation", "sex-kreation", "sexie", "sex-ie"],
    "Nextstep": ["nextstep", "next step", "next-step"],
    "BEES": ["bees", "bee's"],
    "Jubel": ["jubel", "jubileum"],
    "JSA Board": ["jsa board", "board", "styrelsen", "the board"],
    "Spring Inspiration": ["spring inspiration", "spring inspiration committee", "spring", "si"],
    "Ball Committee": ["ball committee", "ball", "summer ball", "winter banquet", "summer ball committee"],
    "Case Academy": ["case academy", "case"],
    "JSA Masters": ["jsa masters", "masters"],
    "Sports Committee": ["sports committee", "sports"],
    "Marketing Committee": ["marketing committee", "marketing"],
    "Education Committee": ["education committee", "education"]
}

NON_DIETARY = {'nej', 'no', 'none', 'nothing', 'inga', 'inget', 'nope', '-', 'x', 'ingen', 'inte', '0', 'nan', '.', '..', 'no allergies'}

def clean_text(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    if s.lower() in NON_DIETARY or s.lower().startswith("no allerg"):
        return ""
    return s

def split_names(val):
    if not val or pd.isna(val): return "", ""
    s = str(val).strip()
    if not s or s.lower() in NON_DIETARY: return "", ""
    for delim in ['\n', ' och ', ' and ', '&', '/', '+']:
        s = s.replace(delim, ',')
    parts = [p.strip() for p in s.split(',') if p.strip()]
    if len(parts) >= 2: return parts[0], parts[1]
    if len(parts) == 1: return parts[0], ""
    return "", ""

def match_committee(val):
    if not val or pd.isna(val): return None
    raw = str(val).strip().lower()
    for canonical, variations in COMMITTEE_ALIASES.items():
        for var in variations:
            if raw == var or f" {var} " in f" {raw} ":
                return canonical
    return None

def parse_guestlist(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    valid_sheets = [s for s in xls.sheet_names if not any(x in s.lower() for x in ['översikt', 'sammanställning', 'summary', 'overview'])]
    if not valid_sheets: valid_sheets = [xls.sheet_names[-1]]

    dfs = [pd.read_excel(uploaded_file, sheet_name=s).dropna(how='all') for s in valid_sheets]
    df_raw = pd.concat(dfs, ignore_index=True)
    cols = [str(c).strip().lower() for c in df_raw.columns]

    def get_col(*keys):
        for orig, low in zip(df_raw.columns, cols):
            if any(k in low for k in keys):
                return orig
        return None

    c_name1 = get_col('name ticket 1', 'your name', 'full name (first & last name)', 'name ticket', 'full name')
    c_fn = get_col('förnamn', 'first name')
    c_ln = get_col('efternamn', 'last name')
    c_dt = get_col('försäljnignsdatum', 'datum', 'date')
    c_p1 = get_col('seating preference 1:', '2 people you want to sit next to', 'preference 1', 'seating preference 1')
    c_p2 = get_col('seating preference 2:', 'preference 2', 'seating preference 2')
    c_diet1 = get_col('dietary requirements', 'allergies/food preference', 'dietary', 'allergies')
    c_comm1 = get_col('committee table', 'committee')
    c_friend = get_col('second guest', 'your friends first and last name', 'friend')
    c_p2_1 = get_col('seating preference 1 second guest', '2 people your friend want')
    c_p2_2 = get_col('seating preference 2 second guest')
    c_diet2 = get_col('dietary requirements second guest', 'your friends allergies')

    attendees = []
    buyer_counter = {}

    for idx, row in df_raw.iterrows():
        fn = str(row.get(c_fn, '')).strip() if c_fn and pd.notna(row.get(c_fn)) else ''
        ln = str(row.get(c_ln, '')).strip() if c_ln and
