import io
import os
import json
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import openpyxl
from PIL import Image

# ==========================================
# 1. PAGE CONFIGURATION & STYLING
# ==========================================
def get_page_icon():
    for name in ["ball_logo.png", "ball logo high res.jpg", "input_file_0.png"]:
        if os.path.exists(name):
            try:
                return Image.open(name)
            except Exception:
                pass
    return "🍾"

st.set_page_config(
    page_title="Ball Committee — Seating Studio",
    page_icon=get_page_icon(),
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@600;700&family=Montserrat:wght@400;500;600&display=swap');
    
    .stApp {
        background-color: #F8F9FA;
        font-family: 'Montserrat', sans-serif;
    }
    
    .gala-title {
        font-family: 'Cinzel', serif !important;
        letter-spacing: 3px;
        color: #111111;
        text-align: center;
        margin-bottom: 0px;
    }
    
    .stButton>button {
        background-color: #111111 !important;
        color: #FFFFFF !important;
        border: 1px solid #111111 !important;
        border-radius: 4px !important;
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 600 !important;
        letter-spacing: 1px;
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        background-color: #374151 !important;
        border-color: #374151 !important;
    }
    
    .upload-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        padding: 16px 20px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

def render_logo():
    for logo_path in ["ball_logo.png", "ball logo high res.jpg", "input_file_0.png"]:
        if os.path.exists(logo_path):
            st.image(logo_path, width=280)
            return
    st.markdown("<h2 class='gala-title'>BALL COMMITTEE</h2>", unsafe_allow_html=True)

col_spacer_l, col_center, col_spacer_r = st.columns([1.5, 1, 1.5])
with col_center:
    render_logo()

st.markdown("<p style='text-align: center; letter-spacing: 2px; color: #6B7280; text-transform: uppercase; font-size: 0.8rem; margin-top: -6px;'>Universal 20-Table Interactive Seating Studio</p>", unsafe_allow_html=True)

# ==========================================
# 2. PARSER & COMMITTEE ALIAS RESOLUTION
# ==========================================
COMMITTEE_ALIASES = {
    "SexK": ["sexk", "sexkreation", "sex kreation", "sex-kreation", "sexie", "sex-ie"],
    "Nextstep": ["nextstep", "next step", "next-step"],
    "BEES": ["bees", "bee's"],
    "Jubel": ["jubel", "jubileum"],
    "JSA Board": ["jsa board", "board", "styrelsen", "the board"],
    "Spring Inspiration": ["spring inspiration", "spring inspiration committee", "spring", "si"],
    "Ball Committee": ["ball committee", "ball", "summer ball", "winter banquet", "summer ball committee", "winter banquet committee"],
    "Case Academy": ["case academy", "case"],
    "JSA Masters": ["jsa masters", "masters"],
    "Sports Committee": ["sports committee", "sports", "sport committee"],
    "Marketing Committee": ["marketing committee", "marketing", "pr & marketing"],
    "Education Committee": ["education committee", "education"],
    "Quality Committee": ["quality committee", "quality"]
}

NON_DIETARY = {
    'nej', 'no', 'none', 'nothing', 'inga', 'inget', 'nope', '-', 'x', 'ingen', 
    'inte', '0', 'nan', '.', '..', 'no allergies', 'no allergy', 'ingenting', 
    'non', 'no allergen', 'inga allergier', 'no restriction', 'none!'
}

def clean_dietary_text(d):
    if pd.isna(d): return ""
    s = str(d).strip()
    if s.lower() in NON_DIETARY or s.lower().startswith('no allerg') or s.lower().startswith('inga allerg'):
        return ""
    return s

def split_pair_names(text):
    if not text or pd.isna(text): return "", ""
    s = str(text).strip()
    if not s or s.lower() in ['-', 'none', 'nej', 'no', 'x']: return "", ""
    for d in ['\n', ' och ', ' and ', '&', '/', '+']:
        s = s.replace(d, ',')
    parts = [p.strip() for p in s.split(',') if p.strip()]
    if len(parts) >= 2:
        return parts[0], parts[1]
    elif len(parts) == 1:
        return parts[0], ""
    return "", ""

def resolve_comm_alias(text):
    if not text or pd.isna(text): return None
    raw = str(text).strip().lower()
    if raw in ['-', 'none', 'nej', 'no', 'x', 'vet inte', 'ingen']: return None
    for canonical, variations in COMMITTEE_ALIASES.items():
        for var in variations:
            if raw == var or f" {var} " in f" {raw} ":
                return canonical
    return None

def parse_hitract_excel_dataset(file_obj):
    xls = pd.ExcelFile(file_obj)
    valid_sheets = []
    for s in xls.sheet_names:
        s_lower = s.lower().strip()
        if s_lower in ['översikt', 'sammanställning', 'summary', 'overview']:
            continue
        if any(w in s_lower for w in ['ticket', 'attendee', 'guest', 'banquet', '1)', '2)', '3)', '4)', '5)', 'seating']):
            valid_sheets.append(s)
            
    if not valid_sheets:
        valid_sheets = [xls.sheet_names[-1]]

    dfs = []
    for sheet_name in valid_sheets:
        try:
            df_s = pd.read_excel(file_obj, sheet_name=sheet_name).dropna(how='all')
            if len(df_s) > 0:
                dfs.append(df_s)
        except Exception:
            pass

    if not dfs:
        return []

    df_raw = pd.concat(dfs, ignore_index=True)
    cols = [str(c).strip().lower() for c in df_raw.columns]

    def find_col(*keywords):
        for orig_c, lower_c in zip(df_raw.columns, cols):
            if any(k in lower_c for k in keywords):
                return orig_c
        return None

    col_name1 = find_col('name ticket 1', 'your name')
    col_fname = find_col('förnamn', 'first name')
    col_lname = find_col('efternamn', 'last name')
    col_date = find_col('försäljnignsdatum', 'datum', 'date')
    col_pref1 = find_col('seating preference 1:', '2 people you want to sit next to', 'preference 1', 'seating preference 1')
    col_pref2 = find_col('seating preference 2:', 'preference 2', 'seating preference 2')
    col_diet1 = find_col('dietary requirements', 'allergies/food preference', 'dietary', 'allergies')
    col_comm1 = find_col('committee table', 'committee')
    col_friend = find_col('second guest', 'your friends first and last name', 'friend')
    col_pref2_1 = find_col('seating preference 1 second guest', '2 people your friend want')
    col_pref2_2 = find_col('seating preference 2 second guest')
    col_diet2 = find_col('dietary requirements second guest', 'your friends allergies')
    col_comm2 = find_col('committee table.1')

    # Clean 5-column support
    if len(df_raw.columns) <= 5 and not col_fname and 'name' in cols:
        name_c = find_col('name')
        comm_c = find_col('committee')
        p1_c = find_col('preference_1', 'pref1')
        p2_c = find_col('preference_2', 'pref2')
        diet_c = find_col('dietary', 'diet')
        attendees = []
        for idx, row in df_raw.iterrows():
            nm = str(row.get(name_c, '')).strip()
            if nm and nm.lower() != 'nan':
                attendees.append({
                    'id': f"g_{idx}",
                    'name': nm,
                    'committee': str(row.get(comm_c, '')).strip() if comm_c and pd.notna
