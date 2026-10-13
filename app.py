import io
import os
import re
import json
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image

# ==========================================
# 1. PAGINA CONFIGURATIE & THEMA
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
    page_title="Ball Committee Seating Studio",
    page_icon=get_page_icon(),
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@500;700&family=Montserrat:wght@300;400;600&display=swap');
    .stApp { background-color: #F8F9FA; font-family: 'Montserrat', sans-serif; }
    h1, h2, h3, .gala-title { font-family: 'Cinzel', serif !important; letter-spacing: 2px; color: #111111; text-align: center; }
    
    /* Grote, duidelijke uploadbox */
    .stFileUploader {
        background: #FFFFFF;
        padding: 15px;
        border: 2px dashed #111111;
        border-radius: 8px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# Logo weergave bovenaan
def render_ball_logo():
    if os.path.exists("ball_logo.png"):
        st.image("ball_logo.png", width=320)
    elif os.path.exists("ball logo high res.jpg"):
        st.image("ball logo high res.jpg", width=320)
    elif os.path.exists("input_file_0.png"):
        st.image("input_file_0.png", width=320)
    else:
        st.markdown("<h2 class='gala-title'>BALL COMMITTEE</h2>", unsafe_allow_html=True)

col_l, col_m, col_r = st.columns([1, 1, 1])
with col_m:
    render_ball_logo()

st.markdown("<h3 style='text-align: center; font-family: Cinzel, serif; letter-spacing: 1px; margin-top: -10px;'>Seating Studio — 20 Tafels Overzicht</h3>", unsafe_allow_html=True)

# ==========================================
# 2. PROMINENTE UPLOAD BALK DIRECT ONDER LOGO
# ==========================================
st.markdown("---")
st.markdown("#### 📂 1. Upload hier je Excel gastenlijst")

col_upload, col_settings, col_reset = st.columns([3, 1, 1])

with col_upload:
    uploaded_file = st.file_uploader(
        "Kies je Hitract Excel bestand (.xlsx)", 
        type=["xlsx"], 
        key="main_guestlist_file"
    )

with col_settings:
    table_max_seats = st.number_input(
        "Stoelen per tafel", 
        min_value=4, 
        max_value=25, 
        value=10, 
        step=1
    )

with col_reset:
    st.write("")
    st.write("")
    if st.button("🔄 Reset / Wis Lijst", use_container_width=True):
        st.session_state.pop('active_guestlist', None)
        st.rerun()

st.markdown("---")

# ==========================================
# 3. UNIVERSELE GASTENLIJST PARSER
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
    "Sports Committee": ["sports committee", "sports", "sport committee"],
    "Marketing Committee": ["marketing committee", "marketing"],
    "Education Committee": ["education committee", "education"],
    "Quality Committee": ["quality committee", "quality"]
}

NON_DIETARY = {
    'nej', 'no', 'none', 'nothing', 'inga', 'inget', 'nope', '-', 'x', 'ingen', 
    'inte', '0', 'nan', '.', '..', 'no allergies', 'no allergy', 'ingenting', 
    'non', 'no allergen', 'inga allergier', 'no restriction', 'none!'
}

def clean_diet_val(d):
    if pd.isna(d): return ""
    s = str(d).strip()
    if s.lower() in NON_DIETARY or s.lower().startswith('no allerg'):
        return ""
    return s

def split_combined_names(text):
    if not text or pd.isna(text): return "", ""
    s = str(text).strip()
    if not s or s.lower() in ['-', 'none', 'nej', 'no', 'x']: return "", ""
    parts = re.split(r',|\boch\b|\band\b|&|/|\+|\n', s, flags=re.IGNORECASE)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) >= 2:
        return parts[0], parts[1]
    elif len(parts) == 1:
        return parts[0], ""
    return "", ""

def resolve_committee_alias(text):
    if not text or pd.isna(text): return None
    raw = str(text).strip().lower()
    for canonical, variations in COMMITTEE_ALIASES.items():
        for var in variations:
            if raw == var or f" {var} " in f" {raw} ":
                return canonical
    return None

def parse_any_guestlist_excel(file_bytes_or_buffer):
    xls = pd.ExcelFile(file_bytes_or_buffer)
    
    ticket_sheets = []
    for s in xls.sheet_names:
        s_lower = s.lower().strip()
        if s_lower in ['översikt', 'sammanställning', 'summary', 'overview']:
            continue
        if any(w in s_lower for w in ['ticket', 'attendee', 'guest', 'banquet', '1)', '2)', '3)', '4)', '5)', 'seating']):
            ticket_sheets.append(s)
    if not ticket_sheets:
        ticket_sheets = [xls.sheet_names[-1]]

    dfs = []
    for sheet_name in ticket_sheets:
        try:
            df_s = pd.read_excel(file_bytes_or_buffer, sheet_name=sheet_name).dropna(how='all')
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

    # Schone 5-koloms data check
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
                    'committee': str(row.get(comm_c, '')).strip() if comm_c and pd.notna(row.get(comm_c)) else "",
                    'pref1': str(row.get(p1_c, '')).strip() if p1_c and pd.notna(row.get(p1_c)) else "",
                    'pref2': str(row.get(p2_c, '')).strip() if p2_c and pd.notna(row.get(p2_c)) else "",
                    'dietary': clean_diet_val(row.get(diet_c, '')) if diet_c else ""
                })
        return attendees

    buyer_counts = {}
    attendees = []
    for idx, row in df_raw.iterrows():
        fn = str(row.get(col_fname, '')).strip() if col_fname and pd.notna(row.get(col_fname)) else ''
        ln = str(row.get(col_lname, '')).strip() if col_lname and pd.notna(row.get(col_lname)) else ''
        sdate = str(row.get(col_date, '')).strip() if col_date and pd.notna(row.get(col_date)) else ''
        b_key = f"{fn}_{ln}_{sdate}"
        occ = buyer_counts.get(b_key, 0) + 1
        buyer_counts[b_key] = occ

        n1_val = str(row.get(col_name1, '')).strip() if col_name1 and pd.notna(row.get(col_name1)) else ''
        n1 = n1_val if n1_val and n1_val.lower() != 'nan' else f"{fn} {ln}".strip()
        if not n1:
            continue

        n2_val = str(row.get(col_friend, '')).strip() if col_friend and pd.notna(row.get(col_friend)) else ''
        n2 = n2_val if n2_val and n2_val.lower() not in NON_DIETARY and len(n2_val) > 2 else ''
        if not n2 and occ > 1:
            n2 = f"Guest of {n1}"

        # Voorkeuren Gast 1
        raw_p1 = row.get(col_pref1, '') if col_pref1 else ''
        p1_a, p1_b = split_combined_names(raw_p1)
        if col_pref2 and pd.notna(row.get(col_pref2)):
            raw_p2 = str(row.get(col_pref2)).strip()
            if raw_p2 and raw_p2.lower() not in NON_DIETARY:
                p1_b = raw_p2

        diet1 = clean_diet_val(row.get(col_diet1, '')) if col_diet1 else ''
        comm1 = resolve_committee_alias(row.get(col_comm1, '')) if col_comm1 else resolve_committee_alias(raw_p1)

        # Voorkeuren Gast 2
        raw_p2_1 = row.get(col_pref2_1, '') if col_pref2_1 else ''
        p2_a, p2_b = split_combined_names(raw_p2_1)
        if col_pref2_2 and pd.notna(row.get(col_pref2_2)):
            raw_p2_2 = str(row.get(col_pref2_2)).strip()
            if raw_p2_2 and raw_p2_2.lower() not in NON_DIETARY:
                p2_b = raw_p2_2

        diet2 = clean_diet_val(row.get(col_diet2, '')) if col_diet
