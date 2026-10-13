import io
import os
import json
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image

# ==========================================
# 1. PAGINA CONFIGURATIE & FAVICON
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
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@500;700&family=Montserrat:wght@300;400;600&display=swap');
    .stApp { background-color: #F8F9FA; font-family: 'Montserrat', sans-serif; }
    h1, h2, h3, .gala-title { font-family: 'Cinzel', serif !important; letter-spacing: 2px; color: #111111; text-align: center; }
    .block-container { padding-top: 1rem; padding-bottom: 1rem; max-width: 98% !important; }
</style>
""", unsafe_allow_html=True)

# Logo weergave
def render_ball_logo():
    if os.path.exists("ball_logo.png"):
        st.image("ball_logo.png", use_container_width=True)
    elif os.path.exists("ball logo high res.jpg"):
        st.image("ball logo high res.jpg", use_container_width=True)
    elif os.path.exists("input_file_0.png"):
        st.image("input_file_0.png", use_container_width=True)
    else:
        st.markdown("<h3 class='gala-title'>BALL COMMITTEE</h3>", unsafe_allow_html=True)

col_spacer_l, col_center, col_spacer_r = st.columns([2, 1.5, 2])
with col_center:
    render_ball_logo()

st.markdown("<p style='text-align: center; letter-spacing: 2px; color: #6B7280; text-transform: uppercase; font-size: 0.8rem; margin-top: -10px;'>20 Table Full-Overview Seating Studio</p>", unsafe_allow_html=True)

# ==========================================
# 2. INTELLIGENTE GASTENLIJST PARSER
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

NON_DIETARY = {'nej', 'no', 'none', 'nothing', 'inga', 'inget', 'nope', '-', 'x', 'ingen', 'inte', '0', 'nan', '.', 'no allergies'}

def clean_dietary_entry(d):
    if pd.isna(d): return ""
    s = str(d).strip
