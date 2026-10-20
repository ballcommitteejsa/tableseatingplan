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
        ln = str(row.get(c_ln, '')).strip() if c_ln and pd.notna(row.get(c_ln)) else ''
        dt = str(row.get(c_dt, '')).strip() if c_dt and pd.notna(row.get(c_dt)) else ''
        b_key = f"{fn}_{ln}_{dt}"
        occ = buyer_counter.get(b_key, 0) + 1
        buyer_counter[b_key] = occ

        n1_val = str(row.get(c_name1, '')).strip() if c_name1 and pd.notna(row.get(c_name1)) else ''
        n1 = n1_val if n1_val and n1_val.lower() != 'nan' else f"{fn} {ln}".strip()
        if not n1: continue

        n2_val = str(row.get(c_friend, '')).strip() if c_friend and pd.notna(row.get(c_friend)) else ''
        n2 = n2_val if n2_val and n2_val.lower() not in NON_DIETARY and len(n2_val) > 1 else ''
        if not n2 and occ > 1: n2 = f"Guest of {n1}"

        p1_a, p1_b = split_names(row.get(c_p1, '')) if c_p1 else ('', '')
        if c_p2 and pd.notna(row.get(c_p2)):
            val = str(row.get(c_p2)).strip()
            if val and val.lower() not in NON_DIETARY: p1_b = val

        diet1 = clean_text(row.get(c_diet1, '')) if c_diet1 else ''
        comm1 = match_committee(row.get(c_comm1, '')) if c_comm1 else match_committee(row.get(c_p1, ''))

        p2_a, p2_b = split_names(row.get(c_p2_1, '')) if c_p2_1 else ('', '')
        if c_p2_2 and pd.notna(row.get(c_p2_2)):
            val = str(row.get(c_p2_2)).strip()
            if val and val.lower() not in NON_DIETARY: p2_b = val

        diet2 = clean_text(row.get(c_diet2, '')) if c_diet2 else ''
        comm2 = match_committee(row.get(c_p2_1, ''))

        clean_pref1_a = "" if match_committee(p1_a) else p1_a
        clean_pref1_b = "" if match_committee(p1_b) else p1_b
        clean_pref2_a = "" if match_committee(p2_a) else p2_a
        clean_pref2_b = "" if match_committee(p2_b) else p2_b

        if occ == 1 or not n2:
            attendees.append({
                'id': f"g_{idx}_1",
                'name': n1,
                'committee': comm1 or "",
                'pref1': clean_pref1_a or p1_a,
                'pref2': clean_pref1_b or p1_b,
                'dietary': diet1
            })
        else:
            attendees.append({
                'id': f"g_{idx}_2",
                'name': n2,
                'committee': comm2 or "",
                'pref1': clean_pref2_a or (n1 if comm2 else p2_a),
                'pref2': clean_pref2_b or p2_b,
                'dietary': diet2
            })

    return attendees

# ==========================================
# UPLOAD BALK
# ==========================================
col_up1, col_up2 = st.columns([3, 1])

with col_up1:
    uploaded_file = st.file_uploader("Upload Hitract Guestlist (.xlsx)", type=["xlsx"], label_visibility="collapsed")

with col_up2:
    table_max_seats = st.number_input("Seats per table", min_value=4, max_value=25, value=10, step=1)

parsed_guests = []
if uploaded_file is not None:
    parsed_guests = parse_guestlist(uploaded_file)
    st.caption(f"Loaded {len(parsed_guests)} guests from {uploaded_file.name}")

guests_json = json.dumps(parsed_guests)

# ==========================================
# INTERACTIEVE STUDIO & FLOOR PLAN (HTML/JS)
# ==========================================
html_code = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    * {{ box-sizing: border-box; }}
    body {{
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        background: #F8F9FA;
        margin: 0;
        padding: 4px;
        user-select: none;
        color: #111;
    }}

    .toolbar {{
        background: #111;
        color: #FFF;
        padding: 8px 12px;
        border-radius: 4px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 8px;
        font-size: 0.8rem;
        flex-wrap: wrap;
        gap: 8px;
    }}
    .mode-tabs {{
        display: inline-flex;
        background: #27272A;
        padding: 2px;
        border-radius: 4px;
    }}
    .mode-tab {{
        background: transparent;
        color: #A1A1AA;
        border: none;
        padding: 5px 12px;
        border-radius: 3px;
        font-size: 0.75rem;
        font-weight: 600;
        cursor: pointer;
    }}
    .mode-tab.active {{
        background: #3F3F46;
        color: #FFF;
    }}

    .btn {{
        background: #374151;
        color: #FFF;
        border: none;
        padding: 5px 11px;
        border-radius: 3px;
        font-size: 0.75rem;
        cursor: pointer;
        font-weight: 500;
        margin-left: 4px;
        transition: background 0.15s;
    }}
    .btn:hover {{ background: #4B5563; }}
    .btn-blue {{ background: #2563EB; font-weight: 600; }}
    .btn-blue:hover {{ background: #1D4ED8; }}
    .btn-blue:disabled {{ background: #4B5563; opacity: 0.45; cursor: not-allowed; }}
    .btn-green {{ background: #15803D; }}
    .btn-green:hover {{ background: #166534; }}

    .action-bar-row {{
        background: #FFF;
        border: 1px solid #D1D5DB;
        border-radius: 4px;
        padding: 6px 10px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
        font-size: 0.75rem;
        flex-wrap: wrap;
    }}
    .action-bar-row input {{
        padding: 4px 6px;
        border: 1px solid #D1D5DB;
        border-radius: 3px;
        font-size: 0.75rem;
    }}

    .pool-section {{
        background: #FFF;
        border: 1px solid #D1D5DB;
        border-radius: 4px;
        margin-bottom: 12px;
        padding: 6px 8px;
    }}
    .pool-title {{
        font-weight: 700;
        font-size: 0.75rem;
        margin-bottom: 6px;
        color: #374151;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }}
    .pool-search-input {{
        padding: 3px 6px;
        border: 1px solid #D1D5DB;
        border-radius: 3px;
        font-size: 0.72rem;
        width: 220px;
    }}
    .pool-container {{
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        min-height: 50px;
        max-height: 200px;
        overflow-y: auto;
        padding: 4px;
        background: #F9FAFB;
        border: 1px dashed #D1D5DB;
        border-radius: 3px;
    }}

    /* GRID VIEW */
    .grid-dynamic {{
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 8px;
        width: 100%;
    }}
    .table-col {{
        background: #FFF;
        border: 1px solid #D1D5DB;
        border-top: 3px solid #111;
        border-radius: 3px;
        display: flex;
        flex-direction: column;
        min-height: 280px;
    }}
    .table-head {{
        padding: 4px 6px;
        background: #F3F4F6;
        border-bottom: 1px solid #E5E7EB;
        font-size: 0.75rem;
        font-weight: 700;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }}
    .table-input {{
        border: 1px dashed transparent;
        background: transparent;
        font-weight: 700;
        font-size: 0.75rem;
        width: 65%;
        padding: 1px 2px;
    }}
    .table-input:focus {{
        border-color: #2563EB;
        background: #FFF;
        outline: none;
    }}
    .tbl-del-btn {{
        background: transparent;
        border: none;
        color: #9CA3AF;
        cursor: pointer;
        font-size: 0.7rem;
        padding: 0 2px;
    }}
    .tbl-del-btn:hover {{
        color: #DC2626;
    }}
    .dropzone {{
        flex: 1;
        padding: 4px;
        min-height: 240px;
        display: flex;
        flex-direction: column;
        gap: 4px;
    }}
    .dropzone.dragover {{
        background: #ECFDF5;
        border: 1px dashed #059669;
    }}

    .guest-card {{
        background: #FFF;
        border: 1px solid #D1D5DB;
        border-left: 3px solid #2563EB;
        border-radius: 2px;
        padding: 4px 5px;
        cursor: grab;
        font-size: 0.72rem;
        display: flex;
        flex-direction: column;
    }}
    .guest-card.selected {{
        background: #EFF6FF !important;
        border-color: #2563EB !important;
    }}
    .guest-card.dragging {{
        opacity: 0.35;
    }}
    .guest-top {{
        display: flex;
        align-items: center;
        gap: 4px;
    }}
    .guest-name {{
        font-weight: 600;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        flex: 1;
        cursor: pointer;
    }}
    .guest-name:hover {{
        text-decoration: underline;
    }}
    .pref-text {{
        font-size: 0.65rem;
        color: #4B5563;
        margin-top: 2px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }}
    .tag-comm {{
        font-size: 0.6rem;
        font-weight: 600;
        color: #1E40AF;
        background: #DBEAFE;
        padding: 0 3px;
        border-radius: 2px;
    }}
    .diet-row {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-top: 2px;
    }}
    .tag-diet {{
        font-size: 0.6rem;
        font-weight: 600;
        color: #92400E;
        background: #FEF3C7;
        padding: 0 3px;
        border-radius: 2px;
        cursor: pointer;
    }}
    .tag-diet:hover {{
        background: #FDE68A;
    }}
    .add-diet-btn {{
        font-size: 0.58rem;
        color: #9CA3AF;
        cursor: pointer;
        padding: 0 2px;
    }}
    .add-diet-btn:hover {{
        color: #2563EB;
        text-decoration: underline;
    }}

    .bundle-card {{
        background: #FEF9C3;
        border: 1px solid #CA8A04;
        border-radius: 3px;
        padding: 4px 5px;
        cursor: grab;
        display: flex;
        flex-direction: column;
        gap: 2px;
        font-size: 0.72rem;
    }}
    .bundle-head {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-weight: 700;
        color: #854D0E;
    }}
    .ungroup-link {{
        background: #FFF;
        border: 1px solid #CA8A04;
        padding: 1px 4px;
        font-size: 0.65rem;
        cursor: pointer;
        border-radius: 2px;
    }}

    /* FLOOR PLAN VIEW */
    #floorplan-view {{
        display: none;
        width: 100%;
        margin-top: 4px;
    }}
    .hall-canvas {{
        width: 100%;
        height: 840px;
        background: #FFFFFF;
        border: 3px solid #111111;
        border-radius: 6px;
        position: relative;
        overflow: hidden;
        box-shadow: 0 4px 10px rgba(0,0,0,0.06);
    }}
    .hall-label {{
        position: absolute;
        top: 8px;
        left: 12px;
        font-size: 0.8rem;
        font-weight: 700;
        letter-spacing: 2px;
        color: #9CA3AF;
        text-transform: uppercase;
        pointer-events: none;
    }}
    .round-table-node {{
        width: 105px;
        height: 105px;
        border-radius: 50%;
        background: #FFFFFF;
        border: 3px solid #111111;
        position: absolute;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        text-align: center;
        cursor: move;
        box-shadow: 0 3px 6px rgba(0,0,0,0.12);
        padding: 6px;
        transition: border-color 0.15s, box-shadow 0.15s;
    }}
    .round-table-node:hover {{
        border-color: #2563EB;
        box-shadow: 0 4px 10px rgba(37,99,235,0.25);
    }}
    .rt-num {{
        font-weight: 800;
        font-size: 0.75rem;
        color: #111111;
    }}
    .rt-name {{
        font-size: 0.62rem;
        color: #4B5563;
        font-weight: 600;
        line-height: 1.1;
        max-width: 90px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }}
    .rt-badge {{
        font-size: 0.62rem;
        margin-top: 3px;
        font-weight: bold;
        padding: 1px 4px;
        border-radius: 3px;
    }}

    .hall-element-node {{
        position: absolute;
        background: #1F2937;
        color: #FFFFFF;
        border: 2px solid #111827;
        border-radius: 4px;
        padding: 6px 10px;
        font-size: 0.8rem;
        font-weight: 700;
        letter-spacing: 1px;
        cursor: move;
        box-shadow: 0 4px 8px rgba(0,0,0,0.18);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        align-items: center;
        text-align: center;
    }}
    .node-header {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        width: 100%;
        margin-bottom: 4px;
    }}
    .node-title {{
        flex: 1;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        font-size: 0.78rem;
    }}
    .node-controls {{
        display: flex;
        gap: 3px;
        align-items: center;
    }}
    .resize-btn {{
        background: #374151;
        color: #FFF;
        border: 1px solid #4B5563;
        border-radius: 2px;
        font-size: 0.65rem;
        cursor: pointer;
        padding: 1px 4px;
        font-weight: bold;
    }}
    .resize-btn:hover {{
        background: #4B5563;
    }}
    .del-btn {{
        background: transparent;
        color: #EF4444;
        border: none;
        cursor: pointer;
        font-weight: bold;
        padding: 0 3px;
        font-size: 0.85rem;
    }}
    .del-btn:hover {{
        color: #F87171;
    }}
</style>
</head>
<body>

<div class="toolbar">
    <div style="display: flex; align-items: center; gap: 8px;">
        <span style="font-weight: 600;">SEATING STUDIO</span>
        <div class="mode-tabs">
            <button class="mode-tab active" id="tab-grid" onclick="switchMode('grid')">Table Columns</button>
            <button class="mode-tab" id="tab-floor" onclick="switchMode('floor')">Hall Floor Plan (Zaalplattegrond)</button>
        </div>
    </div>
    <div>
        <span id="select-label" style="color: #9CA3AF; margin-right: 6px;">0 selected</span>
        <button class="btn btn-blue" id="groupBtn" onclick="bundleSelected()" disabled>Group Selected in Order</button>
        <button class="btn btn-green" onclick="exportCards()">Export Excel (Place Cards)</button>
        <button class="btn" onclick="exportPdf()">Save PDF (Tables & Allergies)</button>
        <button class="btn" onclick="resetLayout()">Reset Layout</button>
    </div>
</div>

<div class="action-bar-row">
    <div style="display: flex; align-items: center; gap: 6px;">
        <span style="font-weight: 600;">Add Guest:</span>
        <input type="text" id="manualName" placeholder="Full name (required)" style="width: 170px;">
        <input type="text" id="manualComm" placeholder="Committee" style="width: 100px;">
        <input type="text" id="manualDiet" placeholder="Dietary" style="width: 100px;">
        <button class="btn" onclick="addGuestManual()" style="margin: 0; background: #111;">Add Guest</button>
    </div>
    <div style="display: flex; align-items: center; gap: 6px;">
        <span style="font-weight: 600;">Manage Tables:</span>
        <button class="btn" onclick="addNewTable()" style="background: #2563EB;">+ Add Table</button>
        <span style="color: #6B7280; font-size: 0.72rem;">(Tip: Double click a name or dietary tag to edit it directly!)</span>
    </div>
</div>

<div class="pool-section" id="pool-section-wrapper">
    <div class="pool-title">
        <span>UNASSIGNED QUEUE (<span id="pool-count">0</span>)</span>
        <input type="text" class="pool-search-input" id="poolSearch" placeholder="Search unassigned guest..." oninput="filterPool(this.value)">
    </div>
    <div class="pool-container" id="pool-zone" ondragover="allowDrop(event)" ondrop="handleDrop(event, 'pool')" ondragleave="dragLeave(event)"></div>
</div>

<!-- MODE 1: GRID VIEW -->
<div class="grid-dynamic" id="tables-grid"></div>

<!-- MODE 2: FLOOR PLAN VIEW -->
<div id="floorplan-view">
    <div style="background: #FFF; border: 1px solid #D1D5DB; padding: 6px 10px; border-radius: 4px; margin-bottom: 8px; display: flex; align-items: center; gap: 8px; font-size: 0.75rem; flex-wrap: wrap;">
        <span style="font-weight: bold;">Add Stage / Hall Item:</span>
        <input type="text" id="itemText" placeholder="Name (e.g. Stage / Podium, DJ, Bar, Entrance)" style="padding: 4px 8px; border: 1px solid #D1D5DB; border-radius: 3px; font-size: 0.75rem; width: 250px;">
        <button class="btn" onclick="addHallItem()" style="background: #111; margin: 0;">Add to Hall</button>
        <button class="btn btn-blue" onclick="printFloorPlan()">Print Floor Plan</button>
        <span style="color: #6B7280; margin-left: auto;">Drag round tables & stage elements freely. Use [+] and [-] on stage to resize!</span>
    </div>
    <div class="hall-canvas" id="hall-canvas">
        <div class="hall-label">Ballroom Floor Plan • Drag & Position Tables</div>
    </div>
</div>

<script>
    const STORAGE_KEY = "BALL_STUDIO_STORAGE_DYNAMIC_V8";
    const maxCapacity = {table_max_seats};
    const rawGuests = {guests_json};

    let currentMode = 'grid';
    let guestMap = {{}};
    rawGuests.forEach(g => guestMap[g.id] = g);

    let selectedIds = new Set();
    let bundles = {{}};
    let activeTableIds = [];
    let tableNames = {{}};
    let tablePositions = {{}};
    let hallItems = [];
    let draggedId = null;

    // Default 20 tafels initialiseren
    for (let i = 1; i <= 20; i++) {{
        activeTableIds.push(i);
        tableNames[i] = "TABLE " + (i < 10 ? '0' + i : i);
        let colIdx = (i - 1) % 5;
        let rowIdx = Math.floor((i - 1) / 5);
        tablePositions[i] = {{ x: 40 + colIdx * 210, y: 70 + rowIdx * 180 }};
    }}

    function switchMode(mode) {{
        currentMode = mode;
        const gridView = document.getElementById('tables-grid');
        const floorView = document.getElementById('floorplan-view');
        const tabGrid = document.getElementById('tab-grid');
        const tabFloor = document.getElementById('tab-floor');

        if (mode === 'floor') {{
            gridView.style.display = 'none';
            floorView.style.display = 'block';
            tabGrid.classList.remove('active');
            tabFloor.classList.add('active');
            renderFloorPlan();
        }} else {{
            gridView.style.display = 'grid';
            floorView.style.display = 'none';
            tabGrid.classList.add('active');
            tabFloor.classList.remove('active');
        }}
    }}

    // ==========================================
    // DYNAMIC TABLE MANAGEMENT
    // ==========================================
    function rebuildTableGrid() {{
        const grid = document.getElementById('tables-grid');
        grid.innerHTML = '';
        activeTableIds.forEach(i => {{
            const col = document.createElement('div');
            col.className = 'table-col';
            col.id = 'table-column-wrapper-' + i;
            col.innerHTML = `
                <div class="table-head">
                    <input type="text" class="table-input" id="tbl-input-${{i}}" value="${{tableNames[i] || ('TABLE ' + i)}}" onchange="setTableName(${{i}}, this.value)">
                    <div style="display: flex; align-items: center; gap: 4px;">
                        <span id="t-count-${{i}}" style="color: #6B7280;">0/${{maxCapacity}}</span>
                        <button class="tbl-del-btn" onclick="removeTable(${{i}})" title="Delete this table">✕</button>
                    </div>
                </div>
                <div class="dropzone" id="zone-table-${{i}}" ondragover="handleDragOverZone(event)" ondrop="handleDrop(event, 'zone-table-${{i}}')" ondragleave="dragLeave(event)"></div>
            `;
            grid.appendChild(col);
        }});
        updateCounters();
    }}

    function addNewTable() {{
        let nextId = 1;
        while (activeTableIds.includes(nextId)) {{
            nextId++;
        }}
        activeTableIds.push(nextId);
        activeTableIds.sort((a, b) => a - b);
        tableNames[nextId] = "TABLE " + (nextId < 10 ? '0' + nextId : nextId);
        
        let colIdx = (nextId - 1) % 5;
        let rowIdx = Math.floor((nextId - 1) / 5);
        tablePositions[nextId] = {{ x: 40 + colIdx * 210, y: 70 + rowIdx * 180 }};

        rebuildTableGrid();
        saveState();
        renderFloorPlan();
    }}

    function removeTable(tableId) {{
        const z = document.getElementById('zone-table-' + tableId);
        if (z && z.children.length > 0) {{
            if (!confirm(`Table ${{tableNames[tableId]}} has guests assigned! Remove table and move guests back to unassigned queue?`)) {{
                return;
            }}
            Array.from(z.children).forEach(child => {{
                poolZone.appendChild(child);
            }});
        }}
        activeTableIds = activeTableIds.filter(id => id !== tableId);
        delete tableNames[tableId];
        delete tablePositions[tableId];

        rebuildTableGrid();
        saveState();
        renderFloorPlan();
    }}

    function setTableName(i, val) {{
        tableNames[i] = val.trim() || ("TABLE " + (i < 10 ? '0' + i : i));
        const inp = document.getElementById('tbl-input-' + i);
        if (inp) inp.value = tableNames[i];
        saveState();
        renderFloorPlan();
    }}

    // ==========================================
    // NAAM & DIET BEWERKEN VIA DUBBELKLIK
    // ==========================================
    function editGuestName(guestId) {{
        const g = guestMap[guestId];
        if (!g) return;
        const newName = prompt("Edit guest name:", g.name);
        if (newName && newName.trim() && newName.trim() !== g.name) {{
            g.name = newName.trim();
            const cardEl = document.getElementById(guestId);
            if (cardEl) {{
                const nameSpan = cardEl.querySelector('.guest-name');
                if (nameSpan) {{
                    nameSpan.innerText = g.name;
                    nameSpan.title = "Double click to edit: " + g.name;
                }}
            }}
            saveState();
        }}
    }}

    function editGuestDiet(guestId) {{
        const g = guestMap[guestId];
        if (!g) return;
        const currentDiet = g.dietary || "";
        const newDiet = prompt("Edit dietary requirements / allergies (leave blank if none):", currentDiet);
        if (newDiet !== null) {{
            g.dietary = newDiet.trim();
            const cardEl = document.getElementById(guestId);
            if (cardEl) {{
                const dietRow = cardEl.querySelector('.diet-row');
                if (dietRow) {{
                    if (g.dietary) {{
                        dietRow.innerHTML = `<span class="tag-diet" onclick="editGuestDiet('${{g.id}}')" title="Click to edit allergy">${{g.dietary}}</span>`;
                    }} else {{
                        dietRow.innerHTML = `<span class="add-diet-btn" onclick="editGuestDiet('${{g.id}}')">+ add allergy</span>`;
                    }}
                }}
            }}
            saveState();
        }}
    }}

    // ==========================================
    // CARD CREATION
    // ==========================================
    function createGuestCard(g) {{
        const div = document.createElement('div');
        div.className = 'guest-card';
        div.id = g.id;
        div.draggable = true;
        div.ondragstart = (e) => dragStart(e, g.id);
        div.ondragend = (e) => dragEnd(e);

        let comm = g.committee ? `<span class="tag-comm">${{g.committee}}</span>` : '';

        let prefs = [];
        if (g.pref1) prefs.push(g.pref1);
        if (g.pref2) prefs.push(g.pref2);
        let prefText = prefs.length > 0 ? `<div class="pref-text">wants: ${{prefs.join(' & ')}}</div>` : '';

        let dietMarkup = g.dietary 
            ? `<span class="tag-diet" onclick="editGuestDiet('${{g.id}}')" title="Click to edit allergy">${{g.dietary}}</span>`
            : `<span class="add-diet-btn" onclick="editGuestDiet('${{g.id}}')">+ add allergy</span>`;

        div.innerHTML = `
            <div class="guest-top">
                <input type="checkbox" onchange="toggleSelect(event, '${{g.id}}')" style="cursor: pointer; margin: 0 4px 0 0;">
                <span class="guest-name" title="Double click to edit name: ${{g.name}}" ondblclick="editGuestName('${{g.id}}')">${{g.name}}</span>
                ${{comm}}
            </div>
            ${{prefText}}
            <div class="diet-row">
                ${{dietMarkup}}
            </div>
        `;
        return div;
    }}

    function createBundleCard(bId, bundleData) {{
        const div = document.createElement('div');
        div.className = 'bundle-card';
        div.id = bId;
        div.draggable = true;
        div.ondragstart = (e) => dragStart(e, bId);
        div.ondragend = (e) => dragEnd(e);

        let memberIds = bundleData.members || [];
        let labelName = bundleData.name || ("GROUP (" + memberIds.length + ")");
        let names = memberIds.map(id => guestMap[id] ? guestMap[id].name : id).join(', ');

        div.innerHTML = `
            <div class="bundle-head">
                <span>${{labelName}}</span>
                <button class="ungroup-link" onclick="ungroup('${{bId}}')">Ungroup</button>
            </div>
            <div style="font-weight: 500; font-size: 0.68rem; color: #713F12;">${{names}}</div>
        `;
        return div;
    }}

    const poolZone = document.getElementById('pool-zone');

    function filterPool(query) {{
        const q = query.toLowerCase().trim();
        Array.from(poolZone.children).forEach(child => {{
            const text = child.innerText.toLowerCase();
            if (!q || text.includes(q)) {{
                child.style.display = '';
            }} else {{
                child.style.display = 'none';
            }}
        }});
    }}

    // ==========================================
    // PRECISE INTRA-TABLE DRAG & DROP REORDERING
    // ==========================================
    function dragStart(e, id) {{
        draggedId = id;
        e.dataTransfer.setData('text/plain', id);
        setTimeout(() => {{
            const el = document.getElementById(id);
            if (el) el.classList.add('dragging');
        }}, 0);
    }}

    function dragEnd(e) {{
        if (draggedId) {{
            const el = document.getElementById(draggedId);
            if (el) el.classList.remove('dragging');
        }}
    }}

    function allowDrop(e) {{
        e.preventDefault();
        e.currentTarget.classList.add('dragover');
    }}

    function dragLeave(e) {{
        e.currentTarget.classList.remove('dragover');
    }}

    function handleDragOverZone(e) {{
        e.preventDefault();
        const dropzone = e.currentTarget;
        dropzone.classList.add('dragover');
        if (!draggedId) return;

        const draggingEl = document.getElementById(draggedId);
        if (!draggingEl) return;

        const afterElement = getDragAfterElement(dropzone, e.clientY);
        if (afterElement == null) {{
            dropzone.appendChild(draggingEl);
        }} else {{
            dropzone.insertBefore(draggingEl, afterElement);
        }}
    }}

    function getDragAfterElement(container, y) {{
        const draggableElements = [...container.querySelectorAll('.guest-card:not(.dragging), .bundle-card:not(.dragging)')];
        return draggableElements.reduce((closest, child) => {{
            const box = child.getBoundingClientRect();
            const offset = y - box.top - box.height / 2;
            if (offset < 0 && offset > closest.offset) {{
                return {{ offset: offset, element: child }};
            }} else {{
                return closest;
            }}
        }}, {{ offset: Number.NEGATIVE_INFINITY }}).element;
    }}

    function handleDrop(e, targetZoneId) {{
        e.preventDefault();
        e.currentTarget.classList.remove('dragover');
        if (!draggedId) return;
        const el = document.getElementById(draggedId);
        if (el) {{
            el.classList.remove('dragging');
            updateCounters();
            saveState();
            renderFloorPlan();
        }}
    }}

    // ==========================================
    // GROEPEREN MET BEHOUD VAN VISUELE VOLGORDE
    // ==========================================
    function toggleSelect(e, id) {{
        e.stopPropagation();
        if (e.target.checked) selectedIds.add(id);
        else selectedIds.delete(id);

        const cnt = selectedIds.size;
        document.getElementById('select-label').innerText = `${{cnt}} selected`;
        document.getElementById('groupBtn').disabled = cnt < 2;
        document.getElementById('groupBtn').innerText = cnt >= 2 ? `Group ${{cnt}} in Exact Order` : 'Group Selected in Order';
    }}

    function bundleSelected() {{
        if (selectedIds.size < 2) return;

        const firstId = Array.from(selectedIds)[0];
        const firstEl = document.getElementById(firstId);
        const parent = firstEl ? firstEl.parentElement : poolZone;

        let mIds = [];
        if (parent) {{
            Array.from(parent.children).forEach(child => {{
                if (selectedIds.has(child.id)) {{
                    mIds.push(child.id);
                }}
            }});
        }}
        selectedIds.forEach(id => {{
            if (!mIds.includes(id)) mIds.push(id);
        }});

        const bId = 'bundle_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
        bundles[bId] = {{ name: 'Group (' + mIds.length + ')', members: mIds }};

        mIds.forEach(id => {{
            const el = document.getElementById(id);
            if (el) el.remove();
        }});

        parent.appendChild(createBundleCard(bId, bundles[bId]));
        selectedIds.clear();
        document.getElementById('groupBtn').disabled = true;
        document.getElementById('groupBtn').innerText = 'Group Selected in Order';
        document.getElementById('select-label').innerText = '0 selected';
        updateCounters();
        saveState();
    }}

    function ungroup(bId) {{
        const bEl = document.getElementById(bId);
        if (!bEl) return;
        const parent = bEl.parentElement;
        const memberIds = bundles[bId] ? bundles[bId].members : [];
        memberIds.forEach(id => {{
            if (guestMap[id]) parent.appendChild(createGuestCard(guestMap[id]));
        }});
        bEl.remove();
        delete bundles[bId];
        updateCounters();
        saveState();
    }}

    // ==========================================
    // FLOOR PLAN: ROUND TABLES & RESIZABLE PODIUM
    // ==========================================
    function renderFloorPlan() {{
        const canvas = document.getElementById('hall-canvas');
        Array.from(canvas.querySelectorAll('.round-table-node, .hall-element-node')).forEach(el => el.remove());

        activeTableIds.forEach(i => {{
            const z = document.getElementById('zone-table-' + i);
            let count = 0;
            if (z) {{
                Array.from(z.children).forEach(c => {{
                    count += c.classList.contains('bundle-card') ? (bundles[c.id] ? bundles[c.id].members.length : 0) : 1;
                }});
            }}

            const pos = tablePositions[i] || {{ x: 40 + ((i-1)%5)*200, y: 70 + Math.floor((i-1)/5)*170 }};
            const rTable = document.createElement('div');
            rTable.className = 'round-table-node';
            rTable.id = 'rt-node-' + i;
            rTable.style.left = pos.x + 'px';
            rTable.style.top = pos.y + 'px';

            let badgeColor = count > maxCapacity ? '#FEE2E2' : (count === maxCapacity ? '#DCFCE7' : '#F3F4F6');
            let textColor = count > maxCapacity ? '#DC2626' : (count === maxCapacity ? '#16A34A' : '#4B5563');

            rTable.innerHTML = `
                <div class="rt-num">T-${{i < 10 ? '0' + i : i}}</div>
                <div class="rt-name" title="${{tableNames[i] || ('TABLE ' + i)}}">${{tableNames[i] || ('TABLE ' + i)}}</div>
                <div class="rt-badge" style="background: ${{badgeColor}}; color: ${{textColor}};">${{count}}/${{maxCapacity}}</div>
            `;
            attachCanvasDraggable(rTable, (x, y) => {{
                tablePositions[i] = {{ x, y }};
                saveState();
            }});
            canvas.appendChild(rTable);
        }});

        hallItems.forEach((item, idx) => {{
            const el = document.createElement('div');
            el.className = 'hall-element-node';
            el.style.left = item.x + 'px';
            el.style.top = item.y + 'px';
            el.style.width = (item.w || 220) + 'px';
            el.style.height = (item.h || 75) + 'px';

            el.innerHTML = `
                <div class="node-header">
                    <span class="node-title" title="${{item.text}}">${{item.text}}</span>
                    <button class="del-btn" onclick="removeHallItem(${{idx}})" title="Delete item">✕</button>
                </div>
                <div class="node-controls">
                    <span style="font-size: 0.65rem; color: #9CA3AF; margin-right: 2px;">W:</span>
                    <button class="resize-btn" onclick="resizeItem(${{idx}}, -30, 0)">-</button>
                    <button class="resize-btn" onclick="resizeItem(${{idx}}, 30, 0)">+</button>
                    <span style="font-size: 0.65rem; color: #9CA3AF; margin-left: 4px; margin-right: 2px;">H:</span>
                    <button class="resize-btn" onclick="resizeItem(${{idx}}, 0, -20)">-</button>
                    <button class="resize-btn" onclick="resizeItem(${{idx}}, 0, 20)">+</button>
                </div>
            `;
            attachCanvasDraggable(el, (x, y) => {{
                item.x = x;
                item.y = y;
                saveState();
            }});
            canvas.appendChild(el);
        }});
    }}

    function addHallItem() {{
        const inp = document.getElementById('itemText');
        const txt = inp.value.trim();
        if (!txt) return;
        hallItems.push({{ text: txt, x: 440, y: 30, w: 220, h: 75 }});
        inp.value = '';
        saveState();
        renderFloorPlan();
    }}

    function removeHallItem(idx) {{
        hallItems.splice(idx, 1);
        saveState();
        renderFloorPlan();
    }}

    function resizeItem(idx, dw, dh) {{
        if (!hallItems[idx]) return;
        hallItems[idx].w = Math.max(90, (hallItems[idx].w || 220) + dw);
        hallItems[idx].h = Math.max(45, (hallItems[idx].h || 75) + dh);
        saveState();
        renderFloorPlan();
    }}

    function attachCanvasDraggable(element, onMoveCallback) {{
        let isDragging = false;
        let startX, startY, origX, origY;

        element.onmousedown = function(e) {{
            if (e.target.tagName.toLowerCase() === 'button' || e.target.tagName.toLowerCase() === 'input') return;
            isDragging = true;
            startX = e.clientX;
            startY = e.clientY;
            origX = parseInt(element.style.left) || 0;
            origY = parseInt(element.style.top) || 0;
            document.onmousemove = onMouseMove;
            document.onmouseup = onMouseUp;
            e.preventDefault();
        }};

        function onMouseMove(e) {{
            if (!isDragging) return;
            let dx = e.clientX - startX;
            let dy = e.clientY - startY;
            let newX = Math.max(5, Math.min(1100, origX + dx));
            let newY = Math.max(5, Math.min(730, origY + dy));
            element.style.left = newX + 'px';
            element.style.top = newY + 'px';
        }}

        function onMouseUp(e) {{
            if (isDragging) {{
                isDragging = false;
                document.onmousemove = null;
                document.onmouseup = null;
                onMoveCallback(parseInt(element.style.left), parseInt(element.style.top));
            }}
        }}
    }}

    function printFloorPlan() {{
        saveState();
        let planDoc = `
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="utf-8">
        <title>Hall Floor Plan</title>
        <style>
            @page {{ size: A4 landscape; margin: 8mm; }}
            body {{ font-family: sans-serif; margin: 0; padding: 10px; color: #111; }}
            h2 {{ text-align: center; margin: 0 0 10px 0; font-size: 14pt; letter-spacing: 2px; }}
            .canvas-print {{
                width: 1050px;
                height: 750px;
                border: 3px solid #111;
                position: relative;
                margin: 0 auto;
                background: #FFF;
            }}
            .rt-print {{
                width: 95px;
                height: 95px;
                border-radius: 50%;
                border: 2px solid #111;
                position: absolute;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                text-align: center;
                font-size: 8pt;
                font-weight: bold;
            }}
            .item-print {{
                position: absolute;
                background: #111;
                color: #FFF;
                border: 1px solid #000;
                display: flex;
                justify-content: center;
                align-items: center;
                font-weight: bold;
                font-size: 9pt;
                text-transform: uppercase;
                letter-spacing: 1px;
            }}
        </style>
        </head>
        <body>
        <h2>BALL COMMITTEE — HALL FLOOR PLAN</h2>
        <div class="canvas-print">
        `;

        activeTableIds.forEach(i => {{
            const z = document.getElementById('zone-table-' + i);
            let count = 0;
            if (z) {{
                Array.from(z.children).forEach(c => {{
                    count += c.classList.contains('bundle-card') ? (bundles[c.id] ? bundles[c.id].members : []) : [c.id];
                }});
            }}
            const pos = tablePositions[i] || {{ x: 40, y: 70 }};
            planDoc += `
                <div class="rt-print" style="left: ${{pos.x}}px; top: ${{pos.y}}px;">
                    <div>T-${{i < 10 ? '0' + i : i}}</div>
                    <div style="font-size: 6pt; color: #555;">${{tableNames[i] || ''}}</div>
                    <div style="font-size: 6pt;">${{count}}/${{maxCapacity}}</div>
                </div>
            `;
        }});

        hallItems.forEach(item => {{
            planDoc += `
                <div class="item-print" style="left: ${{item.x}}px; top: ${{item.y}}px; width: ${{item.w || 220}}px; height: ${{item.h || 75}}px;">
                    ${{item.text}}
                </div>
            `;
        }});

        planDoc += `</div></body></html>`;

        const w = window.open('', '_blank');
        if (w) {{
            w.document.write(planDoc);
            w.document.close();
            setTimeout(() => w.print(), 350);
        }}
    }}

    // ==========================================
    // PERSISTENCE & STORAGE
    // ==========================================
    function saveState() {{
        try {{
            const state = {{
                guestMap: guestMap,
                bundles: bundles,
                activeTableIds: activeTableIds,
                tableNames: tableNames,
                tablePositions: tablePositions,
                hallItems: hallItems,
                pool: Array.from(poolZone.children).map(c => ({{ type: c.classList.contains('bundle-card') ? 'bundle' : 'guest', id: c.id }})),
                tables: {{}}
            }};
            activeTableIds.forEach(i => {{
                const z = document.getElementById('zone-table-' + i);
                if (z) {{
                    state.tables[i] = Array.from(z.children).map(c => ({{ type: c.classList.contains('bundle-card') ? 'bundle' : 'guest', id: c.id }}));
                }}
            }});
            localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        }} catch(e) {{}}
    }}

    function loadState() {{
        try {{
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) return false;
            const state = JSON.parse(raw);

            if (state.guestMap) {{
                Object.keys(state.guestMap).forEach(k => {{
                    if (!guestMap[k]) {{
                        guestMap[k] = state.guestMap[k];
                    }} else {{
                        if (state.guestMap[k].name) guestMap[k].name = state.guestMap[k].name;
                        if (state.guestMap[k].dietary !== undefined) guestMap[k].dietary = state.guestMap[k].dietary;
                    }}
                }});
            }}

            bundles = state.bundles || {{}};
            if (state.activeTableIds && state.activeTableIds.length > 0) {{
                activeTableIds = state.activeTableIds;
            }}
            tableNames = state.tableNames || tableNames;
            tablePositions = state.tablePositions || tablePositions;
            hallItems = state.hallItems || [];

            rebuildTableGrid();

            const placed = new Set();
            poolZone.innerHTML = '';
            activeTableIds.forEach(i => {{
                const z = document.getElementById('zone-table-' + i);
                if (z) z.innerHTML = '';
            }});

            if (state.pool) {{
                state.pool.forEach(item => {{
                    if (item.type === 'bundle' && bundles[item.id]) {{
                        poolZone.appendChild(createBundleCard(item.id, bundles[item.id]));
                        (bundles[item.id].members || []).forEach(id => placed.add(id));
                    }} else if (item.type === 'guest' && guestMap[item.id]) {{
                        poolZone.appendChild(createGuestCard(guestMap[item.id]));
                        placed.add(item.id);
                    }}
                }});
            }}

            if (state.tables) {{
                activeTableIds.forEach(i => {{
                    const z = document.getElementById('zone-table-' + i);
                    if (z) {{
                        (state.tables[i] || []).forEach(item => {{
                            if (item.type === 'bundle' && bundles[item.id]) {{
                                z.appendChild(createBundleCard(item.id, bundles[item.id]));
                                (bundles[item.id].members || []).forEach(id => placed.add(id));
                            }} else if (item.type === 'guest' && guestMap[item.id]) {{
                                z.appendChild(createGuestCard(guestMap[item.id]));
                                placed.add(item.id);
                            }}
                        }});
                    }}
                }});
            }}

            Object.values(guestMap).forEach(g => {{
                if (!placed.has(g.id)) poolZone.appendChild(createGuestCard(g));
            }});
            return true;
        }} catch(e) {{
            return false;
        }}
    }}

    function autoBundleCommitteesOnly() {{
        let commGroups = {{}};
        let regularGuests = [];

        Object.values(guestMap).forEach(g => {{
            if (g.committee && g.committee.trim()) {{
                const c = g.committee.trim();
                if (!commGroups[c]) commGroups[c] = [];
                commGroups[c].push(g.id);
            }} else {{
                regularGuests.push(g);
            }}
        }});

        Object.keys(commGroups).sort().forEach(commName => {{
            const ids = commGroups[commName];
            for (let i = 0; i < ids.length; i += 10) {{
                const chunk = ids.slice(i, i + 10);
                if (chunk.length > 1) {{
                    const bId = 'bundle_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
                    bundles[bId] = {{ name: commName + ' (' + chunk.length + ')', members: chunk }};
                    poolZone.appendChild(createBundleCard(bId, bundles[bId]));
                }} else {{
                    poolZone.appendChild(createGuestCard(guestMap[chunk[0]]));
                }}
            }}
        }});

        regularGuests.forEach(g => {{
            poolZone.appendChild(createGuestCard(g));
        }});

        saveState();
    }}

    rebuildTableGrid();

    if (!loadState()) {{
        if (Object.keys(guestMap).length > 0) {{
            autoBundleCommitteesOnly();
        }}
    }}
    updateCounters();
    renderFloorPlan();

    function addGuestManual() {{
        const nEl = document.getElementById('manualName');
        const cEl = document.getElementById('manualComm');
        const dEl = document.getElementById('manualDiet');
        const name = nEl.value.trim();
        if (!name) return;

        const id = "manual_" + Date.now();
        const g = {{ id: id, name: name, committee: cEl.value.trim(), pref1: "", pref2: "", dietary: dEl.value.trim() }};
        guestMap[id] = g;
        poolZone.appendChild(createGuestCard(g));
        nEl.value = ""; cEl.value = ""; dEl.value = "";
        updateCounters();
        saveState();
    }}

    function updateCounters() {{
        let pCount = 0;
        Array.from(poolZone.children).forEach(c => {{
            pCount += c.classList.contains('bundle-card') ? (bundles[c.id] ? bundles[c.id].members.length : 0) : 1;
        }});
        document.getElementById('pool-count').innerText = pCount;

        activeTableIds.forEach(i => {{
            const z = document.getElementById('zone-table-' + i);
            let count = 0;
            if (z) {{
                Array.from(z.children).forEach(c => {{
                    count += c.classList.contains('bundle-card') ? (bundles[c.id] ? bundles[c.id].members.length : 0) : 1;
                }});
            }}
            const counter = document.getElementById('t-count-' + i);
            if (counter) {{
                counter.innerText = `${{count}}/${{maxCapacity}}`;
                counter.style.color = count > maxCapacity ? '#DC2626' : (count === maxCapacity ? '#16A34A' : '#6B7280');
            }}
        }});
    }}

    function resetLayout() {{
        if (confirm("Reset all seating positions, extra tables, and floor plan positions?")) {{
            localStorage.removeItem(STORAGE_KEY);
            location.reload();
        }}
    }}

    function exportCards() {{
        saveState();
        const rows = [["Table", "Seat", "Guest Name", "Committee", "Dietary", "Preferences"]];
        activeTableIds.forEach(i => {{
            const tName = tableNames[i] || ("TABLE " + (i < 10 ? '0' + i : i));
            const z = document.getElementById('zone-table-' + i);
            let seat = 1;
            if (z) {{
                Array.from(z.children).forEach(child => {{
                    const ids = child.classList.contains('bundle-card') ? (bundles[child.id] ? bundles[child.id].members : []) : [child.id];
                    ids.forEach(id => {{
                        const g = guestMap[id];
                        if (g) {{
                            rows.push([
                                `"${{tName.replace(/"/g, '""')}}"`,
                                seat++,
                                `"${{g.name.replace(/"/g, '""')}}"`,
                                `"${{(g.committee || '').replace(/"/g, '""')}}"`,
                                `"${{(g.dietary || '').replace(/"/g, '""')}}"`,
                                `"${{[g.pref1, g.pref2].filter(Boolean).join(' & ').replace(/"/g, '""')}}"`
                            ]);
                        }}
                    }});
                }});
            }}
        }});
        const csv = "\\uFEFF" + rows.map(r => r.join(";")).join("\\r\\n");
        const blob = new Blob([csv], {{ type: "text/csv;charset=utf-8;" }});
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "Guestlist_Place_Cards_Sorted.csv";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }}

    // =========================================================================
    // CLEAR STRUCTURED TABLE & ALLERGIES PRINTOUT MET RONDE TAFEL MET 10 STOELEN
    // =========================================================================
    function exportPdf() {{
        saveState();
        let doc = `
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="utf-8">
        <title>Ball Banquet — Seating & Dietary Plan</title>
        <style>
            @page {{ size: A4 portrait; margin: 8mm; }}
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 0; padding: 0; color: #111; }}
            h1 {{ text-align: center; font-size: 15pt; margin: 0 0 2px 0; text-transform: uppercase; letter-spacing: 2px; }}
            .sub-title {{ text-align: center; font-size: 7.5pt; color: #6B7280; margin-bottom: 10px; text-transform: uppercase; }}
            .tables-container {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; }}
            .table-block {{ border: 1.5px solid #111; border-radius: 4px; padding: 6px 8px; page-break-inside: avoid; background: #FFF; }}
            .table-heading {{ font-weight: 800; font-size: 8.5pt; border-bottom: 2px solid #111; padding-bottom: 3px; margin-bottom: 5px; display: flex; justify-content: space-between; }}
            
            .table-body-flex {{ display: flex; gap: 6px; align-items: center; justify-content: space-between; }}
            .guest-list-col {{ flex: 1; }}
            .table-diagram-col {{ width: 85px; display: flex; flex-direction: column; align-items: center; justify-content: center; flex-shrink: 0; padding-left: 4px; border-left: 1px dashed #E5E7EB; }}
            
            .guest-entry {{ display: flex; justify-content: space-between; align-items: center; font-size: 7pt; padding: 2.2px 0; border-bottom: 1px dotted #E5E7EB; }}
            .guest-primary {{ font-weight: 600; color: #111; }}
            .guest-comm {{ font-size: 6pt; color: #1D4ED8; font-weight: 700; margin-left: 3px; }}
            .allergy-tag {{ font-weight: 700; font-size: 6.5pt; color: #B45309; background: #FEF3C7; padding: 1px 4px; border-radius: 2px; text-align: right; }}
            .no-allergy {{ color: #9CA3AF; font-size: 6pt; }}
        </style>
        </head>
        <body>
        <h1>BALL COMMITTEE — BANQUET SEATING LIST</h1>
        <div class="sub-title">Official Table Seating Plan & Guest Allergies</div>
        <div class="tables-container">
        `;

        const roundTableSvg = `
        <svg width="80" height="80" viewBox="0 0 90 90">
            <circle cx="45" cy="45" r="19" fill="#F9FAFB" stroke="#111" stroke-width="1.3" />
            <text x="45" y="47" text-anchor="middle" font-size="6.5" font-family="sans-serif" font-weight="bold" fill="#4B5563">TABLE</text>
            <circle cx="45.0" cy="15.0" r="6.8" fill="#FFF" stroke="#111" stroke-width="1" />
            <circle cx="62.6" cy="20.7" r="6.8" fill="#FFF" stroke="#111" stroke-width="1" />
            <circle cx="73.5" cy="35.7" r="6.8" fill="#FFF" stroke="#111" stroke-width="1" />
            <circle cx="73.5" cy="54.3" r="6.8" fill="#FFF" stroke="#111" stroke-width="1" />
            <circle cx="62.6" cy="69.3" r="6.8" fill="#FFF" stroke="#111" stroke-width="1" />
            <circle cx="45.0" cy="75.0" r="6.8" fill="#FFF" stroke="#111" stroke-width="1" />
            <circle cx="27.4" cy="69.3" r="6.8" fill="#FFF" stroke="#111" stroke-width="1" />
            <circle cx="16.5" cy="54.3" r="6.8" fill="#FFF" stroke="#111" stroke-width="1" />
            <circle cx="16.5" cy="35.7" r="6.8" fill="#FFF" stroke="#111" stroke-width="1" />
            <circle cx="27.4" cy="20.7" r="6.8" fill="#FFF" stroke="#111" stroke-width="1" />
        </svg>
        <span style="font-size: 5.5pt; color: #6B7280; margin-top: 1px;">Seat Numbers</span>
        `;

        activeTableIds.forEach(i => {{
            const tName = tableNames[i] || ("TABLE " + (i < 10 ? '0' + i : i));
            const z = document.getElementById('zone-table-' + i);
            const cnt = document.getElementById('t-count-' + i) ? document.getElementById('t-count-' + i).innerText : '';

            doc += `
            <div class="table-block">
                <div class="table-heading">
                    <span>${{tName}}</span>
                    <span>${{cnt}}</span>
                </div>
                <div class="table-body-flex">
                    <div class="guest-list-col">
            `;

            let seat = 1;
            if (z && z.children.length > 0) {{
                Array.from(z.children).forEach(child => {{
                    const ids = child.classList.contains('bundle-card') ? (bundles[child.id] ? bundles[child.id].members : []) : [child.id];
                    ids.forEach(id => {{
                        const g = guestMap[id];
                        if (g) {{
                            let allergyDisplay = g.dietary ? `<span class="allergy-tag">Allergy: ${{g.dietary}}</span>` : `<span class="no-allergy">None</span>`;
                            let commDisplay = g.committee ? `<span class="guest-comm">[${{g.committee}}]</span>` : '';
                            doc += `
                            <div class="guest-entry">
                                <div><b>${{seat++}}.</b> <span class="guest-primary">${{g.name}}</span> ${{commDisplay}}</div>
                                <div>${{allergyDisplay}}</div>
                            </div>
                            `;
                        }}
                    }});
                }});
            }} else {{
                doc += `<div style="font-size: 7pt; color: #9CA3AF; font-style: italic; padding: 4px 0;">No guests assigned yet.</div>`;
            }}

            doc += `
                    </div>
                    <div class="table-diagram-col">
                        ${{roundTableSvg}}
                    </div>
                </div>
            </div>
            `;
        }});

        doc += `</div></body></html>`;

        const w = window.open('', '_blank');
        if (w) {{
            w.document.write(doc);
            w.document.close();
            setTimeout(() => w.print(), 350);
        }}
    }}
</script>
</body>
</html>
"""

components.html(html_code, height=1400, scrolling=True)

if parsed_guests:
    st.markdown("---")
    st.markdown("### Kitchen Dietary Overview")
    diet_attendees = [g for g in parsed_guests if g.get('dietary')]
    if diet_attendees:
        st.dataframe(pd.DataFrame([{
            'Guest': g['name'],
            'Committee': g['committee'] or "Guest",
            'Dietary': g['dietary'],
            'Preferences': " & ".join(filter(None, [g['pref1'], g['pref2']]))
        } for g in diet_attendees]), use_container_width=True, hide_index=True)
    else:
        st.info("No dietary requirements reported.")
