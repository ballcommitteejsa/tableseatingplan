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

# PDF Generation imports
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

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
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@500;700&family=Montserrat:wght@300;400;600&display=swap');
    .stApp { background-color: #F8F9FA; font-family: 'Montserrat', sans-serif; }
    h1, h2, h3, .gala-title { font-family: 'Cinzel', serif !important; letter-spacing: 2px; color: #111111; text-align: center; }
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
        st.markdown("<h2 class='gala-title'>BALL COMMITTEE</h2>", unsafe_allow_html=True)

col_spacer_l, col_center, col_spacer_r = st.columns([1.5, 2, 1.5])
with col_center:
    render_ball_logo()

st.markdown("<p style='text-align: center; letter-spacing: 3px; color: #6B7280; text-transform: uppercase; font-size: 0.85rem; margin-top: -10px;'>Interactive Drag & Drop Seating Studio</p>", unsafe_allow_html=True)
st.markdown("<hr style='border: none; border-top: 1px solid #E5E7EB; margin-bottom: 1.5rem;'>", unsafe_allow_html=True)

# ==========================================
# 2. INTELLIGENTE GASTENLIJST PARSER (MET DUO-TICKET FILTER)
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
    s = str(d).strip()
    if s.lower() in NON_DIETARY or s.lower().startswith('no allerg'):
        return ""
    return s

def resolve_committee_alias(text):
    if not text or pd.isna(text): return None
    raw = str(text).strip().lower()
    for canonical, variations in COMMITTEE_ALIASES.items():
        for var in variations:
            if raw == var or f" {var} " in f" {raw} ":
                return canonical
    return None

def parse_uploaded_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    target_sheet = None
    for s in xls.sheet_names:
        if any(w in s.lower() for w in ['ticket', 'attendee', 'guest', 'banquet', '1)', '2)']):
            target_sheet = s
            break
    if not target_sheet:
        target_sheet = xls.sheet_names[0]

    df_raw = pd.read_excel(uploaded_file, sheet_name=target_sheet).dropna(how='all').reset_index(drop=True)
    is_hitract = 'Name ticket 1 ' in df_raw.columns or 'Förnamn' in df_raw.columns

    attendees = []
    
    if is_hitract:
        buyer_counts = {}
        for idx, row in df_raw.iterrows():
            fname = str(row.get('Förnamn', '')).strip() if pd.notna(row.get('Förnamn')) else ''
            lname = str(row.get('Efternamn', '')).strip() if pd.notna(row.get('Efternamn')) else ''
            sdate = str(row.get('Försäljnignsdatum', '')).strip() if pd.notna(row.get('Försäljnignsdatum')) else ''
            b_key = f"{fname}_{lname}_{sdate}"
            occ = buyer_counts.get(b_key, 0) + 1
            buyer_counts[b_key] = occ

            n1 = str(row.get('Name ticket 1 ', '')).strip() if pd.notna(row.get('Name ticket 1 ')) else f"{fname} {lname}".strip()
            n2 = str(row.get("Second guest's full name (if purchasing a second ticket)", '')).strip() if pd.notna(row.get("Second guest's full name (if purchasing a second ticket)")) else ''
            if not n2 and occ > 1:
                n2 = f"Guest of {n1}"

            c1 = str(row.get('Committee table', '')).strip() if pd.notna(row.get('Committee table')) else ''
            p1_1 = str(row.get('Seating preference 1: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 1: Full name (First & Last name) or committee name. ')) else ''
            p1_2 = str(row.get('Seating preference 2: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 2: Full name (First & Last name) or committee name. ')) else ''
            diet1 = clean_dietary_entry(row.get('Dietary requirements (leave blank if none)'))

            c2 = str(row.get('Committee table.1', '')).strip() if pd.notna(row.get('Committee table.1')) else ''
            p2_1 = str(row.get('Seating preference 1 second guest: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 1 second guest: Full name (First & Last name) or committee name. ')) else ''
            p2_2 = str(row.get('Seating preference 2 second guest: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 2 second guest: Full name (First & Last name) or committee name. ')) else ''
            diet2 = clean_dietary_entry(row.get('Dietary requirements second guest (leave blank if none)'))

            if occ == 1 or not n2:
                comm = resolve_committee_alias(c1) or resolve_committee_alias(p1_1) or resolve_committee_alias(p1_2) or (c1 if c1 else None)
                attendees.append({
                    'id': f"guest_{idx}_1",
                    'name': n1,
                    'committee': comm or "",
                    'pref1': p1_1 if not resolve_committee_alias(p1_1) else "",
                    'pref2': p1_2 if not resolve_committee_alias(p1_2) else "",
                    'dietary': diet1
                })
            else:
                comm = resolve_committee_alias(c2) or resolve_committee_alias(p2_1) or resolve_committee_alias(p2_2) or (c2 if c2 else None)
                attendees.append({
                    'id': f"guest_{idx}_2",
                    'name': n2,
                    'committee': comm or "",
                    'pref1': p2_1 if not resolve_committee_alias(p2_1) else n1,
                    'pref2': p2_2 if not resolve_committee_alias(p2_2) else "",
                    'dietary': diet2
                })
    else:
        df = df_raw.iloc[:, 0:5].copy()
        for idx, row in df.iterrows():
            name = str(row.iloc[0]).strip()
            comm = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() not in ['nan', 'None'] else ""
            p1 = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) and str(row.iloc[2]).strip() not in ['nan', 'None'] else ""
            p2 = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) and str(row.iloc[3]).strip() not in ['nan', 'None'] else ""
            diet = clean_dietary_entry(row.iloc[4]) if len(row) > 4 else ""
            if name:
                attendees.append({
                    'id': f"guest_{idx}",
                    'name': name,
                    'committee': resolve_committee_alias(comm) or comm,
                    'pref1': p1,
                    'pref2': p2,
                    'dietary': diet
                })
    return attendees

# ==========================================
# 3. ZIJBALK & UPLOAD
# ==========================================
with st.sidebar:
    render_ball_logo()
    st.markdown("### ⚙️ Studio Instellingen")
    seats_per_table = st.number_input("Max Stoelen per Tafel", min_value=4, max_value=20, value=10, step=1)
    st.markdown("---")
    st.caption("💡 **Tip:** Sleep gastenblokken van de linkerlijst direct in één van de 20 tafels.")

uploaded_file = st.file_uploader("Upload Hitract Gastenlijst (.xlsx)", type=["xlsx"])

if uploaded_file:
    if 'current_file_name' not in st.session_state or st.session_state['current_file_name'] != uploaded_file.name:
        st.session_state['current_file_name'] = uploaded_file.name
        parsed_guests = parse_uploaded_excel(uploaded_file)
        st.session_state['all_guests'] = parsed_guests
        
        # Init 20 tafels
        st.session_state['seating_plan'] = {f"Tafel {i}": [] for i in range(1, 21)}
        st.session_state['unassigned'] = [g['id'] for g in parsed_guests]

    all_guests_dict = {g['id']: g for g in st.session_state['all_guests']}
    
    st.success(f"✅ **{len(st.session_state['all_guests'])}** gasten ingeladen (duo-tickets automatisch gesplitst).")

    # ==========================================
    # 4. INTERACTIEF DRAG & DROP BORD (HTML/JS)
    # ==========================================
    guests_json = json.dumps(st.session_state['all_guests'])

    dnd_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: #F8F9FA;
            margin: 0;
            padding: 10px;
        }}
        .board-container {{
            display: flex;
            gap: 20px;
            height: 800px;
        }}
        .pool-panel {{
            width: 320px;
            background: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-top: 3px solid #111111;
            border-radius: 6px;
            display: flex;
            flex-direction: column;
            box-shadow: 0 4px 6px rgba(0,0,0,0.03);
        }}
        .pool-header {{
            padding: 12px 16px;
            background: #111111;
            color: #FFFFFF;
            font-weight: 700;
            font-size: 0.95rem;
            letter-spacing: 1px;
            display: flex;
            justify-content: space-between;
        }}
        .pool-search {{
            padding: 10px 14px;
            border-bottom: 1px solid #E5E7EB;
        }}
        .pool-search input {{
            width: 90%;
            padding: 8px 10px;
            border: 1px solid #D1D5DB;
            border-radius: 4px;
            font-size: 0.85rem;
        }}
        .drop-list {{
            flex: 1;
            overflow-y: auto;
            padding: 10px;
            min-height: 200px;
        }}
        .tables-grid {{
            flex: 1;
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 16px;
            overflow-y: auto;
            padding-right: 8px;
        }}
        .table-card {{
            background: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-top: 3px solid #374151;
            border-radius: 6px;
            display: flex;
            flex-direction: column;
            min-height: 260px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        }}
        .table-header {{
            padding: 8px 12px;
            background: #F3F4F6;
            border-bottom: 1px solid #E5E7EB;
            font-weight: 700;
            font-size: 0.88rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .guest-card {{
            background: #FFFFFF;
            border: 1px solid #D1D5DB;
            border-radius: 4px;
            padding: 8px 10px;
            margin-bottom: 8px;
            cursor: grab;
            transition: transform 0.15s, box-shadow 0.15s;
        }}
        .guest-card:hover {{
            box-shadow: 0 4px 8px rgba(0,0,0,0.08);
            border-color: #111111;
        }}
        .guest-card.dragging {{
            opacity: 0.4;
        }}
        .guest-name {{
            font-weight: 600;
            font-size: 0.88rem;
            color: #111111;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .guest-prefs {{
            font-size: 0.75rem;
            color: #6B7280;
            margin-top: 4px;
            line-height: 1.25;
        }}
        .badge-diet {{
            display: inline-block;
            background: #FEF3C7;
            color: #92400E;
            font-size: 0.70rem;
            font-weight: 600;
            padding: 2px 6px;
            border-radius: 3px;
            border: 1px solid #FDE68A;
            margin-top: 4px;
        }}
        .badge-comm {{
            font-size: 0.68rem;
            font-weight: 700;
            color: #1E40AF;
            background: #EFF6FF;
            padding: 1px 5px;
            border-radius: 3px;
            text-transform: uppercase;
        }}
        .drag-over {{
            background-color: #F0FDF4 !important;
            border: 2px dashed #16A34A !important;
        }}
    </style>
    </head>
    <body>
    <div class="board-container">
        <!-- LINKER PANEEL: ONGEPLAATSTE GASTEN -->
        <div class="pool-panel">
            <div class="pool-header">
                <span>ONGEPLAATSTE GASTEN</span>
                <span id="pool-count">0</span>
            </div>
            <div class="pool-search">
                <input type="text" id="searchInput" placeholder="Zoek op naam of voorkeur..." onkeyup="filterPool()">
            </div>
            <div class="drop-list" id="unassigned-zone" ondragover="allowDrop(event)" ondrop="handleDrop(event, 'unassigned')" ondragleave="dragLeave(event)">
            </div>
        </div>

        <!-- RECHTS: 20 TAFELS -->
        <div class="tables-grid" id="tablesContainer">
            <!-- Wordt dynamisch gevuld met 20 tafels -->
        </div>
    </div>

    <script>
        const allGuests = {guests_json};
        const guestMap = {{}};
        allGuests.forEach(g => guestMap[g.id] = g);

        // Tafels 1 t/m 20
        const tables = {{}};
        for (let i = 1; i <= 20; i++) {{
            tables['Tafel ' + i] = [];
        }}

        // Initialiseer HTML structuur van de 20 tafels
        const grid = document.getElementById('tablesContainer');
        for (let i = 1; i <= 20; i++) {{
            const tName = 'Tafel ' + i;
            const card = document.createElement('div');
            card.className = 'table-card';
            card.innerHTML = `
                <div class="table-header">
                    <span>${{tName.toUpperCase()}}</span>
                    <span id="count-${{i}}" style="font-weight: normal; color: #4B5563; font-size: 0.8rem;">0/{seats_per_table}</span>
                </div>
                <div class="drop-list" id="table-${{i}}" ondragover="allowDrop(event)" ondrop="handleDrop(event, '${{tName}}')" ondragleave="dragLeave(event)">
                </div>
            `;
            grid.appendChild(card);
        }}

        // Maak gast-kaart element
        function createGuestCard(g) {{
            const div = document.createElement('div');
            div.className = 'guest-card';
            div.id = g.id;
            div.draggable = true;
            div.ondragstart = (e) => dragStart(e, g.id);

            let commBadge = g.committee ? `<span class="badge-comm">${{g.committee}}</span>` : '';
            let dietBadge = g.dietary ? `<div class="badge-diet">⚠️ ${{g.dietary}}</div>` : '';
            
            let prefText = '';
            if (g.pref1 || g.pref2) {{
                let pList = [];
                if (g.pref1) pList.push('1: ' + g.pref1);
                if (g.pref2) pList.push('2: ' + g.pref2);
                prefText = `<div class="guest-prefs">${{pList.join(' • ')}}</div>`;
            }}

            div.innerHTML = `
                <div class="guest-name">
                    <span>${{g.name}}</span>
                    ${{commBadge}}
                </div>
                ${{prefText}}
                ${{dietBadge}}
            `;
            return div;
        }}

        // Plaats alle gasten initieel in de ongeplaatste pool
        const poolZone = document.getElementById('unassigned-zone');
        allGuests.forEach(g => {{
            poolZone.appendChild(createGuestCard(g));
        }});
        document.getElementById('pool-count').innerText = allGuests.length;

        // Drag & Drop handlers
        let draggedId = null;

        function dragStart(e, id) {{
            draggedId = id;
            e.dataTransfer.setData('text/plain', id);
        }}

        function allowDrop(e) {{
            e.preventDefault();
            e.currentTarget.classList.add('drag-over');
        }}

        function dragLeave(e) {{
            e.currentTarget.classList.remove('drag-over');
        }}

        function handleDrop(e, targetZoneName) {{
            e.preventDefault();
            e.currentTarget.classList.remove('drag-over');
            if (!draggedId) return;

            const element = document.getElementById(draggedId);
            if (element) {{
                e.currentTarget.appendChild(element);
                updateCounts();
            }}
        }}

        function updateCounts() {{
            document.getElementById('pool-count').innerText = document.getElementById('unassigned-zone').children.length;
            for (let i = 1; i <= 20; i++) {{
                const count = document.getElementById('table-' + i).children.length;
                const counter = document.getElementById('count-' + i);
                counter.innerText = count + '/{seats_per_table}';
                counter.style.color = count > {seats_per_table} ? '#DC2626' : '#4B5563';
                counter.style.fontWeight = count > {seats_per_table} ? 'bold' : 'normal';
            }}
        }}

        function filterPool() {{
            const val = document.getElementById('searchInput').value.toLowerCase();
            const cards = document.getElementById('unassigned-zone').getElementsByClassName('guest-card');
            for (let card of cards) {{
                const text = card.innerText.toLowerCase();
                card.style.display = text.includes(val) ? 'block' : 'none';
            }}
        }}
    </script>
    </body>
    </html>
    """

    components.html(dnd_html, height=840, scrolling=True)

    # ==========================================
    # 5. EXPORT & DIEET OVERZICHT (A4 & EXCEL)
    # ==========================================
    st.markdown("<hr style='border: none; border-top: 1px solid #E5E7EB; margin: 2rem 0 1rem;'>", unsafe_allow_html=True)
    st.markdown("### 📋 Keuken & Service Briefing (Geregistreerde Allergieën)")

    diet_guests = [g for g in st.session_state['all_guests'] if g.get('dietary')]
    if diet_guests:
        diet_df = pd.DataFrame([{
            'Gast': g['name'],
            'Commissie': g['committee'] or "Gast",
            'Dieet / Allergie': g['dietary'],
            'Voorkeur 1': g['pref1'],
            'Voorkeur 2': g['pref2']
        } for g in diet_guests])
        st.dataframe(diet_df, use_container_width=True, hide_index=True)
    else:
        st.info("Geen dieetwensen gemeld in de ingeladen lijst.")
