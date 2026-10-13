import io
import os
import re
import json
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import openpyxl
from PIL import Image

# ==========================================
# 1. PAGINA CONFIGURATIE & HUISSTIJL
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
    
    /* Moderne Knoppen */
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
    
    /* Upload venster styling */
    .upload-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 8px;
        padding: 20px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

# Logo weergave
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
# 2. INTELLIGENTE PARSER & ALIAS NORMALISATIE
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
    parts = re.split(r'[,&/+\n]|\s+och\s+|\s+and\s+', s, flags=re.IGNORECASE)
    parts = [p.strip() for p in parts if p.strip()]
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

    # Schone 5-koloms data ondersteuning
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
                    'dietary': clean_dietary_text(row.get(diet_c, '')) if diet_c else ""
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

        # Voorkeuren Koper 1
        raw_p1 = row.get(col_pref1, '') if col_pref1 else ''
        p1_a, p1_b = split_pair_names(raw_p1)
        if col_pref2 and pd.notna(row.get(col_pref2)):
            raw_p2 = str(row.get(col_pref2)).strip()
            if raw_p2 and raw_p2.lower() not in NON_DIETARY:
                p1_b = raw_p2

        diet1 = clean_dietary_text(row.get(col_diet1, '')) if col_diet1 else ''
        comm1 = resolve_comm_alias(row.get(col_comm1, '')) if col_comm1 else resolve_comm_alias(raw_p1)

        # Voorkeuren Gast 2
        raw_p2_1 = row.get(col_pref2_1, '') if col_pref2_1 else ''
        p2_a, p2_b = split_pair_names(raw_p2_1)
        if col_pref2_2 and pd.notna(row.get(col_pref2_2)):
            raw_p2_2 = str(row.get(col_pref2_2)).strip()
            if raw_p2_2 and raw_p2_2.lower() not in NON_DIETARY:
                p2_b = raw_p2_2

        diet2 = clean_dietary_text(row.get(col_diet2, '')) if col_diet2 else ''
        comm2 = resolve_comm_alias(row.get(col_comm2, '')) if col_comm2 else resolve_comm_alias(raw_p2_1)

        if occ == 1 or not n2:
            attendees.append({
                'id': f"g_{idx}_1",
                'name': n1,
                'committee': comm1 or "",
                'pref1': p1_a if not comm1 else "",
                'pref2': p1_b if not comm1 else "",
                'dietary': diet1
            })
        else:
            attendees.append({
                'id': f"g_{idx}_2",
                'name': n2,
                'committee': comm2 or "",
                'pref1': p2_a if not comm2 else n1,
                'pref2': p2_b if not comm2 else "",
                'dietary': diet2
            })

    return attendees

# ==========================================
# 3. UPLOAD & BEDIENINGSBALK
# ==========================================
st.markdown("<div class='upload-card'>", unsafe_allow_html=True)
col_file, col_seats, col_action = st.columns([3, 1, 1])

with col_file:
    uploaded_file = st.file_uploader(
        "Kies je Hitract Excel bestand (.xlsx)",
        type=["xlsx"],
        key="uploader_input_field"
    )

with col_seats:
    table_max_seats = st.number_input(
        "Stoelen per tafel",
        min_value=4,
        max_value=25,
        value=10,
        step=1
    )

with col_action:
    st.write("")
    st.write("")
    btn_start = st.button("🚀 Start / Laad Lijst", use_container_width=True)

st.markdown("</div>", unsafe_allow_html=True)

# Verwerking via Start-knop of directe upload
if btn_start or (uploaded_file is not None and 'active_guestlist' not in st.session_state):
    if uploaded_file is not None:
        with st.spinner("Gastenlijst analyseren en duo-tickets splitsen..."):
            loaded_guests = parse_hitract_excel_dataset(uploaded_file)
            st.session_state['active_guestlist'] = loaded_guests
            st.session_state['loaded_file_name'] = uploaded_file.name
            st.success(f"🎉 **{len(loaded_guests)}** gasten succesvol ingeladen uit `{uploaded_file.name}`!")
    else:
        st.warning("⚠️ Upload eerst een Excel-bestand voordat je op Start klikt.")

parsed_guests = st.session_state.get('active_guestlist', [])
guests_json_payload = json.dumps(parsed_guests)

# ==========================================
# 4. OVERZICHTELIJK 20-TAFEL DRAG & DROP BORD
# ==========================================
dnd_html_template = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    * { box-sizing: border-box; }
    body {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
        background-color: #F8F9FA;
        margin: 0;
        padding: 6px;
        user-select: none;
    }
    
    .top-action-bar {
        background: #111111;
        color: #FFFFFF;
        padding: 10px 14px;
        border-radius: 6px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 12px;
    }
    .group-btn {
        background: #2563EB;
        color: #FFFFFF;
        border: none;
        padding: 6px 14px;
        border-radius: 4px;
        font-weight: 600;
        cursor: pointer;
        font-size: 0.85rem;
        transition: background 0.2s;
    }
    .group-btn:hover { background: #1D4ED8; }
    .group-btn:disabled { background: #4B5563; cursor: not-allowed; opacity: 0.6; }

    details.pool-tray {
        background: #FFFFFF;
        border: 1px solid #D1D5DB;
        border-radius: 6px;
        margin-bottom: 14px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.03);
    }
    details.pool-tray summary {
        padding: 10px 14px;
        font-weight: 700;
        cursor: pointer;
        background: #F3F4F6;
        color: #111111;
        border-radius: 6px;
        display: flex;
        justify-content: space-between;
    }
    .pool-cards-container {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        padding: 10px;
        max-height: 240px;
        overflow-y: auto;
        background: #F9FAFB;
    }

    /* 20 TAFELS RASTER: 5 Kolommen x 4 Rijen */
    .tables-grid-20 {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 12px;
        width: 100%;
    }
    .table-column {
        background: #FFFFFF;
        border: 1px solid #D1D5DB;
        border-top: 3px solid #111111;
        border-radius: 4px;
        display: flex;
        flex-direction: column;
        min-height: 320px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .table-header {
        padding: 6px 8px;
        background: #F3F4F6;
        border-bottom: 1px solid #E5E7EB;
        font-size: 0.78rem;
        font-weight: 700;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .table-dropzone {
        flex: 1;
        padding: 6px;
        min-height: 270px;
        display: flex;
        flex-direction: column;
        gap: 5px;
    }
    .table-dropzone.drag-over {
        background-color: #ECFDF5;
        border: 2px dashed #059669;
    }

    /* GASTEN BLOK */
    .guest-item {
        background: #FFFFFF;
        border: 1px solid #D1D5DB;
        border-left: 3px solid #2563EB;
        border-radius: 3px;
        padding: 5px 6px;
        cursor: grab;
        display: flex;
        flex-direction: column;
        transition: all 0.15s;
    }
    .guest-item:hover {
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        border-color: #111111;
    }
    .guest-item.is-selected {
        background: #EFF6FF !important;
        border-color: #2563EB !important;
    }
    .guest-top-row {
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .guest-name {
        font-weight: 600;
        font-size: 0.78rem;
        color: #111111;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        flex: 1;
    }
    .pref-line {
        font-size: 0.68rem;
        color: #4B5563;
        margin-top: 2px;
        line-height: 1.15;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .tag-diet {
        font-size: 0.62rem;
        font-weight: bold;
        color: #B45309;
        background: #FEF3C7;
        padding: 1px 3px;
        border-radius: 2px;
        margin-top: 2px;
        display: inline-block;
        width: fit-content;
    }
    .tag-comm {
        font-size: 0.62rem;
        font-weight: bold;
        color: #1E3A8A;
        background: #DBEAFE;
        padding: 1px 3px;
        border-radius: 2px;
    }

    /* GROEPSBLOK / BUNDEL */
    .bundle-item {
        background: #FEF9C3;
        border: 1.5px solid #CA8A04;
        border-radius: 4px;
        padding: 5px;
        cursor: grab;
        display: flex;
        flex-direction: column;
        gap: 4px;
    }
    .bundle-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 0.72rem;
        font-weight: 700;
        color: #854D0E;
        border-bottom: 1px dashed #EAB308;
        padding-bottom: 3px;
    }
    .ungroup-btn {
        background: #FFFFFF;
        border: 1px solid #CA8A04;
        color: #854D0E;
        border-radius: 3px;
        padding: 1px 5px;
        font-size: 0.65rem;
        cursor: pointer;
        font-weight: bold;
    }
    .ungroup-btn:hover { background: #FEF08A; }
</style>
</head>
<body>

<script id="guests-raw-data" type="application/json">
__GUESTS_JSON__
</script>

<div class="top-action-bar">
    <div style="font-weight: bold; font-size: 0.9rem;">
        🎯 TAFELINDELING STUDIO &nbsp;|&nbsp; <span style="font-size: 0.8rem; font-weight: normal; color: #9CA3AF;">Alle 20 tafels in beeld</span>
    </div>
    <div>
        <span id="selected-count-label" style="font-size: 0.8rem; margin-right: 12px; color: #D1D5DB;">0 gasten aangevinkt</span>
        <button class="group-btn" id="groupBtn" onclick="bundleSelectedGuests()" disabled>🔗 Groepeer als 1 Blok</button>
    </div>
</div>

<details class="pool-tray" open>
    <summary>
        <span>📥 ONGEPLAATSTE GASTEN WACHTRIJ (<span id="unassigned-count">0</span>) — Klik om in/uit te klappen</span>
        <span style="font-size: 0.78rem; font-weight: normal; color: #6B7280;">Vink aan om een groep te vormen of sleep direct naar een tafel</span>
    </summary>
    <div class="pool-cards-container" id="pool-zone" ondragover="allowDrop(event)" ondrop="handleDrop(event, 'pool')" ondragleave="dragLeave(event)">
    </div>
</details>

<div class="tables-grid-20" id="tables-grid">
</div>

<script>
    const tableMaxCapacity = __MAX_CAPACITY__;
    const rawDataText = document.getElementById('guests-raw-data').textContent;
    let initialGuests = [];
    try {
        initialGuests = JSON.parse(rawDataText);
    } catch(err) {
        console.error("Kon gastenlijst niet parsen:", err);
    }

    const guestMap = {};
    initialGuests.forEach(g => guestMap[g.id] = g);

    let selectedGuestIds = new Set();
    let bundles = {};
    let draggedElementId = null;

    const grid = document.getElementById('tables-grid');
    for (let i = 1; i <= 20; i++) {
        const tCol = document.createElement('div');
        tCol.className = 'table-column';
        tCol.innerHTML = `
            <div class="table-header">
                <span>TAFEL ${i < 10 ? '0' + i : i}</span>
                <span id="t-count-${i}" style="color: #6B7280;">0/${tableMaxCapacity}</span>
            </div>
            <div class="table-dropzone" id="zone-table-${i}" ondragover="allowDrop(event)" ondrop="handleDrop(event, 'table-${i}')" ondragleave="dragLeave(event)">
            </div>
        `;
        grid.appendChild(tCol);
    }

    function renderGuestCard(g) {
        const div = document.createElement('div');
        div.className = 'guest-item';
        div.id = g.id;
        div.draggable = true;
        div.ondragstart = (e) => dragStart(e, g.id);

        let commTag = g.committee ? `<span class="tag-comm">${g.committee}</span>` : '';
        let dietTag = g.dietary ? `<div class="tag-diet">⚠️ ${g.dietary}</div>` : '';

        let prefs = [];
        if (g.pref1) prefs.push(g.pref1);
        if (g.pref2) prefs.push(g.pref2);
        let prefText = '';
        if (prefs.length > 0) {
            prefText = `<div class="pref-line" title="Voorkeuren: ${prefs.join(', ')}">👉 ${prefs.join(' & ')}</div>`;
        }

        div.innerHTML = `
            <div class="guest-top-row">
                <input type="checkbox" onchange="toggleSelect(event, '${g.id}')" style="cursor: pointer; margin: 0 4px 0 0;">
                <span class="guest-name" title="${g.name}">${g.name}</span>
                ${commTag}
            </div>
            ${prefText}
            ${dietTag}
        `;
        return div;
    }

    function renderBundleCard(bundleId, memberIds) {
        const bundleDiv = document.createElement('div');
        bundleDiv.className = 'bundle-item';
        bundleDiv.id = bundleId;
        bundleDiv.draggable = true;
        bundleDiv.ondragstart = (e) => dragStart(e, bundleId);

        let namesList = memberIds.map(mid => guestMap[mid].name).join(', ');

        bundleDiv.innerHTML = `
            <div class="bundle-header">
                <span>👥 GROEP (${memberIds.length} pers.)</span>
                <button class="ungroup-btn" onclick="ungroupBundle('${bundleId}')">✂️ Ontkoppel</button>
            </div>
            <div style="font-size: 0.72rem; font-weight: 600; color: #713F12;">${namesList}</div>
        `;
        return bundleDiv;
    }

    const poolZone = document.getElementById('pool-zone');
    if (initialGuests.length > 0) {
        initialGuests.forEach(g => {
            poolZone.appendChild(renderGuestCard(g));
        });
    } else {
        poolZone.innerHTML = '<div style="font-size: 0.85rem; color: #9CA3AF; padding: 12px; font-style: italic;">Nog geen gastenlijst gestart. Upload een Excel-bestand hierboven en klik op Start.</div>';
    }
    updateCounters();

    function toggleSelect(e, gid) {
        e.stopPropagation();
        if (e.target.checked) {
            selectedGuestIds.add(gid);
            document.getElementById(gid).classList.add('is-selected');
        } else {
            selectedGuestIds.delete(gid);
            document.getElementById(gid).classList.remove('is-selected');
        }
        
        const btn = document.getElementById('groupBtn');
        const lbl = document.getElementById('selected-count-label');
        const cnt = selectedGuestIds.size;
        lbl.innerText = `${cnt} gasten aangevinkt`;
        btn.disabled = cnt < 2;
        btn.innerText = cnt >= 2 ? `🔗 Maak Groepsblok van ${cnt} pers.` : `🔗 Groepeer als 1 Blok`;
    }

    function bundleSelectedGuests() {
        if (selectedGuestIds.size < 2) return;

        const bundleId = 'bundle_' + Date.now();
        const memberIds = Array.from(selectedGuestIds);
        bundles[bundleId] = memberIds;

        const firstEl = document.getElementById(memberIds[0]);
        const targetContainer = firstEl ? firstEl.parentElement : poolZone;

        memberIds.forEach(mid => {
            const el = document.getElementById(mid);
            if (el) el.remove();
        });

        const bundleEl = renderBundleCard(bundleId, memberIds);
        targetContainer.appendChild(bundleEl);

        selectedGuestIds.clear();
        document.getElementById('groupBtn').disabled = true;
        document.getElementById('selected-count-label').innerText = '0 gasten aangevinkt';
        updateCounters();
    }

    function ungroupBundle(bundleId) {
        const bundleEl = document.getElementById(bundleId);
        if (!bundleEl) return;
        const parent = bundleEl.parentElement;
        const memberIds = bundles[bundleId];

        memberIds.forEach(mid => {
            parent.appendChild(renderGuestCard(guestMap[mid]));
        });

        bundleEl.remove();
        delete bundles[bundleId];
        updateCounters();
    }

    function dragStart(e, id) {
        draggedElementId = id;
        e.dataTransfer.setData('text/plain', id);
    }

    function allowDrop(e) {
        e.preventDefault();
        e.currentTarget.classList.add('drag-over');
    }

    function dragLeave(e) {
        e.currentTarget.classList.remove('drag-over');
    }

    function handleDrop(e, targetZone) {
        e.preventDefault();
        e.currentTarget.classList.remove('drag-over');
        if (!draggedElementId) return;

        const el = document.getElementById(draggedElementId);
        if (el) {
            e.currentTarget.appendChild(el);
            updateCounters();
        }
    }

    function updateCounters() {
        let poolCount = 0;
        Array.from(poolZone.children).forEach(child => {
            if (child.classList && child.classList.contains('bundle-item')) {
                poolCount += bundles[child.id].length;
            } else if (child.classList && child.classList.contains('guest-item')) {
                poolCount += 1;
            }
        });
        document.getElementById('unassigned-count').innerText = poolCount;

        for (let i = 1; i <= 20; i++) {
            const zone = document.getElementById('zone-table-' + i);
            let tCount = 0;
            Array.from(zone.children).forEach(child => {
                if (child.classList && child.classList.contains('bundle-item')) {
                    tCount += bundles[child.id].length;
                } else if (child.classList && child.classList.contains('guest-item')) {
                    tCount += 1;
                }
            });
            const counter = document.getElementById('t-count-' + i);
            counter.innerText = `${tCount}/${tableMaxCapacity}`;
            counter.style.color = tCount > tableMaxCapacity ? '#DC2626' : (tCount === tableMaxCapacity ? '#16A34A' : '#6B7280');
            counter.style.fontWeight = tCount > tableMaxCapacity ? 'bold' : 'normal';
        }
    }
</script>
</body>
</html>
"""

final_html = dnd_html_template.replace("__GUESTS_JSON__", guests_json_payload).replace("__MAX_CAPACITY__", str(table_max_seats))
components.html(final_html, height=1550, scrolling=True)

# ==========================================
# 5. DIEET & ALLERGIE BRIEFING VOOR KEUKEN
# ==========================================
if parsed_guests:
    st.markdown("---")
    st.markdown("### 🍽️ Keuken & Allergieën Briefing")
    
    diet_attendees = [g for g in parsed_guests if g.get('dietary')]
    if diet_attendees:
        df_diet = pd.DataFrame([{
            'Gast': g['name'],
            'Commissie': g['committee'] or "Gast",
            'Dieet / Allergie': g['dietary'],
            'Voorkeur 1': g['pref1'],
            'Voorkeur 2': g['pref2']
        } for g in diet_attendees])
        st.dataframe(df_diet, use_container_width=True, hide_index=True)
    else:
        st.info("Geen dieetwensen gemeld in dit bestand.")
