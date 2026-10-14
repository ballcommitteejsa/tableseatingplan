import io
import os
import json
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image

st.set_page_config(
    page_title="Ball Committee — Seating Studio",
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
    
    .header-container {
        text-align: center;
        margin-bottom: 15px;
    }
    
    .gala-title {
        font-family: 'Cinzel', serif !important;
        letter-spacing: 4px;
        color: #111111;
        font-size: 1.6rem;
        margin: 0;
        text-transform: uppercase;
    }
    
    .gala-sub {
        font-size: 0.75rem;
        letter-spacing: 2px;
        color: #6B7280;
        text-transform: uppercase;
        margin-top: 4px;
    }

    .control-panel {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 6px;
        padding: 12px 18px;
        margin-bottom: 12px;
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

# Logo weergave
def render_header():
    for logo in ["ball_logo.png", "ball logo high res.jpg", "input_file_0.png"]:
        if os.path.exists(logo):
            st.image(logo, width=240)
            return
    st.markdown("<h1 class='gala-title'>BALL COMMITTEE</h1>", unsafe_allow_html=True)

col_l, col_m, col_r = st.columns([1.5, 1, 1.5])
with col_m:
    render_header()

st.markdown("<p style='text-align: center; letter-spacing: 2px; color: #6B7280; text-transform: uppercase; font-size: 0.75rem; margin-top: -8px; margin-bottom: 14px;'>20-Table Seating Studio</p>", unsafe_allow_html=True)

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

        # Voorkeuren
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
# INTERACTIEVE STUDIO (HTML/JS)
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
    }}
    .btn {{
        background: #374151;
        color: #FFF;
        border: none;
        padding: 5px 10px;
        border-radius: 3px;
        font-size: 0.75rem;
        cursor: pointer;
        font-weight: 500;
        margin-left: 4px;
    }}
    .btn:hover {{ background: #4B5563; }}
    .btn-blue {{ background: #2563EB; }}
    .btn-blue:hover {{ background: #1D4ED8; }}
    .btn-blue:disabled {{ background: #4B5563; opacity: 0.5; cursor: not-allowed; }}
    .btn-green {{ background: #15803D; }}
    .btn-green:hover {{ background: #166534; }}

    .add-bar {{
        background: #FFF;
        border: 1px solid #D1D5DB;
        border-radius: 4px;
        padding: 6px 10px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 6px;
        font-size: 0.75rem;
    }}
    .add-bar input {{
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
    }}
    .pool-container {{
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        min-height: 50px;
        max-height: 180px;
        overflow-y: auto;
        padding: 4px;
        background: #F9FAFB;
        border: 1px dashed #D1D5DB;
        border-radius: 3px;
    }}

    .grid-20 {{
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
        width: 75%;
        padding: 1px 2px;
    }}
    .table-input:focus {{
        border-color: #2563EB;
        background: #FFF;
        outline: none;
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
    .tag-diet {{
        font-size: 0.6rem;
        font-weight: 600;
        color: #92400E;
        background: #FEF3C7;
        padding: 0 3px;
        border-radius: 2px;
        margin-top: 2px;
        width: fit-content;
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
</style>
</head>
<body>

<div class="toolbar">
    <div style="font-weight: 600;">SEATING STUDIO</div>
    <div>
        <span id="select-label" style="color: #9CA3AF; margin-right: 6px;">0 selected</span>
        <button class="btn btn-blue" id="groupBtn" onclick="bundleSelected()" disabled>Group Selected</button>
        <button class="btn btn-green" onclick="exportCards()">Export Excel (Place Cards)</button>
        <button class="btn" onclick="exportPdf()">Save PDF</button>
        <button class="btn" onclick="resetLayout()">Reset Layout</button>
    </div>
</div>

<div class="add-bar">
    <span style="font-weight: 600;">Add Guest:</span>
    <input type="text" id="manualName" placeholder="Full name (required)" style="flex: 2;">
    <input type="text" id="manualComm" placeholder="Committee (optional)" style="flex: 1;">
    <input type="text" id="manualDiet" placeholder="Dietary (optional)" style="flex: 1;">
    <button class="btn" onclick="addGuestManual()" style="margin: 0; background: #111;">Add</button>
</div>

<div class="pool-section">
    <div class="pool-title">
        <span>UNASSIGNED QUEUE (<span id="pool-count">0</span>)</span>
        <span style="font-weight: normal; color: #6B7280;">Drag guests to tables below</span>
    </div>
    <div class="pool-container" id="pool-zone" ondragover="allowDrop(event)" ondrop="handleDrop(event, 'pool')" ondragleave="dragLeave(event)"></div>
</div>

<div class="grid-20" id="tables-grid"></div>

<script>
    const STORAGE_KEY = "BALL_STUDIO_STORAGE_CLEAN";
    const maxCapacity = {table_max_seats};
    const rawGuests = {guests_json};

    let guestMap = {{}};
    rawGuests.forEach(g => guestMap[g.id] = g);

    let selectedIds = new Set();
    let bundles = {{}};
    let tableNames = {{}};
    let draggedId = null;

    for (let i = 1; i <= 20; i++) {{
        tableNames[i] = "TABLE " + (i < 10 ? '0' + i : i);
    }}

    // 20 Tafels opbouwen
    const grid = document.getElementById('tables-grid');
    for (let i = 1; i <= 20; i++) {{
        const col = document.createElement('div');
        col.className = 'table-col';
        col.innerHTML = `
            <div class="table-head">
                <input type="text" class="table-input" id="tbl-input-${{i}}" value="${{tableNames[i]}}" onchange="setTableName(${{i}}, this.value)">
                <span id="t-count-${{i}}" style="color: #6B7280;">0/${{maxCapacity}}</span>
            </div>
            <div class="dropzone" id="zone-table-${{i}}" ondragover="allowDrop(event)" ondrop="handleDrop(event, 'zone-table-${{i}}')" ondragleave="dragLeave(event)"></div>
        `;
        grid.appendChild(col);
    }}

    function setTableName(i, val) {{
        tableNames[i] = val.trim() || ("TABLE " + (i < 10 ? '0' + i : i));
        saveState();
    }}

    function createGuestCard(g) {{
        const div = document.createElement('div');
        div.className = 'guest-card';
        div.id = g.id;
        div.draggable = true;
        div.ondragstart = (e) => dragStart(e, g.id);

        let comm = g.committee ? `<span class="tag-comm">${{g.committee}}</span>` : '';
        let diet = g.dietary ? `<div class="tag-diet">${{g.dietary}}</div>` : '';

        let prefs = [];
        if (g.pref1) prefs.push(g.pref1);
        if (g.pref2) prefs.push(g.pref2);
        let prefText = prefs.length > 0 ? `<div class="pref-text">wants: ${{prefs.join(' & ')}}</div>` : '';

        div.innerHTML = `
            <div class="guest-top">
                <input type="checkbox" onchange="toggleSelect(event, '${{g.id}}')" style="cursor: pointer; margin: 0 4px 0 0;">
                <span class="guest-name" title="${{g.name}}">${{g.name}}</span>
                ${{comm}}
            </div>
            ${{prefText}}
            ${{diet}}
        `;
        return div;
    }}

    function createBundleCard(bId, memberIds) {{
        const div = document.createElement('div');
        div.className = 'bundle-card';
        div.id = bId;
        div.draggable = true;
        div.ondragstart = (e) => dragStart(e, bId);

        let names = memberIds.map(id => guestMap[id] ? guestMap[id].name : id).join(', ');

        div.innerHTML = `
            <div class="bundle-head">
                <span>GROUP (${{memberIds.length}})</span>
                <button class="ungroup-link" onclick="ungroup('${{bId}}')">Ungroup</button>
            </div>
            <div style="font-weight: 500; font-size: 0.68rem; color: #713F12;">${{names}}</div>
        `;
        return div;
    }}

    const poolZone = document.getElementById('pool-zone');

    function saveState() {{
        try {{
            const state = {{
                bundles: bundles,
                tableNames: tableNames,
                custom: Object.values(guestMap).filter(g => g.id.startsWith("manual_")),
                pool: Array.from(poolZone.children).map(c => ({{ type: c.classList.contains('bundle-card') ? 'bundle' : 'guest', id: c.id }})),
                tables: {{}}
            }};
            for (let i = 1; i <= 20; i++) {{
                const z = document.getElementById('zone-table-' + i);
                state.tables[i] = Array.from(z.children).map(c => ({{ type: c.classList.contains('bundle-card') ? 'bundle' : 'guest', id: c.id }}));
            }}
            localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        }} catch(e) {{}}
    }}

    function loadState() {{
        try {{
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) return false;
            const state = JSON.parse(raw);
            bundles = state.bundles || {{}};

            if (state.custom) {{
                state.custom.forEach(g => guestMap[g.id] = g);
            }}

            if (state.tableNames) {{
                tableNames = state.tableNames;
                for (let i = 1; i <= 20; i++) {{
                    const el = document.getElementById('tbl-input-' + i);
                    if (el && tableNames[i]) el.value = tableNames[i];
                }}
            }}

            const placed = new Set();
            poolZone.innerHTML = '';
            for (let i = 1; i <= 20; i++) document.getElementById('zone-table-' + i).innerHTML = '';

            if (state.pool) {{
                state.pool.forEach(item => {{
                    if (item.type === 'bundle' && bundles[item.id]) {{
                        poolZone.appendChild(createBundleCard(item.id, bundles[item.id]));
                        bundles[item.id].forEach(id => placed.add(id));
                    }} else if (item.type === 'guest' && guestMap[item.id]) {{
                        poolZone.appendChild(createGuestCard(guestMap[item.id]));
                        placed.add(item.id);
                    }}
                }});
            }}

            if (state.tables) {{
                for (let i = 1; i <= 20; i++) {{
                    const z = document.getElementById('zone-table-' + i);
                    (state.tables[i] || []).forEach(item => {{
                        if (item.type === 'bundle' && bundles[item.id]) {{
                            z.appendChild(createBundleCard(item.id, bundles[item.id]));
                            bundles[item.id].forEach(id => placed.add(id));
                        }} else if (item.type === 'guest' && guestMap[item.id]) {{
                            z.appendChild(createGuestCard(guestMap[item.id]));
                            placed.add(item.id);
                        }}
                    }});
                }}
            }}

            Object.values(guestMap).forEach(g => {{
                if (!placed.has(g.id)) poolZone.appendChild(createGuestCard(g));
            }});
            return true;
        }} catch(e) {{
            return false;
        }}
    }}

    // Start-initiatie
    if (!loadState()) {{
        Object.values(guestMap).forEach(g => {{
            poolZone.appendChild(createGuestCard(g));
        }});
    }}
    updateCounters();

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

    function toggleSelect(e, id) {{
        e.stopPropagation();
        if (e.target.checked) selectedIds.add(id);
        else selectedIds.delete(id);

        const cnt = selectedIds.size;
        document.getElementById('select-label').innerText = `${{cnt}} selected`;
        document.getElementById('groupBtn').disabled = cnt < 2;
    }}

    function bundleSelected() {{
        if (selectedIds.size < 2) return;
        const bId = 'bundle_' + Date.now();
        const mIds = Array.from(selectedIds);
        bundles[bId] = mIds;

        const first = document.getElementById(mIds[0]);
        const parent = first ? first.parentElement : poolZone;

        mIds.forEach(id => {{
            const el = document.getElementById(id);
            if (el) el.remove();
        }});

        parent.appendChild(createBundleCard(bId, mIds));
        selectedIds.clear();
        document.getElementById('groupBtn').disabled = true;
        document.getElementById('select-label').innerText = '0 selected';
        updateCounters();
        saveState();
    }}

    function ungroup(bId) {{
        const bEl = document.getElementById(bId);
        if (!bEl) return;
        const parent = bEl.parentElement;
        (bundles[bId] || []).forEach(id => {{
            if (guestMap[id]) parent.appendChild(createGuestCard(guestMap[id]));
        }});
        bEl.remove();
        delete bundles[bId];
        updateCounters();
        saveState();
    }}

    function dragStart(e, id) {{
        draggedId = id;
        e.dataTransfer.setData('text/plain', id);
    }}
    function allowDrop(e) {{
        e.preventDefault();
        e.currentTarget.classList.add('dragover');
    }}
    function dragLeave(e) {{
        e.currentTarget.classList.remove('dragover');
    }}
    function handleDrop(e, target) {{
        e.preventDefault();
        e.currentTarget.classList.remove('dragover');
        if (!draggedId) return;
        const el = document.getElementById(draggedId);
        if (el) {{
            e.currentTarget.appendChild(el);
            updateCounters();
            saveState();
        }}
    }}

    function updateCounters() {{
        let pCount = 0;
        Array.from(poolZone.children).forEach(c => {{
            pCount += c.classList.contains('bundle-card') ? (bundles[c.id] || []).length : 1;
        }});
        document.getElementById('pool-count').innerText = pCount;

        for (let i = 1; i <= 20; i++) {{
            const z = document.getElementById('zone-table-' + i);
            let count = 0;
            Array.from(z.children).forEach(c => {{
                count += c.classList.contains('bundle-card') ? (bundles[c.id] || []).length : 1;
            }});
            const counter = document.getElementById('t-count-' + i);
            counter.innerText = `${{count}}/${{maxCapacity}}`;
            counter.style.color = count > maxCapacity ? '#DC2626' : (count === maxCapacity ? '#16A34A' : '#6B7280');
        }}
    }}

    function resetLayout() {{
        if (confirm("Reset all seating positions?")) {{
            localStorage.removeItem(STORAGE_KEY);
            location.reload();
        }}
    }}

    function exportCards() {{
        saveState();
        const rows = [["Table", "Seat", "Guest Name", "Committee", "Dietary", "Preferences"]];
        for (let i = 1; i <= 20; i++) {{
            const tName = tableNames[i] || ("TABLE " + (i < 10 ? '0' + i : i));
            const z = document.getElementById('zone-table-' + i);
            let seat = 1;
            Array.from(z.children).forEach(child => {{
                const ids = child.classList.contains('bundle-card') ? (bundles[child.id] || []) : [child.id];
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
        const csv = "\\uFEFF" + rows.map(r => r.join(";")).join("\\r\\n");
        const blob = new Blob([csv], {{ type: "text/csv;charset=utf-8;" }});
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "Guestlist_Place_Cards_Sorted.csv";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }}

    function exportPdf() {{
        saveState();
        let doc = `
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="utf-8">
        <title>Seating Plan</title>
        <style>
            @page {{ size: A4 landscape; margin: 8mm; }}
            body {{ font-family: sans-serif; margin: 0; padding: 10px; color: #111; }}
            h2 {{ text-align: center; margin: 0 0 10px 0; font-size: 14pt; }}
            .grid {{ display: grid; grid-template-columns: repeat(5, 1fr); gap: 8px; }}
            .box {{ border: 1px solid #111; padding: 5px; page-break-inside: avoid; }}
            .title {{ font-weight: bold; font-size: 8pt; border-bottom: 1px solid #111; padding-bottom: 2px; margin-bottom: 4px; display: flex; justify-content: space-between; }}
            .row {{ font-size: 7pt; padding: 2px 0; border-bottom: 1px dotted #ccc; }}
        </style>
        </head>
        <body>
        <h2>BALL COMMITTEE — SEATING PLAN</h2>
        <div class="grid">
        `;
        for (let i = 1; i <= 20; i++) {{
            const tName = tableNames[i] || ("TABLE " + (i < 10 ? '0' + i : i));
            const z = document.getElementById('zone-table-' + i);
            const cnt = document.getElementById('t-count-' + i).innerText;
            doc += `<div class="box"><div class="title"><span>${{tName}}</span><span>${{cnt}}</span></div>`;
            let seat = 1;
            Array.from(z.children).forEach(child => {{
                const ids = child.classList.contains('bundle-card') ? (bundles[child.id] || []) : [child.id];
                ids.forEach(id => {{
                    const g = guestMap[id];
                    if (g) {{
                        doc += `<div class="row"><b>${{seat++}}.</b> ${{g.name}} ${{g.committee ? '['+g.committee+']' : ''}} ${{g.dietary ? '('+g.dietary+')' : ''}}</div>`;
                    }}
                }});
            }});
            doc += `</div>`;
        }}
        doc += `</div></body></html>`;

        const w = window.open('', '_blank');
        if (w) {{
            w.document.write(doc);
            w.document.close();
            setTimeout(() => w.print(), 300);
        }}
    }}
</script>
</body>
</html>
"""

components.html(html_code, height=1350, scrolling=True)

# Keukenoverzicht onderaan
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
