import io
import pandas as pd
import streamlit as st
from ortools.sat.python import cp_model

# ==========================================
# 1. LUXE GALA HUISSTIJL & THEMA
# ==========================================
st.set_page_config(
    page_title="Ultimate Ball Committee Seating Engine",
    page_icon="🍾",
    layout="wide"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@500;700&family=Montserrat:wght@300;400;600&display=swap');

    .stApp {
        background-color: #FAFAFA;
        font-family: 'Montserrat', sans-serif;
    }
    
    h1, h2, h3, .gala-title {
        font-family: 'Cinzel', serif !important;
        letter-spacing: 2px;
        color: #111111;
        text-align: center;
    }

    .stButton>button {
        background-color: #111111 !important;
        color: #FFFFFF !important;
        border: 1px solid #111111 !important;
        border-radius: 0px !important;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        font-family: 'Cinzel', serif !important;
        padding: 0.6rem 2rem !important;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #FFFFFF !important;
        color: #111111 !important;
        border: 1px solid #111111 !important;
    }

    .table-card {
        background-color: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-top: 3px solid #111111;
        border-radius: 4px;
        padding: 20px;
        margin-bottom: 24px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
    }
    .table-header {
        font-family: 'Cinzel', serif;
        font-size: 1.15rem;
        font-weight: 700;
        letter-spacing: 1px;
        color: #111111;
        border-bottom: 1px solid #111111;
        padding-bottom: 8px;
        margin-bottom: 14px;
        display: flex;
        justify-content: space-between;
    }
    .seat-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 6px 0;
        border-bottom: 1px dotted #E5E7EB;
        font-size: 0.88rem;
    }
    .seat-number {
        font-weight: 600;
        color: #6B7280;
        margin-right: 8px;
    }
    .tag-comm { 
        color: #111111; 
        font-weight: 600; 
        background: #F3F4F6;
        padding: 2px 8px;
        font-size: 0.75rem;
        letter-spacing: 0.5px;
        text-transform: uppercase;
    }
    .tag-indiv { 
        color: #9CA3AF; 
        font-style: italic; 
        font-size: 0.8rem;
    }
    .diet-badge {
        font-size: 0.75rem;
        background: #FEF3C7;
        color: #92400E;
        padding: 1px 6px;
        border-radius: 3px;
        margin-left: 6px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. HEADER MET LOGO
# ==========================================
col_spacer_l, col_center, col_spacer_r = st.columns([1.5, 2, 1.5])
with col_center:
    try:
        st.image("ball_logo.png", use_container_width=True)
    except Exception:
        st.markdown("<h2 class='gala-title'>BALL COMMITTEE</h2>", unsafe_allow_html=True)

st.markdown("<p style='text-align: center; letter-spacing: 3px; color: #6B7280; text-transform: uppercase; font-size: 0.85rem; margin-top: -10px;'>Banquet Seating Arrangement Engine</p>", unsafe_allow_html=True)
st.markdown("<hr style='border: none; border-top: 1px solid #E5E7EB; margin-bottom: 2rem;'>", unsafe_allow_html=True)

# ==========================================
# 3. ZIJBALK INSTELLINGEN
# ==========================================
with st.sidebar:
    try:
        st.image("ball_logo.png", use_container_width=True)
    except Exception:
        pass
    st.markdown("### ⚙️ Gala Parameters")
    table_cap = st.slider("Seats per Table", min_value=4, max_value=16, value=10)
    max_time = st.slider("Max Solver Time (sec)", min_value=5, max_value=60, value=25)
    st.markdown("---")
    st.caption("✅ Supports raw **Hitract exports** with dual tickets as well as standard 4-column sheets.")

# ==========================================
# 4. HITRACT EN STANDAARD DATA PARSER
# ==========================================
KNOWN_COMMITTEES = [
    "JSA Board", "Board", "Spring Ball Committee", "Ball Committee", "Winter Banquet",
    "SexIE", "SexK", "Nextstep", "Case Academy", "JSA Masters", "Sports Committee",
    "Marketing Committee", "Education Committee", "Quality Committee"
]

def extract_comm_or_pref(text):
    if not text or pd.isna(text):
        return None, None
    t = str(text).strip()
    if not t:
        return None, None
    for c in KNOWN_COMMITTEES:
        if c.lower() in t.lower():
            return c, None
    return None, t

def parse_uploaded_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    target_sheet = None
    for s in xls.sheet_names:
        if any(w in s.lower() for w in ['ticket', 'attendee', 'guest', '1)', '2)']):
            target_sheet = s
            break
    if not target_sheet:
        target_sheet = xls.sheet_names[0]

    df_raw = pd.read_excel(uploaded_file, sheet_name=target_sheet).dropna(how='all').reset_index(drop=True)

    # Controleer of het een Hitract export is
    is_hitract = 'Name ticket 1 ' in df_raw.columns or 'Förnamn' in df_raw.columns

    parsed = []
    if is_hitract:
        buyer_counts = {}
        for _, row in df_raw.iterrows():
            fname = str(row.get('Förnamn', '')).strip() if pd.notna(row.get('Förnamn')) else ''
            lname = str(row.get('Efternamn', '')).strip() if pd.notna(row.get('Efternamn')) else ''
            sdate = str(row.get('Försäljnignsdatum', '')).strip() if pd.notna(row.get('Försäljnignsdatum')) else ''
            b_key = f"{fname}_{lname}_{sdate}"
            occ = buyer_counts.get(b_key, 0) + 1
            buyer_counts[b_key] = occ

            n1 = str(row.get('Name ticket 1 ', '')).strip() if pd.notna(row.get('Name ticket 1 ')) else f"{fname} {lname}".strip()
            n2 = str(row.get("Second guest's full name (if purchasing a second ticket)", '')).strip() if pd.notna(row.get("Second guest's full name (if purchasing a second ticket)")) else ''

            p1_1 = str(row.get('Seating preference 1: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 1: Full name (First & Last name) or committee name. ')) else ''
            p1_2 = str(row.get('Seating preference 2: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 2: Full name (First & Last name) or committee name. ')) else ''
            diet1 = str(row.get('Dietary requirements (leave blank if none)', '')).strip() if pd.notna(row.get('Dietary requirements (leave blank if none)')) else ''

            p2_1 = str(row.get('Seating preference 1 second guest: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 1 second guest: Full name (First & Last name) or committee name. ')) else ''
            p2_2 = str(row.get('Seating preference 2 second guest: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 2 second guest: Full name (First & Last name) or committee name. ')) else ''
            diet2 = str(row.get('Dietary requirements second guest (leave blank if none)', '')).strip() if pd.notna(row.get('Dietary requirements second guest (leave blank if none)')) else ''

            if occ == 1 or not n2:
                c1_a, pf1_a = extract_comm_or_pref(p1_1)
                c1_b, pf1_b = extract_comm_or_pref(p1_2)
                parsed.append({
                    'Name': n1,
                    'Committee': c1_a or c1_b,
                    'Preference_1': pf1_a if not c1_a else None,
                    'Preference_2': pf1_b if not c1_b else None,
                    'Dietary': diet1
                })
            elif occ == 2 and n2:
                c2_a, pf2_a = extract_comm_or_pref(p2_1)
                c2_b, pf2_b = extract_comm_or_pref(p2_2)
                parsed.append({
                    'Name': n2,
                    'Committee': c2_a or c2_b,
                    'Preference_1': pf2_a if not c2_a else None,
                    'Preference_2': pf2_b if not c2_b else None,
                    'Dietary': diet2
                })
        return pd.DataFrame(parsed)
    else:
        df = df_raw.iloc[:, 0:4].copy()
        df.columns = ['Name', 'Committee', 'Preference_1', 'Preference_2']
        df['Dietary'] = ""
        for col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace({'nan': None, 'None': None, '': None})
        return df.dropna(subset=['Name']).reset_index(drop=True)

# ==========================================
# 5. UPLOAD & OPTIMALISATIE
# ==========================================
uploaded_file = st.file_uploader("Upload Ball Guestlist (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = parse_uploaded_excel(uploaded_file)
    st.info(f"✅ Successfully processed **{len(df)}** attendees from your list.")

    if st.button("Generate Seating Plan"):
        with st.spinner("Calculating optimal table allocation..."):
            attendees = df['Name'].tolist()
            n_attendees = len(attendees)
            n_tables = (n_attendees + table_cap - 1) // table_cap
            att_map = {name.lower(): i for i, name in enumerate(attendees)}

            model = cp_model.CpModel()
            x = {}
            for i in range(n_attendees):
                for t in range(n_tables):
                    x[i, t] = model.NewBoolVar(f'x_{i}_{t}')

            for i in range(n_attendees):
                model.Add(sum(x[i, t] for t in range(n_tables)) == 1)

            for t in range(n_tables):
                model.Add(sum(x[i, t] for i in range(n_attendees)) <= table_cap)

            # Commissies samen vergrendelen
            if 'Committee' in df.columns:
                for _, comm_df in df.dropna(subset=['Committee']).groupby('Committee'):
                    members = [att_map[n.lower()] for n in comm_df['Name'] if n.lower() in att_map]
                    if 1 < len(members) <= table_cap:
                        for m in members[1:]:
                            for t in range(n_tables):
                                model.Add(x[m, t] == x[members[0], t])

            # Voorkeuren maximaliseren (losse gasten)
            score_terms = []
            for _, row in df.iterrows():
                if pd.notna(row['Committee']):
                    continue
                i = att_map[row['Name'].lower()]
                for p_col in ['Preference_1', 'Preference_2']:
                    p = row.get(p_col)
                    if p and str(p).lower() in att_map:
                        j = att_map[str(p).lower()]
                        if i != j:
                            for t in range(n_tables):
                                together = model.NewBoolVar(f'p_{i}_{j}_{t}')
                                model.Add(x[i, t] + x[j, t] == 2).OnlyEnforceIf(together)
                                model.Add(x[i, t] + x[j, t] < 2).OnlyEnforceIf(together.Not())
                                score_terms.append(together * 5)

            model.Maximize(sum(score_terms))

            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = float(max_time)
            status = solver.Solve(model)

            if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                seating = {}
                for i in range(n_attendees):
                    for t in range(n_tables):
                        if solver.Value(x[i, t]) == 1:
                            seating[attendees[i]] = t + 1
                
                df['Assigned_Table'] = df['Name'].map(seating)
                df_sorted = df.sort_values(by=['Assigned_Table', 'Committee', 'Name']).reset_index(drop=True)

                st.markdown("<br><h3 class='gala-title'>Seating Chart Overview</h3>", unsafe_allow_html=True)

                tables = df_sorted.groupby('Assigned_Table')
                cols = st.columns(3)
                for idx, (tbl_num, group) in enumerate(tables):
                    with cols[idx % 3]:
                        st.markdown(f"""
                        <div class="table-card">
                            <div class="table-header">
                                <span>TABLE {tbl_num:02d}</span>
                                <span style="font-size: 0.85rem; font-weight: normal; color: #6B7280;">{len(group)}/{table_cap} Guests</span>
                            </div>
                        """, unsafe_allow_html=True)
                        for seat_idx, (_, r) in enumerate(group.iterrows(), 1):
                            comm_label = f"<span class='tag-comm'>{r['Committee']}</span>" if pd.notna(r['Committee']) else "<span class='tag-indiv'>Guest</span>"
                            diet_label = f"<span class='diet-badge'>{r['Dietary']}</span>" if r['Dietary'] else ""
                            st.markdown(f"""
                            <div class="seat-row">
                                <span><span class="seat-number">{seat_idx:02d}.</span> {r['Name']}{diet_label}</span>
                                {comm_label}
                            </div>
                            """, unsafe_allow_html=True)
                        st.markdown("</div>", unsafe_allow_html=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_sorted.to_excel(writer, sheet_name="Full Seating", index=False)
                    df_sorted.sort_values('Name').to_excel(writer, sheet_name="Hostess Entrance List", index=False)
                
                st.download_button(
                    label="📥 Download Official Seating Plan (Excel)",
                    data=output.getvalue(),
                    file_name="Ball_Committee_Seating_Plan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Could not find a feasible arrangement. Check group sizes against table capacity.")
