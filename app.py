import io
import pandas as pd
import streamlit as st
from ortools.sat.python import cp_model

# ==========================================
# 1. LUXE GALA HUISSTIJL & THEMA
# ==========================================
st.set_page_config(
    page_title="Ball Committee Seating Engine",
    page_icon="🍾",
    layout="wide"
)

# Aangepaste styling afgestemd op het Ball Committee logo
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@500;700&family=Montserrat:wght@300;400;600&display=swap');

    /* Achtergrond en typografie */
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

    /* Knoppen: Minimalistisch zwart/goud */
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

    /* Tafelkaartjes met luxe omlijning */
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
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. HEADER MET HET LOGO
# ==========================================
col_spacer_l, col_center, col_spacer_r = st.columns([1.5, 2, 1.5])
with col_center:
    # Toont het geüploade logo
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
    st.caption("📋 **Excel Format Order:**\n1. Name\n2. Committee\n3. Preference 1\n4. Preference 2")

# ==========================================
# 4. BESTANDSUPLOAD & OPTIMALISATIE
# ==========================================
uploaded_file = st.file_uploader("Upload Ball Guestlist (.xlsx)", type=["xlsx"])

if uploaded_file:
    raw_df = pd.read_excel(uploaded_file)
    df = raw_df.iloc[:, 0:4].copy()
    df.columns = ['Name', 'Committee', 'Preference_1', 'Preference_2']

    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({'nan': None, 'None': None, '': None})
    df = df.dropna(subset=['Name']).reset_index(drop=True)

    st.info(f"Loaded **{len(df)}** attendees. Ready for seat allocation.")

    if st.button("Generate Seating Plan"):
        with st.spinner("Calculating optimal table allocation..."):
            attendees = df['Name'].tolist()
            n_attendees = len(attendees)
            n_tables = (n_attendees + table_cap - 1) // table_cap
            att_map = {name.lower(): i for i, name in enumerate(attendees)}

            # OR-Tools Constraint Model
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

            # Voorkeuren optimaliseren (alleen voor losse gasten)
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

                # Visuele tafels in kaarten
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
                            st.markdown(f"""
                            <div class="seat-row">
                                <span><span class="seat-number">{seat_idx:02d}.</span> {r['Name']}</span>
                                {comm_label}
                            </div>
                            """, unsafe_allow_html=True)
                        st.markdown("</div>", unsafe_allow_html=True)

                # Excel download
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
