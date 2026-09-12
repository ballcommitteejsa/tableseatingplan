import io
import pandas as pd
import streamlit as st
from ortools.sat.python import cp_model

# ==========================================
# 1. HUISSTIJL & THEMA (Aanpasbaar)
# ==========================================
st.set_page_config(
    page_title="JIBS Gala Seating Optimizer",
    page_icon="🍸",
    layout="wide"
)

# Aangepaste CSS voor huisstijl (bijv. Navy/Goud gala-look)
st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    .stButton>button {
        background-color: #1a365d;
        color: white;
        border-radius: 8px;
        font-weight: bold;
        height: 3em;
        width: 100%;
    }
    .table-card {
        background-color: white;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .table-header {
        font-size: 1.2rem;
        font-weight: bold;
        color: #1a365d;
        border-bottom: 2px solid #ecc94b;
        padding-bottom: 8px;
        margin-bottom: 12px;
    }
    .seat-row {
        display: flex;
        justify-content: space-between;
        padding: 4px 0;
        border-bottom: 1px dashed #edf2f7;
        font-size: 0.95rem;
    }
    .tag-comm { color: #2b6cb0; font-weight: 600; }
    .tag-indiv { color: #718096; font-style: italic; }
</style>
""", unsafe_allow_html=True)

# Header
st.title("🍸 Gala Table Seating Optimizer")
st.markdown("Upload the registration sheet to automatically generate optimal table arrangements.")

# ==========================================
# 2. ZIJBALK: INSTELLINGEN
# ==========================================
with st.sidebar:
    st.header("⚙️ Event Settings")
    table_cap = st.slider("Seats per Table", min_value=4, max_value=16, value=10)
    max_time = st.slider("Max Solver Time (seconds)", min_value=5, max_value=60, value=20)
    st.info("💡 **Required Columns (in order):**\n1. Name\n2. Committee\n3. Preference 1\n4. Preference 2")

# ==========================================
# 3. BESTANDSUPLOAD & VERWERKING
# ==========================================
uploaded_file = st.file_uploader("Upload Attendees Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Inlezen en opschonen
    raw_df = pd.read_excel(uploaded_file)
    df = raw_df.iloc[:, 0:4].copy()
    df.columns = ['Name', 'Committee', 'Preference_1', 'Preference_2']

    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({'nan': None, 'None': None, '': None})
    df = df.dropna(subset=['Name']).reset_index(drop=True)

    st.success(f"✅ Loaded {len(df)} attendees successfully.")

    if st.button("🚀 Generate Optimal Seating Arrangement"):
        with st.spinner("Optimizing seating chart..."):
            attendees = df['Name'].tolist()
            n_attendees = len(attendees)
            n_tables = (n_attendees + table_cap - 1) // table_cap
            att_map = {name.lower(): i for i, name in enumerate(attendees)}

            # OR-Tools Model
            model = cp_model.CpModel()
            x = {}
            for i in range(n_attendees):
                for t in range(n_tables):
                    x[i, t] = model.NewBoolVar(f'x_{i}_{t}')

            for i in range(n_attendees):
                model.Add(sum(x[i, t] for t in range(n_tables)) == 1)

            for t in range(n_tables):
                model.Add(sum(x[i, t] for i in range(n_attendees)) <= table_cap)

            # Commissie constraint
            if 'Committee' in df.columns:
                for _, comm_df in df.dropna(subset=['Committee']).groupby('Committee'):
                    members = [att_map[n.lower()] for n in comm_df['Name'] if n.lower() in att_map]
                    if 1 < len(members) <= table_cap:
                        for m in members[1:]:
                            for t in range(n_tables):
                                model.Add(x[m, t] == x[members[0], t])

            # Voorkeuren maximaliseren (alleen niet-commissieleden)
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

                st.markdown("### 🍽️ Generated Seating Chart")

                # Visuele weergave in 3 kolommen
                tables = df_sorted.groupby('Assigned_Table')
                cols = st.columns(3)
                for idx, (tbl_num, group) in enumerate(tables):
                    with cols[idx % 3]:
                        st.markdown(f"""
                        <div class="table-card">
                            <div class="table-header">Table {tbl_num:02d} ({len(group)}/{table_cap})</div>
                        """, unsafe_allow_html=True)
                        for seat_idx, (_, r) in enumerate(group.iterrows(), 1):
                            comm_label = f"<span class='tag-comm'>[{r['Committee']}]</span>" if pd.notna(r['Committee']) else "<span class='tag-indiv'>Individual</span>"
                            st.markdown(f"""
                            <div class="seat-row">
                                <span><b>#{seat_idx:02d}</b> {r['Name']}</span>
                                {comm_label}
                            </div>
                            """, unsafe_allow_html=True)
                        st.markdown("</div>", unsafe_allow_html=True)

                # Excel download genereren
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_sorted.to_excel(writer, sheet_name="Full Seating", index=False)
                    df_sorted.sort_values('Name').to_excel(writer, sheet_name="Hostess List", index=False)
                
                st.download_button(
                    label="📥 Download Excel Seating Chart",
                    data=output.getvalue(),
                    file_name="Gala_Seating_Plan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("Could not find a valid seating arrangement. Check group sizes vs table capacity.")
