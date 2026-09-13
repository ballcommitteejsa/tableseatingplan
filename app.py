import io
import os
import time
import pandas as pd
import streamlit as st
from ortools.sat.python import cp_model
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
# 1. PAGE CONFIGURATION & LOGO FAVICON
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
    page_title="Ball Committee Seating Engine",
    page_icon=get_page_icon(),
    layout="wide"
)

custom_css = """
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
        font-weight: 600;
        padding: 2px 6px;
        border-radius: 3px;
        margin-left: 6px;
        border: 1px solid #FDE68A;
    }
    .diet-banner {
        background-color: #FFFBEB;
        border-left: 4px solid #F59E0B;
        padding: 12px 16px;
        border-radius: 4px;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 4px;
        padding: 14px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==========================================
# 2. LOGO & HEADER
# ==========================================
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

st.markdown("<p style='text-align: center; letter-spacing: 3px; color: #6B7280; text-transform: uppercase; font-size: 0.85rem; margin-top: -10px;'>Universal Seating Arrangement & Staff Engine</p>", unsafe_allow_html=True)
st.markdown("<hr style='border: none; border-top: 1px solid #E5E7EB; margin-bottom: 2rem;'>", unsafe_allow_html=True)

# ==========================================
# 3. ZIJBALK INSTELLINGEN
# ==========================================
with st.sidebar:
    render_ball_logo()
    st.markdown("### ⚙️ Gala Parameters")
    table_cap = st.number_input("Seats per Table", min_value=2, max_value=30, value=10, step=1)
    st.markdown("---")
    st.caption("✅ Supports **Spring Inspiration** & **Ball Committee** as distinct separate committees.")

# ==========================================
# 4. UNIVERSELE DYNAMISCHE DATA PARSER
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

def resolve_committee_alias(text):
    if not text or pd.isna(text):
        return None
    raw = str(text).strip().lower()
    if not raw or raw in ['-', 'none', 'nej', 'no', 'vet inte', 'ingen', 'x']:
        return None
    
    for canonical, variations in COMMITTEE_ALIASES.items():
        for var in variations:
            if raw == var or f" {var} " in f" {raw} " or raw.startswith(f"{var} ") or raw.endswith(f" {var}"):
                return canonical
    return None

def parse_uploaded_excel(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    
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
        df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet_name).dropna(how='all')
        if len(df_sheet) > 0:
            dfs.append(df_sheet)

    df_raw = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    is_hitract = 'Name ticket 1 ' in df_raw.columns or 'Förnamn' in df_raw.columns

    if is_hitract:
        all_guest_names = set()
        for _, row in df_raw.iterrows():
            fn = str(row.get('Förnamn', '')).strip() if pd.notna(row.get('Förnamn')) else ''
            ln = str(row.get('Efternamn', '')).strip() if pd.notna(row.get('Efternamn')) else ''
            n1 = str(row.get('Name ticket 1 ', '')).strip() if pd.notna(row.get('Name ticket 1 ')) else f"{fn} {ln}".strip()
            n2 = str(row.get("Second guest's full name (if purchasing a second ticket)", '')).strip() if pd.notna(row.get("Second guest's full name (if purchasing a second ticket)")) else ''
            if n1: all_guest_names.add(n1.lower())
            if n2: all_guest_names.add(n2.lower())

        buyer_counts = {}
        parsed = []
        for _, row in df_raw.iterrows():
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

            c1_raw = str(row.get('Committee table', '')).strip() if pd.notna(row.get('Committee table')) else ''
            p1_1 = str(row.get('Seating preference 1: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 1: Full name (First & Last name) or committee name. ')) else ''
            p1_2 = str(row.get('Seating preference 2: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 2: Full name (First & Last name) or committee name. ')) else ''
            diet1 = str(row.get('Dietary requirements (leave blank if none)', '')).strip() if pd.notna(row.get('Dietary requirements (leave blank if none)')) else ''

            c2_raw = str(row.get('Committee table.1', '')).strip() if pd.notna(row.get('Committee table.1')) else ''
            p2_1 = str(row.get('Seating preference 1 second guest: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 1 second guest: Full name (First & Last name) or committee name. ')) else ''
            p2_2 = str(row.get('Seating preference 2 second guest: Full name (First & Last name) or committee name. ', '')).strip() if pd.notna(row.get('Seating preference 2 second guest: Full name (First & Last name) or committee name. ')) else ''
            diet2 = str(row.get('Dietary requirements second guest (leave blank if none)', '')).strip() if pd.notna(row.get('Dietary requirements second guest (leave blank if none)')) else ''

            if occ == 1 or not n2:
                alias_comm = resolve_committee_alias(c1_raw) or resolve_committee_alias(p1_1) or resolve_committee_alias(p1_2)
                comm = alias_comm or c1_raw or (p1_1 if p1_1 and p1_1.lower() not in all_guest_names else None) or (p1_2 if p1_2 and p1_2.lower() not in all_guest_names else None)
                parsed.append({
                    'Name': n1,
                    'Committee': comm if comm else None,
                    'Preference_1': p1_1 if p1_1 and not alias_comm else None,
                    'Preference_2': p1_2 if p1_2 and not alias_comm else None,
                    'Dietary': diet1
                })
            else:
                alias_comm = resolve_committee_alias(c2_raw) or resolve_committee_alias(p2_1) or resolve_committee_alias(p2_2)
                comm = alias_comm or c2_raw or (p2_1 if p2_1 and p2_1.lower() not in all_guest_names else None) or (p2_2 if p2_2 and p2_2.lower() not in all_guest_names else None)
                parsed.append({
                    'Name': n2,
                    'Committee': comm if comm else None,
                    'Preference_1': p2_1 if p2_1 and not alias_comm else n1,
                    'Preference_2': p2_2 if p2_2 and not alias_comm else None,
                    'Dietary': diet2
                })
        return pd.DataFrame(parsed)
    else:
        df = df_raw.iloc[:, 0:4].copy()
        df.columns = ['Name', 'Committee', 'Preference_1', 'Preference_2']
        df['Dietary'] = df_raw.iloc[:, 4].fillna("").astype(str).str.strip() if df_raw.shape[1] > 4 else ""
        
        for idx, row in df.iterrows():
            c_alias = resolve_committee_alias(row['Committee']) or resolve_committee_alias(row['Preference_1']) or resolve_committee_alias(row['Preference_2'])
            if c_alias:
                df.at[idx, 'Committee'] = c_alias
                if resolve_committee_alias(row['Preference_1']): df.at[idx, 'Preference_1'] = None
                if resolve_committee_alias(row['Preference_2']): df.at[idx, 'Preference_2'] = None

        for col in ['Name', 'Committee', 'Preference_1', 'Preference_2']:
            df[col] = df[col].astype(str).str.strip().replace({'nan': None, 'None': None, '': None})
        return df.dropna(subset=['Name']).reset_index(drop=True)

# ==========================================
# 5. GENERATE 2-PAGE A4 PDF
# ==========================================
def generate_2page_visual_pdf(df_sorted, table_cap):
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(A4),
        leftMargin=12,
        rightMargin=12,
        topMargin=12,
        bottomMargin=12
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, leading=16, alignment=1, textColor=colors.HexColor('#111111')
    )
    subtitle_style = ParagraphStyle(
        'SubTitleStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=10, alignment=1, textColor=colors.HexColor('#666666')
    )
    table_header_style = ParagraphStyle(
        'TblHdr', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.0, leading=11, textColor=colors.HexColor('#111111')
    )
    seat_style = ParagraphStyle(
        'SeatText', parent=styles['Normal'], fontName='Helvetica', fontSize=7.4, leading=9.0, textColor=colors.HexColor('#222222')
    )
    comm_style = ParagraphStyle(
        'CommText', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.8, leading=8.5, alignment=2, textColor=colors.HexColor('#1A365D')
    )
    
    story = []
    tables = list(df_sorted.groupby('Assigned_Table'))
    total_tables = len(tables)
    mid_point = (total_tables + 1) // 2
    
    table_cards = []
    for tbl_num, grp in tables:
        card_data = []
        hdr_cell = Paragraph(f"<b>TABLE {tbl_num:02d}</b> ({len(grp)}/{table_cap})", table_header_style)
        card_data.append([hdr_cell, ""])
        
        for seat_idx, (_, r) in enumerate(grp.iterrows(), 1):
            name_text = f"<b>{seat_idx:02d}.</b> {r['Name']}"
            if r.get('Dietary'):
                name_text += f" <font color='#B45309'>[{r['Dietary']}]</font>"
            left_p = Paragraph(name_text, seat_style)
            comm_text = r['Committee'] if pd.notna(r['Committee']) else ""
            right_p = Paragraph(comm_text, comm_style)
            card_data.append([left_p, right_p])
            
        card_table = Table(card_data, colWidths=[114, 46])
        card_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F3F4F6')),
            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.HexColor('#111111')),
            ('TOPPADDING', (0, 0), (-1, -1), 1),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
            ('LEFTPADDING', (0, 0), (-1, -1), 2.5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2.5),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1D5DB')),
            ('LINEBELOW', (0, 1), (-1, -1), 0.3, colors.HexColor('#F3F4F6')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        table_cards.append(card_table)

    def make_grid(cards_subset, page_num):
        grid_data = []
        t_par = Paragraph("<b>BALL COMMITTEE — BANQUET SEATING PLAN</b>", title_style)
        sub_par = Paragraph(f"Official Banquet Seating Layout • Page {page_num} of 2", subtitle_style)
        ncols = 5
        for r_idx in range(0, len(cards_subset), ncols):
            row_cards = cards_subset[r_idx : r_idx + ncols]
            while len(row_cards) < ncols:
                row_cards.append("")
            grid_data.append(row_cards)
            
        col_w = 163
        grid_table = Table(grid_data, colWidths=[col_w]*ncols)
        grid_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 1),
            ('RIGHTPADDING', (0, 0), (-1, -1), 1),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        return [t_par, sub_par, Spacer(1, 4), grid_table]

    story.extend(make_grid(table_cards[:mid_point], 1))
    if len(table_cards) > mid_point:
        story.append(PageBreak())
        story.extend(make_grid(table_cards[mid_point:], 2))

    doc.build(story)
    return pdf_buffer.getvalue()

# ==========================================
# 6. PRINTKLAAR EXCEL DOCUMENT MAKEN (A4)
# ==========================================
def build_printable_excel_workbook(df_sorted, unfulfilled_df):
    wb = openpyxl.Workbook()
    
    # TAB 1: KITCHEN STAFF BRIEFING (A4)
    ws_kitchen = wb.active
    ws_kitchen.title = "Kitchen Dietary Briefing (Print)"
    ws_kitchen.page_setup.orientation = ws_kitchen.ORIENTATION_PORTRAIT
    ws_kitchen.page_setup.paperSize = ws_kitchen.PAPERSIZE_A4
    ws_kitchen.page_setup.fitToWidth = 1
    ws_kitchen.page_setup.fitToHeight = 0
    ws_kitchen.sheet_properties.pageSetUpPr.fitToPage = True

    ws_kitchen.merge_cells('A1:D1')
    ws_kitchen['A1'] = "BALL COMMITTEE — KITCHEN & SERVICE DIETARY BRIEFING"
    ws_kitchen['A1'].font = Font(name='Calibri', size=14, bold=True, color='111111')
    ws_kitchen['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws_kitchen.merge_cells('A2:D2')
    ws_kitchen['A2'] = "Printable summary of all allergies, dietary restrictions, and special meals per table."
    ws_kitchen['A2'].font = Font(name='Calibri', size=10, italic=True, color='666666')
    ws_kitchen['A2'].alignment = Alignment(horizontal='center')

    k_headers = ["Table", "Seat", "Guest Name", "Dietary Requirement / Allergy"]
    ws_kitchen.append([])
    ws_kitchen.append(k_headers)

    h_fill = PatternFill(start_color="111111", end_color="111111", fill_type="solid")
    h_font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='DDDDDD'), right=Side(style='thin', color='DDDDDD'),
        top=Side(style='thin', color='DDDDDD'), bottom=Side(style='thin', color='DDDDDD')
    )

    for col_idx in range(1, 5):
        cell = ws_kitchen.cell(row=4, column=col_idx)
        cell.fill = h_fill
        cell.font = h_font
        cell.alignment = Alignment(horizontal='left' if col_idx > 2 else 'center', vertical='center')

    row_num = 5
    for tbl_num, grp in df_sorted.groupby('Assigned_Table'):
        diet_in_tbl = grp[grp['Dietary'].fillna('').str.strip() != '']
        if len(diet_in_tbl) > 0:
            for seat_idx, (_, r) in enumerate(grp.iterrows(), 1):
                if str(r['Dietary']).strip():
                    ws_kitchen.cell(row=row_num, column=1, value=f"Table {tbl_num:02d}").alignment = Alignment(horizontal='center')
                    ws_kitchen.cell(row=row_num, column=2, value=f"Seat {seat_idx:02d}").alignment = Alignment(horizontal='center')
                    ws_kitchen.cell(row=row_num, column=3, value=r['Name'])
                    d_cell = ws_kitchen.cell(row=row_num, column=4, value=r['Dietary'])
                    d_cell.font = Font(name='Calibri', size=10, bold=True, color='92400E')
                    d_cell.fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
                    for c_idx in range(1, 5):
                        ws_kitchen.cell(row=row_num, column=c_idx).border = thin_border
                    row_num += 1

    for col in ws_kitchen.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_kitchen.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # TAB 2: VISUAL SEATING LAYOUT
    ws_visual = wb.create_sheet(title="Visual Seating Layout")
    ws_visual.page_setup.orientation = ws_visual.ORIENTATION_PORTRAIT
    ws_visual.page_setup.paperSize = ws_visual.PAPERSIZE_A4

    ws_visual.merge_cells('A1:C1')
    ws_visual['A1'] = "OFFICIAL BANQUET SEATING CHART"
    ws_visual['A1'].font = Font(name='Calibri', size=14, bold=True, color='111111')

    v_headers = ["Table / Seat", "Guest Name", "Committee / Dietary Note"]
    ws_visual.append([])
    ws_visual.append(v_headers)
    for col_idx in range(1, 4):
        c = ws_visual.cell(row=3, column=col_idx)
        c.fill = h_fill
        c.font = h_font

    v_row = 4
    for tbl_num, grp in df_sorted.groupby('Assigned_Table'):
        ws_visual.cell(row=v_row, column=1, value=f"=== TABLE {tbl_num:02d} ===").font = Font(name='Calibri', bold=True)
        v_row += 1
        for seat_idx, (_, r) in enumerate(grp.iterrows(), 1):
            comm = f"[{r['Committee']}]" if pd.notna(r['Committee']) else ""
            diet = f"({r['Dietary']})" if r['Dietary'] else ""
            note = f"{comm} {diet}".strip() or "Guest"

            ws_visual.cell(row=v_row, column=1, value=f"Seat {seat_idx:02d}")
            ws_visual.cell(row=v_row, column=2, value=r['Name'])
            ws_visual.cell(row=v_row, column=3, value=note)
            for c_idx in range(1, 4):
                ws_visual.cell(row=v_row, column=c_idx).border = thin_border
            v_row += 1
        v_row += 1

    for col in ws_visual.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_visual.column_dimensions[col_letter].width = max(max_len + 4, 16)

    # TAB 3: UNFULFILLED WISHES REPORT
    ws_unfulfilled = wb.create_sheet(title="Unfulfilled Wishes Report")
    u_headers = ["Guest Name", "Assigned Table", "Requested Wish", "Requested Guest Table", "Status / Reason"]
    ws_unfulfilled.append(u_headers)
    for c_idx in range(1, 6):
        c = ws_unfulfilled.cell(row=1, column=c_idx)
        c.fill = h_fill
        c.font = h_font

    for _, r in unfulfilled_df.iterrows():
        ws_unfulfilled.append([r['Guest Name'], r['Assigned Table'], r['Requested Wish'], r['Requested Guest Table'], r['Status / Reason']])

    for col in ws_unfulfilled.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_unfulfilled.column_dimensions[col_letter].width = max(max_len + 4, 16)

    # TAB 4: HOSTESS ENTRANCE LIST
    ws_hostess = wb.create_sheet(title="Hostess Entrance List")
    h_headers = ["Guest Name", "Table", "Committee", "Dietary Requirement"]
    ws_hostess.append(h_headers)
    for c_idx in range(1, 5):
        c = ws_hostess.cell(row=1, column=c_idx)
        c.fill = h_fill
        c.font = h_font

    for _, r in df_sorted.sort_values('Name').iterrows():
        ws_hostess.append([r['Name'], f"Table {r['Assigned_Table']:02d}", r['Committee'] or "", r['Dietary'] or ""])

    for col in ws_hostess.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_hostess.column_dimensions[col_letter].width = max(max_len + 4, 14)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 7. APP UI & OPTIMALISATIE ENGINE
# ==========================================
uploaded_file = st.file_uploader("Upload Ball Guestlist (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = parse_uploaded_excel(uploaded_file)
    detected_comms = df['Committee'].dropna().unique()
    comm_txt = f"Detected {len(detected_comms)} committees: *{', '.join(detected_comms)}*" if len(detected_comms) > 0 else "Individual guestlist"
    st.info(f"✅ Processed **{len(df)}** attendees. {comm_txt}")

    if st.button("Generate Seating Plan"):
        progress_placeholder = st.empty()
        progress_bar = progress_placeholder.progress(0, text="Initializing optimization engine... 0%")
        
        for pct in range(5, 45, 5):
            time.sleep(0.04)
            progress_bar.progress(pct, text=f"Analyzing committee constraints & seating wishes... {pct}%")

        attendees = df['Name'].tolist()
        n_attendees = len(attendees)
        table_capacity_val = int(table_cap)
        n_tables = (n_attendees + table_capacity_val - 1) // table_capacity_val
        att_map = {name.lower(): i for i, name in enumerate(attendees)}

        model = cp_model.CpModel()
        x = {}
        for i in range(n_attendees):
            for t in range(n_tables):
                x[i, t] = model.NewBoolVar(f'x_{i}_{t}')

            model.Add(sum(x[i, t] for t in range(n_tables)) == 1)

        for t in range(n_tables):
            model.Add(sum(x[i, t] for i in range(n_attendees)) <= table_capacity_val)

        comm_members_map = {}
        if 'Committee' in df.columns:
            for comm_name, comm_df in df.dropna(subset=['Committee']).groupby('Committee'):
                c_members = [att_map[n.lower()] for n in comm_df['Name'] if n.lower() in att_map]
                comm_members_map[str(comm_name).lower()] = c_members
                if 1 < len(c_members) <= table_capacity_val:
                    for m in c_members[1:]:
                        for t in range(n_tables):
                            model.Add(x[m, t] == x[c_members[0], t])

        score_terms = []
        for _, row in df.iterrows():
            if pd.notna(row['Committee']):
                continue
            i = att_map[row['Name'].lower()]
            for p_col in ['Preference_1', 'Preference_2']:
                pref_raw = row.get(p_col)
                if not pref_raw or pd.isna(pref_raw):
                    continue
                pref_clean = str(pref_raw).strip().lower()

                if pref_clean in att_map:
                    j = att_map[pref_clean]
                    if i != j:
                        for t in range(n_tables):
                            together = model.NewBoolVar(f'p_pers_{i}_{j}_{t}')
                            model.Add(x[i, t] + x[j, t] == 2).OnlyEnforceIf(together)
                            model.Add(x[i, t] + x[j, t] < 2).OnlyEnforceIf(together.Not())
                            score_terms.append(together * 6)

                for c_name, members in comm_members_map.items():
                    if c_name in pref_clean or pref_clean in c_name:
                        for target_m in members:
                            for t in range(n_tables):
                                together_c = model.NewBoolVar(f'p_comm_{i}_{target_m}_{t}')
                                model.Add(x[i, t] + x[target_m, t] == 2).OnlyEnforceIf(together_c)
                                model.Add(x[i, t] + x[target_m, t] < 2).OnlyEnforceIf(together_c.Not())
                                score_terms.append(together_c * 3)

        model.Maximize(sum(score_terms))

        for pct in range(45, 80, 5):
            time.sleep(0.03)
            progress_bar.progress(pct, text=f"Solving table arrangement matrix... {pct}%")

        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 10.0
        status = solver.Solve(model)

        for pct in range(80, 101, 5):
            time.sleep(0.02)
            progress_bar.progress(pct, text=f"Finalizing report & layouts... {pct}%")
            
        time.sleep(0.15)
        progress_placeholder.empty()

        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            seating = {}
            for i in range(n_attendees):
                for t in range(n_tables):
                    if solver.Value(x[i, t]) == 1:
                        seating[attendees[i]] = t + 1
            
            df['Assigned_Table'] = df['Name'].map(seating)
            df_sorted = df.sort_values(by=['Assigned_Table', 'Committee', 'Name']).reset_index(drop=True)
            seating_map_lower = {name.lower(): tbl for name, tbl in seating.items()}

            unfulfilled_rows = []
            total_wishes = 0
            fulfilled_wishes = 0

            for _, r in df.iterrows():
                if pd.notna(r['Committee']):
                    continue
                g_name = r['Name']
                g_table = r['Assigned_Table']

                for p_col in ['Preference_1', 'Preference_2']:
                    pref_val = r.get(p_col)
                    if pd.isna(pref_val) or not str(pref_val).strip():
                        continue
                    pref_str = str(pref_val).strip()
                    total_wishes += 1

                    if pref_str.lower() in seating_map_lower:
                        target_t = seating_map_lower[pref_str.lower()]
                        if target_t == g_table:
                            fulfilled_wishes += 1
                        else:
                            unfulfilled_rows.append({
                                'Guest Name': g_name,
                                'Assigned Table': f"Table {g_table:02d}",
                                'Requested Wish': pref_str,
                                'Requested Guest Table': f"Table {target_t:02d}",
                                'Status / Reason': 'Different table due to table capacity limits'
                            })
                    else:
                        matched_comm = False
                        for c_name, members in comm_members_map.items():
                            if c_name in pref_str.lower() or pref_str.lower() in c_name:
                                if members:
                                    first_member_name = attendees[members[0]]
                                    comm_t = seating[first_member_name]
                                    if comm_t == g_table:
                                        fulfilled_wishes += 1
                                        matched_comm = True
                                        break
                                    else:
                                        unfulfilled_rows.append({
                                            'Guest Name': g_name,
                                            'Assigned Table': f"Table {g_table:02d}",
                                            'Requested Wish': f"[{pref_str}] Committee Table",
                                            'Requested Guest Table': f"Table {comm_t:02d}",
                                            'Status / Reason': 'Committee table reached capacity'
                                        })
                                        matched_comm = True
                                        break
                        if not matched_comm:
                            unfulfilled_rows.append({
                                'Guest Name': g_name,
                                'Assigned Table': f"Table {g_table:02d}",
                                'Requested Wish': pref_str,
                                'Requested Guest Table': 'Not Found',
                                'Status / Reason': 'Name / Committee not registered in guestlist'
                            })

            unfulfilled_df = pd.DataFrame(unfulfilled_rows)
            satisfaction_rate = (fulfilled_wishes / total_wishes * 100) if total_wishes > 0 else 100.0

            st.session_state['seating_computed'] = True
            st.session_state['df_sorted'] = df_sorted
            st.session_state['unfulfilled_df'] = unfulfilled_df
            st.session_state['fulfilled_wishes'] = fulfilled_wishes
            st.session_state['total_wishes'] = total_wishes
            st.session_state['satisfaction_rate'] = satisfaction_rate
            st.session_state['table_cap'] = table_capacity_val
            st.session_state['pdf_bytes'] = generate_2page_visual_pdf(df_sorted, table_capacity_val)
            st.session_state['excel_bytes'] = build_printable_excel_workbook(df_sorted, unfulfilled_df)
        else:
            st.session_state['seating_computed'] = False
            st.error("Could not find a feasible arrangement. Check group sizes against table capacity.")

# ==========================================
# 8. RENDER RESULTS (PERSISTENT VIA SESSION STATE)
# ==========================================
if st.session_state.get('seating_computed', False):
    df_sorted = st.session_state['df_sorted']
    unfulfilled_df = st.session_state['unfulfilled_df']
    fulfilled_wishes = st.session_state['fulfilled_wishes']
    total_wishes = st.session_state['total_wishes']
    satisfaction_rate = st.session_state['satisfaction_rate']
    table_cap_val = st.session_state['table_cap']
    pdf_bytes = st.session_state['pdf_bytes']
    excel_bytes = st.session_state['excel_bytes']

    # ----------------------------------------------------
    # SCORECARD METRICS
    # ----------------------------------------------------
    st.markdown("<br><h3 class='gala-title'>📊 Seating Satisfaction & Audit</h3>", unsafe_allow_html=True)
    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
    with m_col1:
        st.metric("Total Attendees Seated", len(df_sorted))
    with m_col2:
        st.metric("Total Tables", int(df_sorted['Assigned_Table'].max()))
    with m_col3:
        st.metric("Wishes Fulfilled", f"{fulfilled_wishes} / {total_wishes}")
    with m_col4:
        st.metric("Satisfaction Rate", f"{satisfaction_rate:.1f}%")

    if len(unfulfilled_df) > 0:
        with st.expander(f"⚠️ View Unfulfilled Wishes ({len(unfulfilled_df)} instances)", expanded=False):
            st.dataframe(unfulfilled_df, use_container_width=True, hide_index=True)
    else:
        st.success("🎉 100% of valid seating wishes were perfectly fulfilled!")

    # ----------------------------------------------------
    # KEUKENOVERZICHT
    # ----------------------------------------------------
    st.markdown("<br><h3 class='gala-title'>🍽️ Kitchen & Service Staff Dietary Overview</h3>", unsafe_allow_html=True)
    diet_summary_rows = []
    for tbl_num, grp in df_sorted.groupby('Assigned_Table'):
        with_diet = grp[grp['Dietary'].fillna('').str.strip() != '']
        if len(with_diet) > 0:
            details = " • " + "\n • ".join([f"{r['Name']}: {r['Dietary']}" for _, r in with_diet.iterrows()])
            diet_summary_rows.append({
                'Table': f"Table {tbl_num:02d}",
                'Special Meals Count': len(with_diet),
                'Allergy / Requirement Details': details
            })
    
    if diet_summary_rows:
        st.dataframe(pd.DataFrame(diet_summary_rows), use_container_width=True, hide_index=True)
    else:
        st.success("No dietary restrictions reported.")

    # ----------------------------------------------------
    # VISUELE TAFELKAARTEN
    # ----------------------------------------------------
    st.markdown("<br><h3 class='gala-title'>Seating Chart Overview</h3>", unsafe_allow_html=True)
    tables = df_sorted.groupby('Assigned_Table')
    cols = st.columns(3)
    for idx, (tbl_num, group) in enumerate(tables):
        with cols[idx % 3]:
            st.markdown(f"""
            <div class="table-card">
                <div class="table-header">
                    <span>TABLE {tbl_num:02d}</span>
                    <span style="font-size: 0.85rem; font-weight: normal; color: #6B7280;">{len(group)}/{table_cap_val} Guests</span>
                </div>
            """, unsafe_allow_html=True)
            for seat_idx, (_, r) in enumerate(group.iterrows(), 1):
                comm_label = f"<span class='tag-comm'>{r['Committee']}</span>" if pd.notna(r['Committee']) else "<span class='tag-indiv'>Guest</span>"
                diet_badge = f"<span class='diet-badge'>{r['Dietary']}</span>" if r['Dietary'] else ""
                st.markdown(f"""
                <div class="seat-row">
                    <span><span class="seat-number">{seat_idx:02d}.</span> {r['Name']}{diet_badge}</span>
                    {comm_label}
                </div>
                """, unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    # ----------------------------------------------------
    # DOWNLOAD BUTTONS
    # ----------------------------------------------------
    st.markdown("<hr style='border: none; border-top: 1px solid #E5E7EB; margin: 2rem 0;'>", unsafe_allow_html=True)
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label="📄 Download Visual Seating Plan (2-Page A4 PDF)",
            data=pdf_bytes,
            file_name="Ball_Committee_Seating_Plan_2Page_A4.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="btn_dl_pdf"
        )
    with col_dl2:
        st.download_button(
            label="📥 Download Full Seating Plan, Kitchen Briefing & Audit (Excel)",
            data=excel_bytes,
            file_name="Ball_Committee_Seating_Plan_Printable.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="btn_dl_excel"
        )
