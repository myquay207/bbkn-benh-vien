"""
HỆ THỐNG TỰ ĐỘNG HÓA BIÊN BẢN KIỂM NHẬP
Bệnh viện Đà Nẵng – Khoa Dược
"""

import io
import math
import copy
import datetime
import warnings
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
#  PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="BBKN – Bệnh viện Đà Nẵng",
    page_icon="🏥",
    layout="centered",
)

# ══════════════════════════════════════════════════════════════════════════════
#  CSS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Be Vietnam Pro', sans-serif; }

    /* Header banner */
    .hero {
        background: linear-gradient(135deg, #1a3a5c 0%, #2563a8 60%, #1e7fcb 100%);
        border-radius: 16px;
        padding: 36px 40px 28px;
        margin-bottom: 28px;
        color: white;
        box-shadow: 0 8px 32px rgba(37,99,168,0.25);
    }
    .hero h1 {
        font-size: 1.65rem;
        font-weight: 700;
        letter-spacing: 0.5px;
        margin: 0 0 6px 0;
        line-height: 1.3;
    }
    .hero .sub {
        font-size: 0.92rem;
        font-weight: 300;
        opacity: 0.85;
        margin: 0;
    }
    .hero .badge {
        display: inline-block;
        background: rgba(255,255,255,0.18);
        border-radius: 20px;
        padding: 3px 12px;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 1px;
        margin-bottom: 14px;
        text-transform: uppercase;
    }

    /* Upload box */
    [data-testid="stFileUploader"] {
        border: 2px dashed #2563a8 !important;
        border-radius: 12px !important;
        padding: 8px !important;
        background: #f0f6ff !important;
    }

    /* Button */
    .stButton > button {
        background: linear-gradient(135deg, #1a3a5c, #2563a8) !important;
        color: white !important;
        font-family: 'Be Vietnam Pro', sans-serif !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 14px 0 !important;
        width: 100% !important;
        letter-spacing: 0.5px !important;
        box-shadow: 0 4px 14px rgba(37,99,168,0.35) !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 20px rgba(37,99,168,0.45) !important;
    }

    /* Download button */
    [data-testid="stDownloadButton"] > button {
        background: linear-gradient(135deg, #166534, #16a34a) !important;
        color: white !important;
        font-family: 'Be Vietnam Pro', sans-serif !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 16px 0 !important;
        width: 100% !important;
        letter-spacing: 0.5px !important;
        box-shadow: 0 4px 14px rgba(22,163,74,0.35) !important;
    }

    /* Stat cards */
    .stat-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 14px;
        margin: 20px 0;
    }
    .stat-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 18px 14px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
    .stat-card .num {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1a3a5c;
        line-height: 1;
    }
    .stat-card .lbl {
        font-size: 0.78rem;
        color: #64748b;
        margin-top: 5px;
        font-weight: 400;
    }

    /* Info box */
    .info-box {
        background: #eff6ff;
        border-left: 4px solid #2563a8;
        border-radius: 0 10px 10px 0;
        padding: 14px 18px;
        margin: 16px 0;
        font-size: 0.88rem;
        color: #1e3a5f;
        line-height: 1.6;
    }

    /* Success box */
    .success-box {
        background: #f0fdf4;
        border: 1.5px solid #86efac;
        border-radius: 12px;
        padding: 20px 22px;
        margin: 20px 0;
        text-align: center;
    }
    .success-box .icon { font-size: 2.4rem; }
    .success-box h3 { color: #166534; margin: 8px 0 4px; font-size: 1.1rem; }
    .success-box p  { color: #15803d; font-size: 0.88rem; margin: 0; }

    /* Divider */
    hr { border: none; border-top: 1px solid #e2e8f0; margin: 24px 0; }

    /* Section label */
    .section-label {
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        color: #94a3b8;
        margin-bottom: 10px;
    }

    /* Hide streamlit branding */
    #MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
SKIP_KEYWORDS = ['Tổng cộng', 'Hội đồng', 'Trưởng', 'Trang',
                 'Đã kiểm nhập', 'Ông/bà', 'kiểm nhập những']

COL_WIDTH = {
    1:  5.5,   # A  STT
    2:  10.0,  # B  Số CT
    3:  38.0,  # C  Tên thuốc
    4:  17.0,  # D  Nồng độ hàm lượng
    5:  7.5,   # E  Đơn vị
    6:  11.0,  # F  Số lô
    7:  24.0,  # G  Hãng sản xuất
    8:  11.0,  # H  Hạn dùng
    9:  12.0,  # I  Đơn giá
    10: 9.5,   # J  Số lượng
    11: 14.0,  # K  Thành tiền
    12: 8.0,   # L  Ghi chú
}

COL_ALIGN = {
    1:  ('center', 'center'),
    2:  ('center', 'center'),
    3:  ('left',   'center'),
    4:  ('left',   'center'),
    5:  ('center', 'center'),
    6:  ('center', 'center'),
    7:  ('left',   'center'),
    8:  ('center', 'center'),
    9:  ('right',  'center'),
    10: ('right',  'center'),
    11: ('right',  'center'),
    12: ('center', 'center'),
}

WRAP_COLS   = {3, 4, 7}
NUMBER_COLS = {9, 10, 11}

THIN = Side(style='thin')
MED  = Side(style='medium')

def b_thin():
    return Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def b_med():
    return Border(left=THIN, right=THIN, top=MED, bottom=MED)

NO_FILL = PatternFill(fill_type=None)

# ══════════════════════════════════════════════════════════════════════════════
#  CORE LOGIC
# ══════════════════════════════════════════════════════════════════════════════

def parse_raw_data(raw_df: pd.DataFrame):
    """
    Quét file dữ liệu thô HPT, trả về:
    - companies_data: list of (company_name, [drug_rows])
    - stats: dict thống kê
    """
    companies_data = []
    current_company = None
    drug_rows = []
    skipped = 0

    for _, row in raw_df.iterrows():
        val0 = row[0]

        # Nhận diện dòng tên công ty
        is_company_row = (
            isinstance(val0, str)
            and pd.isna(row[1])
            and pd.isna(row[2])
            and not any(kw in str(val0) for kw in SKIP_KEYWORDS)
        )

        if is_company_row:
            if current_company is not None and drug_rows:
                companies_data.append((current_company, drug_rows))
            current_company = str(val0).strip()
            drug_rows = []
        else:
            # Nhận diện dòng thuốc
            try:
                stt = int(str(val0).strip()) if not pd.isna(val0) else None
            except (ValueError, TypeError):
                stt = None

            col2_val = row[2]
            is_drug = (
                stt is not None
                and not pd.isna(col2_val)
                and isinstance(col2_val, str)
                and not col2_val.strip().isdigit()
            )

            if is_drug:
                qty_val = row[9]
                try:
                    qty = float(qty_val) if not pd.isna(qty_val) else 0
                except (ValueError, TypeError):
                    qty = 0
                if qty > 0:
                    drug_rows.append(row)
                else:
                    skipped += 1

    if current_company is not None and drug_rows:
        companies_data.append((current_company, drug_rows))

    total_drugs = sum(len(d) for _, d in companies_data)
    stats = {
        "companies": len(companies_data),
        "drugs": total_drugs,
        "skipped": skipped,
    }
    return companies_data, stats


def estimate_row_height(row_num, ws, font_size=12):
    line_height = font_size * 1.3
    max_lines   = 1
    for col in WRAP_COLS:
        val = ws.cell(row=row_num, column=col).value
        if not val or not isinstance(val, str):
            continue
        chars_per_line = COL_WIDTH.get(col, 15) * 1.1
        chars_per_line = max(chars_per_line, 1)
        wrapped = sum(
            max(1, math.ceil(len(ln) / chars_per_line))
            for ln in val.split('\n')
        )
        max_lines = max(max_lines, wrapped)
    return max(22, min(max_lines * line_height + 4, 120))


def build_excel(template_bytes: bytes, companies_data: list) -> bytes:
    """
    Điền dữ liệu vào template và áp dụng toàn bộ định dạng.
    Trả về bytes của file .xlsx hoàn chỉnh.
    """
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb.active

    # ── Lấy style mẫu từ template ────────────────────────────────────────────
    def get_style(row, col):
        c = ws.cell(row=row, column=col)
        return {
            'font':          copy.copy(c.font),
            'border':        copy.copy(c.border),
            'alignment':     copy.copy(c.alignment),
            'fill':          copy.copy(c.fill),
            'number_format': c.number_format,
        }

    company_styles = {col: get_style(15, col) for col in range(1, 13)}
    drug_styles    = {col: get_style(16, col) for col in range(1, 13)}
    total_k_style  = get_style(213, 11)

    def apply_style(cell, style):
        cell.font          = copy.copy(style['font'])
        cell.border        = copy.copy(style['border'])
        cell.alignment     = copy.copy(style['alignment'])
        cell.fill          = copy.copy(style['fill'])
        if style['number_format']:
            cell.number_format = style['number_format']

    # ── Tìm footer ────────────────────────────────────────────────────────────
    footer_start = None
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == 'HỘI ĐỒNG KIỂM NHẬP':
                footer_start = cell.row
                break
        if footer_start:
            break

    DATA_START = 15
    total_data_rows = sum(1 + len(drugs) for _, drugs in companies_data) + 1
    data_end_row    = DATA_START + total_data_rows - 1

    # Chèn dòng nếu cần
    rows_needed = data_end_row - footer_start + 1
    if rows_needed > 0:
        ws.insert_rows(footer_start, rows_needed)
        footer_start += rows_needed

    # Xóa merge trong vùng dữ liệu
    for m in [str(mr) for mr in ws.merged_cells.ranges
              if DATA_START <= mr.min_row < footer_start]:
        ws.merged_cells.remove(m)

    # Xóa giá trị cũ
    for r in range(DATA_START, footer_start):
        for c in range(1, 13):
            try:
                ws.cell(row=r, column=c).value = None
            except AttributeError:
                pass

    # ── Ghi dữ liệu ──────────────────────────────────────────────────────────
    def write_company_row(row_num, name):
        cell = ws.cell(row=row_num, column=1, value=name)
        apply_style(cell, company_styles[1])
        cell.font = Font(name='Times New Roman', bold=True, size=12)
        cell.fill = NO_FILL
        for c in range(2, 13):
            c2 = ws.cell(row=row_num, column=c)
            apply_style(c2, company_styles[c])
            c2.fill = NO_FILL
        ws.row_dimensions[row_num].height = 20

    def write_drug_row(row_num, stt, dr):
        # A: STT
        ca = ws.cell(row=row_num, column=1, value=stt)
        apply_style(ca, drug_styles[1])
        ca.font      = Font(name='Times New Roman', size=12)
        ca.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # B: Số CT
        so_ct = '' if pd.isna(dr[1]) else str(dr[1]).strip()
        cb = ws.cell(row=row_num, column=2, value=so_ct)
        apply_style(cb, drug_styles[2])
        cb.font = Font(name='Times New Roman', size=12)
        cb.alignment = Alignment(horizontal='center', vertical='center')

        # C: Tên thuốc
        ten = '' if pd.isna(dr[2]) else str(dr[2]).strip()
        cc = ws.cell(row=row_num, column=3, value=ten)
        apply_style(cc, drug_styles[3])
        cc.font      = Font(name='Times New Roman', size=12)
        cc.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        # D: Nồng độ
        nd = '' if pd.isna(dr[3]) else str(dr[3]).strip()
        cd = ws.cell(row=row_num, column=4, value=nd)
        apply_style(cd, drug_styles[4])
        cd.font      = Font(name='Times New Roman', size=12)
        cd.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        # E: Đơn vị
        dv = '' if pd.isna(dr[4]) else str(dr[4]).strip()
        ce = ws.cell(row=row_num, column=5, value=dv)
        apply_style(ce, drug_styles[5])
        ce.font      = Font(name='Times New Roman', size=12)
        ce.alignment = Alignment(horizontal='center', vertical='center')

        # F: Số lô
        sl = '' if pd.isna(dr[5]) else str(dr[5]).strip()
        cf = ws.cell(row=row_num, column=6, value=sl)
        apply_style(cf, drug_styles[6])
        cf.font      = Font(name='Times New Roman', size=12)
        cf.alignment = Alignment(horizontal='center', vertical='center')

        # G: Hãng sản xuất
        hang = '' if pd.isna(dr[6]) else str(dr[6]).strip()
        cg = ws.cell(row=row_num, column=7, value=hang)
        apply_style(cg, drug_styles[7])
        cg.font      = Font(name='Times New Roman', size=12)
        cg.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        # H: Hạn dùng
        hd = dr[7]
        if not isinstance(hd, datetime.datetime) and pd.isna(hd):
            hd = ''
        ch = ws.cell(row=row_num, column=8, value=hd)
        apply_style(ch, drug_styles[8])
        ch.font      = Font(name='Times New Roman', size=12)
        ch.alignment = Alignment(horizontal='center', vertical='center')
        if isinstance(hd, datetime.datetime):
            ch.number_format = 'DD/MM/YYYY'

        # I: Đơn giá
        dg = dr[8] if not pd.isna(dr[8]) else 0
        ci = ws.cell(row=row_num, column=9, value=dg)
        apply_style(ci, drug_styles[9])
        ci.font          = Font(name='Times New Roman', size=12)
        ci.alignment     = Alignment(horizontal='right', vertical='center')
        ci.number_format = '#,##0'

        # J: Số lượng
        qty = dr[9] if not pd.isna(dr[9]) else 0
        cj = ws.cell(row=row_num, column=10, value=int(qty))
        apply_style(cj, drug_styles[10])
        cj.font          = Font(name='Times New Roman', size=12)
        cj.alignment     = Alignment(horizontal='right', vertical='center')
        cj.number_format = '#,##0'

        # K: Thành tiền (formula)
        ck = ws.cell(row=row_num, column=11, value=f'=I{row_num}*J{row_num}')
        apply_style(ck, drug_styles[11])
        ck.font          = Font(name='Times New Roman', size=12)
        ck.alignment     = Alignment(horizontal='right', vertical='center')
        ck.number_format = '#,##0'

        # L: Ghi chú
        cl = ws.cell(row=row_num, column=12, value='')
        apply_style(cl, drug_styles[12])

    # Ghi toàn bộ dữ liệu
    current_row   = DATA_START
    drug_row_nums = []

    for company_name, drugs in companies_data:
        write_company_row(current_row, company_name)
        current_row += 1
        for stt_i, dr in enumerate(drugs, 1):
            write_drug_row(current_row, stt_i, dr)
            drug_row_nums.append(current_row)
            current_row += 1

    # ── Dòng tổng cộng ───────────────────────────────────────────────────────
    total_row = current_row
    lbl = ws.cell(row=total_row, column=1, value='Tổng cộng: ')
    lbl.font      = Font(name='Times New Roman', bold=True, size=12)
    lbl.alignment = Alignment(horizontal='left', vertical='center')
    lbl.border    = b_med()

    for c in range(2, 11):
        try:
            ws.cell(row=total_row, column=c).border = b_med()
        except AttributeError:
            pass

    sum_formula = f'=SUM({",".join(f"K{r}" for r in drug_row_nums)})'
    ck_total = ws.cell(row=total_row, column=11, value=sum_formula)
    apply_style(ck_total, total_k_style)
    ck_total.font          = Font(name='Times New Roman', bold=True, size=12)
    ck_total.alignment     = Alignment(horizontal='right', vertical='center')
    ck_total.number_format = '#,##0'
    ck_total.border        = b_med()
    ws.row_dimensions[total_row].height = 22

    # ── Định dạng thẩm mỹ ────────────────────────────────────────────────────
    HEADER_ROW = 13
    SUBHDR_ROW = 14

    # Sửa header cột C
    ws.cell(row=HEADER_ROW, column=3).value = 'Tên thuốc'

    # Bỏ fill header, giữ border + font
    for col in range(1, 13):
        for r in (HEADER_ROW, SUBHDR_ROW):
            cell = ws.cell(row=r, column=col)
            try:
                cell.fill      = NO_FILL
                cell.font      = Font(name='Times New Roman', bold=True, size=12)
                cell.border    = b_med()
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            except AttributeError:
                pass
    ws.row_dimensions[HEADER_ROW].height = 42
    ws.row_dimensions[SUBHDR_ROW].height = 18

    # Style & auto height vùng dữ liệu
    for r in range(DATA_START, total_row + 1):
        a_val = ws.cell(row=r, column=1).value
        c_val = ws.cell(row=r, column=3).value
        is_company = isinstance(a_val, str) and not str(a_val).strip().lstrip('-').isdigit() and not c_val
        is_total_r = (r == total_row)

        if is_total_r:
            pass  # already styled above
        elif is_company:
            ws.row_dimensions[r].height = 20
            for col in range(1, 13):
                cell = ws.cell(row=r, column=col)
                try:
                    cell.fill      = NO_FILL
                    cell.border    = b_thin()
                    cell.font      = Font(name='Times New Roman', bold=True, size=12)
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                except AttributeError:
                    pass
        else:
            auto_h = estimate_row_height(r, ws)
            ws.row_dimensions[r].height = auto_h
            for col in range(1, 13):
                cell = ws.cell(row=r, column=col)
                try:
                    wrap       = col in WRAP_COLS
                    h_a, v_a   = COL_ALIGN.get(col, ('left', 'center'))
                    cell.fill   = NO_FILL
                    cell.border = b_thin()
                    cell.font   = Font(name='Times New Roman', size=12)
                    cell.alignment = Alignment(horizontal=h_a, vertical=v_a, wrap_text=wrap)
                    if col in NUMBER_COLS and cell.value is not None:
                        cell.number_format = '#,##0'
                    if col == 8 and isinstance(cell.value, datetime.datetime):
                        cell.number_format = 'DD/MM/YYYY'
                except AttributeError:
                    pass

    # Độ rộng cột
    for col, width in COL_WIDTH.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    # Page setup
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left   = 0.4
    ws.page_margins.right  = 0.4
    ws.page_margins.top    = 0.5
    ws.page_margins.bottom = 0.5
    ws.page_margins.header = 0.2
    ws.page_margins.footer = 0.2
    ws.print_title_rows    = f'1:{SUBHDR_ROW}'
    ws.freeze_panes        = ws.cell(row=DATA_START, column=1)

    # ── Xuất ra bytes ─────────────────────────────────────────────────────────
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════════════════════════════════════

# Hero banner
st.markdown("""
<div class="hero">
    <div class="badge">🏥 Bệnh viện Đà Nẵng · Khoa Dược</div>
    <h1>HỆ THỐNG TỰ ĐỘNG HÓA<br>BIÊN BẢN KIỂM NHẬP</h1>
    <p class="sub">Xử lý dữ liệu từ phần mềm HPT · Xuất Biên Bản Kiểm Nhập chuẩn in hội đồng</p>
</div>
""", unsafe_allow_html=True)

# Hướng dẫn
st.markdown("""
<div class="info-box">
    <b>Cách sử dụng:</b><br>
    1️⃣ &nbsp;Upload <b>file dữ liệu thô từ HPT</b> (.xls / .xlsx) vào ô bên dưới<br>
    2️⃣ &nbsp;Upload <b>file form chuẩn kiểm nhập 2026</b> (.xlsx) làm mẫu<br>
    3️⃣ &nbsp;Nhấn <b>Bắt đầu xử lý</b> → Tải file hoàn chỉnh về in ký hội đồng
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-label">📂 Tải file lên</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    raw_file = st.file_uploader(
        "File dữ liệu thô HPT",
        type=["xls", "xlsx"],
        help="File xuất từ phần mềm HPT, định dạng .xls hoặc .xlsx",
    )
with col2:
    template_file = st.file_uploader(
        "File form chuẩn 2026",
        type=["xlsx"],
        help="File form_chuan_kiểm_nhập_2026.xlsx",
    )

st.markdown("<hr>", unsafe_allow_html=True)

# ── Nút xử lý ────────────────────────────────────────────────────────────────
if st.button("⚡  Bắt đầu xử lý", disabled=(not raw_file or not template_file)):

    with st.spinner("Đang đọc và xử lý dữ liệu..."):
        try:
            # Đọc file thô – hỗ trợ cả .xls và .xlsx
            raw_bytes = raw_file.read()
            if raw_file.name.endswith('.xls'):
                # Chuyển xls → xlsx qua openpyxl-compatible path bằng xlrd/pandas
                try:
                    raw_df = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=0,
                                           header=None, engine='xlrd')
                except Exception:
                    raw_df = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=0,
                                           header=None)
            else:
                raw_df = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=0, header=None)

            # Parse
            companies_data, stats = parse_raw_data(raw_df)

            if not companies_data:
                st.error("❌ Không tìm thấy dữ liệu hợp lệ trong file. Vui lòng kiểm tra lại file HPT.")
                st.stop()

            # Build Excel
            template_bytes = template_file.read()
            result_bytes   = build_excel(template_bytes, companies_data)

            st.session_state["result_bytes"] = result_bytes
            st.session_state["stats"]        = stats
            st.session_state["processed"]    = True

        except Exception as e:
            st.error(f"❌ Lỗi khi xử lý: {str(e)}")
            st.exception(e)
            st.stop()

# ── Hiển thị kết quả ─────────────────────────────────────────────────────────
if st.session_state.get("processed"):
    stats  = st.session_state["stats"]
    result = st.session_state["result_bytes"]

    st.markdown(f"""
    <div class="success-box">
        <div class="icon">✅</div>
        <h3>Xử lý hoàn tất!</h3>
        <p>File biên bản kiểm nhập đã sẵn sàng để tải về và in.</p>
    </div>
    <div class="stat-grid">
        <div class="stat-card">
            <div class="num">{stats['companies']}</div>
            <div class="lbl">Công ty cung cấp</div>
        </div>
        <div class="stat-card">
            <div class="num">{stats['drugs']}</div>
            <div class="lbl">Mặt hàng có nhập</div>
        </div>
        <div class="stat-card">
            <div class="num">{stats['skipped']}</div>
            <div class="lbl">Dòng đã lọc bỏ (SL = 0)</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Tên file theo tháng của dữ liệu
    now = datetime.datetime.now()
    filename = f"BBKN_T{now.month}_{now.year}_HoanChinh.xlsx"

    st.download_button(
        label="⬇️  Tải Biên Bản Hoàn Chỉnh (.xlsx)",
        data=result,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown("""
    <div class="info-box" style="margin-top:16px;">
        <b>Lưu ý khi in:</b> Mở file Excel → <b>File → Print</b> → 
        Kiểm tra khổ giấy <b>A4 Ngang (Landscape)</b> và chế độ 
        <b>"Fit Sheet on One Page"</b> đã được thiết lập sẵn.
    </div>
    """, unsafe_allow_html=True)
