"""
===============================================================
  TOOL: Import Log Note từ Excel vào Odoo (th.project / th.project.task / th.issue)
===============================================================

FILE EXCEL CÓ 2 SHEET CHÍNH:
  - 'Log Notes - Dự án'      : Log của th.project  (match bằng mã dự án)
  - 'Log Notes - Task dự án' : Log của th.project.task (match bằng mã dự án + tên task)
  - 'Tổng hợp - Tất cả'      : (dự phòng) chứa toàn bộ, bao gồm itc.wms.bug -> th.issue

CÁCH DÙNG:
  python import_lognote_du_an.py                    # Chạy tất cả, môi trường default
  python import_lognote_du_an.py --env production   # Môi trường production
  python import_lognote_du_an.py --scope project    # Chỉ chạy log dự án
  python import_lognote_du_an.py --scope task       # Chỉ chạy log task
  python import_lognote_du_an.py --scope issue      # Chỉ chạy log lỗi/sự cố
  python import_lognote_du_an.py --dry-run          # Xem trước, không ghi DB
  python import_lognote_du_an.py --list-envs        # Liệt kê môi trường
  python import_lognote_du_an.py --reset            # Xoá log cũ đã import (cẩn thận!)
"""

import argparse
import sys
import os
import psycopg2
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
from collections import defaultdict

# =============================================================
#  ⚙️  CẤU HÌNH MÔI TRƯỜNG
# =============================================================
ENVIRONMENTS = {
    "default": {
        "description": "Local dev (omega)",
        "db_host":     "127.0.0.1",
        "db_port":     5432,
        "db_name":     "omega",
        "db_user":     "odoo",
        "db_password": "1",
        "excel_path":  r"d:\omega\lognote_du_an.xlsx",
        "timezone_offset_hours": 7,
        "admin_login": "admin",
    },
    # Thêm môi trường mới tại đây:
    # "production": {
    #     "description": "Production server",
    #     "db_host":     "10.0.0.100",
    #     "db_port":     5432,
    #     "db_name":     "omega_prod",
    #     "db_user":     "odoo",
    #     "db_password": "your_password",
    #     "excel_path":  r"C:\path\to\lognote_du_an.xlsx",
    #     "timezone_offset_hours": 7,
    #     "admin_login": "admin",
    # },
}

# Nhãn phân biệt log note đã import (dùng để kiểm tra trùng + reset)
IMPORT_MARKER = "data-imported-from-excel"

# =============================================================
#   HELPER: Tạo HTML log note
# =============================================================

def _esc(s):
    """Escape HTML basic chars."""
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def build_html_project(ma_du_an, ten_du_an, entries):
    """HTML cho log note dự án (sheet 'Log Notes - Dự án')."""
    rows_html = []
    for e in entries:
        t = e["ngay_gio"]
        t_str = t.strftime("%d/%m/%Y %H:%M:%S") if isinstance(t, datetime) else _esc(t or "")
        nguoi = _esc(e["nguoi_ghi"] or "")
        if e["noi_dung"]:
            rows_html.append(
                f'<tr><td style="padding:4px 8px;border:1px solid #ddd;white-space:nowrap">{t_str}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd">{nguoi}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd;color:#0055cc;font-weight:bold" colspan="3">'
                f'{_esc(e["noi_dung"])}</td></tr>'
            )
        elif e["ten_truong"]:
            cu  = _esc(e["gia_tri_cu"] or "").strip() or '<em style="color:#999">(trống)</em>'
            moi = _esc(e["gia_tri_moi"] or "").strip() or '<em style="color:#999">(trống)</em>'
            rows_html.append(
                f'<tr><td style="padding:4px 8px;border:1px solid #ddd;white-space:nowrap">{t_str}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd">{nguoi}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd">{_esc(e["ten_truong"])}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd;color:#cc0000">{cu}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd;color:#006600">{moi}</td></tr>'
            )
    if not rows_html:
        return ""
    return f"""<div {IMPORT_MARKER}="1" style="font-family:Arial,sans-serif;font-size:13px">
<p style="margin:0 0 8px;font-weight:bold;color:#333">
📋 Lịch sử thay đổi dự án: <strong>{_esc(ten_du_an)}</strong> [{_esc(ma_du_an)}]
</p>
<table style="border-collapse:collapse;width:100%;font-size:12px">
<thead><tr style="background:#f0f0f0">
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Ngày giờ (GMT+7)</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Người ghi</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Trường / Nội dung</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Giá trị cũ</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Giá trị mới</th>
</tr></thead>
<tbody>{''.join(rows_html)}</tbody>
</table>
<p style="margin:8px 0 0;font-size:11px;color:#888">
(Nhập từ log Excel — {len(entries)} bản ghi)
</p></div>"""


def build_html_task(ma_du_an, ten_task, entries):
    """HTML cho log note task (sheet 'Log Notes - Task dự án')."""
    rows_html = []
    for e in entries:
        t = e["ngay_gio"]
        t_str = t.strftime("%d/%m/%Y %H:%M:%S") if isinstance(t, datetime) else _esc(t or "")
        nguoi  = _esc(e["nguoi_ghi"] or "")
        tieu_de = _esc(e.get("tieu_de") or "")

        if e["noi_dung"]:
            rows_html.append(
                f'<tr><td style="padding:4px 8px;border:1px solid #ddd;white-space:nowrap">{t_str}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd">{nguoi}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd;color:#0055cc;font-weight:bold" colspan="3">'
                f'{_esc(e["noi_dung"])}</td></tr>'
            )
        elif e["ten_truong"]:
            cu  = _esc(e["gia_tri_cu"] or "").strip() or '<em style="color:#999">(trống)</em>'
            moi = _esc(e["gia_tri_moi"] or "").strip() or '<em style="color:#999">(trống)</em>'
            rows_html.append(
                f'<tr><td style="padding:4px 8px;border:1px solid #ddd;white-space:nowrap">{t_str}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd">{nguoi}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd">{_esc(e["ten_truong"])}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd;color:#cc0000">{cu}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd;color:#006600">{moi}</td></tr>'
            )
    if not rows_html:
        return ""
    return f"""<div {IMPORT_MARKER}="1" style="font-family:Arial,sans-serif;font-size:13px">
<p style="margin:0 0 8px;font-weight:bold;color:#333">
📋 Lịch sử task: <strong>{_esc(ten_task)}</strong>
<span style="color:#666;font-size:11px"> (dự án [{_esc(ma_du_an)}])</span>
</p>
<table style="border-collapse:collapse;width:100%;font-size:12px">
<thead><tr style="background:#e8f4e8">
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Ngày giờ (GMT+7)</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Người ghi</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Trường / Nội dung</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Giá trị cũ</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Giá trị mới</th>
</tr></thead>
<tbody>{''.join(rows_html)}</tbody>
</table>
<p style="margin:8px 0 0;font-size:11px;color:#888">
(Nhập từ log Excel — {len(entries)} bản ghi)
</p></div>"""


def build_html_issue(ten_issue, entries):
    """HTML cho log note lỗi/sự cố (từ sheet 'Tổng hợp - Tất cả', model itc.wms.bug)."""
    rows_html = []
    for e in entries:
        t = e["ngay_gio"]
        t_str = t.strftime("%d/%m/%Y %H:%M:%S") if isinstance(t, datetime) else _esc(t or "")
        nguoi = _esc(e["nguoi_ghi"] or "")
        if e["noi_dung"]:
            rows_html.append(
                f'<tr><td style="padding:4px 8px;border:1px solid #ddd;white-space:nowrap">{t_str}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd">{nguoi}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd;color:#cc3300;font-weight:bold" colspan="3">'
                f'{_esc(e["noi_dung"])}</td></tr>'
            )
        elif e["ten_truong"]:
            cu  = _esc(e["gia_tri_cu"] or "").strip() or '<em style="color:#999">(trống)</em>'
            moi = _esc(e["gia_tri_moi"] or "").strip() or '<em style="color:#999">(trống)</em>'
            rows_html.append(
                f'<tr><td style="padding:4px 8px;border:1px solid #ddd;white-space:nowrap">{t_str}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd">{nguoi}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd">{_esc(e["ten_truong"])}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd;color:#cc0000">{cu}</td>'
                f'<td style="padding:4px 8px;border:1px solid #ddd;color:#006600">{moi}</td></tr>'
            )
    if not rows_html:
        return ""
    return f"""<div {IMPORT_MARKER}="1" style="font-family:Arial,sans-serif;font-size:13px">
<p style="margin:0 0 8px;font-weight:bold;color:#cc3300">
🐛 Lịch sử lỗi/sự cố: <strong>{_esc(ten_issue)}</strong>
</p>
<table style="border-collapse:collapse;width:100%;font-size:12px">
<thead><tr style="background:#fdf0f0">
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Ngày giờ (GMT+7)</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Người ghi</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Trường / Nội dung</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Giá trị cũ</th>
<th style="padding:6px 8px;border:1px solid #ddd;text-align:left">Giá trị mới</th>
</tr></thead>
<tbody>{''.join(rows_html)}</tbody>
</table>
<p style="margin:8px 0 0;font-size:11px;color:#888">
(Nhập từ log Excel — {len(entries)} bản ghi)
</p></div>"""


# =============================================================
#   ĐỌC EXCEL
# =============================================================

def _adj_time(dt, offset_h):
    if isinstance(dt, datetime):
        return dt + timedelta(hours=offset_h)
    return dt


def read_sheet_project(wb, tz_offset):
    """Sheet 'Log Notes - Dự án': group by mã dự án."""
    ws = wb['Log Notes - Dự án']
    data = defaultdict(list)
    for row in range(2, ws.max_row + 1):
        ma_du_an = ws.cell(row=row, column=2).value
        if not ma_du_an:
            continue
        ma_du_an = str(ma_du_an).strip()
        data[ma_du_an].append({
            "ngay_gio":    _adj_time(ws.cell(row=row, column=1).value, tz_offset),
            "ten_du_an":   ws.cell(row=row, column=3).value,
            "nguoi_ghi":   ws.cell(row=row, column=4).value,
            "noi_dung":    ws.cell(row=row, column=5).value,
            "ten_truong":  ws.cell(row=row, column=6).value,
            "gia_tri_cu":  ws.cell(row=row, column=7).value,
            "gia_tri_moi": ws.cell(row=row, column=8).value,
        })
    print(f"  📄 Sheet dự án: {sum(len(v) for v in data.values())} dòng, {len(data)} mã dự án")
    return data


def read_sheet_task(wb, tz_offset):
    """Sheet 'Log Notes - Task dự án': group by (mã dự án, tên task)."""
    ws = wb['Log Notes - Task dự án']
    data = defaultdict(list)
    for row in range(2, ws.max_row + 1):
        ma_du_an = ws.cell(row=row, column=3).value
        ten_task = ws.cell(row=row, column=5).value
        if not ma_du_an or not ten_task:
            continue
        key = (str(ma_du_an).strip(), str(ten_task).strip())
        data[key].append({
            "ngay_gio":    _adj_time(ws.cell(row=row, column=2).value, tz_offset),
            "ten_du_an":   ws.cell(row=row, column=4).value,
            "trang_thai":  ws.cell(row=row, column=6).value,
            "nguoi_ghi":   ws.cell(row=row, column=7).value,
            "loai_msg":    ws.cell(row=row, column=8).value,
            "subtype":     ws.cell(row=row, column=9).value,
            "noi_bo":      ws.cell(row=row, column=10).value,
            "tieu_de":     ws.cell(row=row, column=11).value,
            "noi_dung":    ws.cell(row=row, column=12).value,
            "ten_truong":  ws.cell(row=row, column=13).value,
            "field_name":  ws.cell(row=row, column=14).value,
            "kieu_dl":     ws.cell(row=row, column=15).value,
            "gia_tri_cu":  ws.cell(row=row, column=16).value,
            "gia_tri_moi": ws.cell(row=row, column=17).value,
        })
    print(f"  📄 Sheet task: {sum(len(v) for v in data.values())} dòng, {len(data)} cặp (dự án, task)")
    return data


def read_sheet_issue(wb, tz_offset):
    """Sheet 'Tổng hợp - Tất cả': lọc model=itc.wms.bug, group by tên record."""
    ws = wb['Tổng hợp - Tất cả']
    data = defaultdict(list)
    for row in range(2, ws.max_row + 1):
        model = ws.cell(row=row, column=3).value
        if model != 'itc.wms.bug':
            continue
        rec_name = ws.cell(row=row, column=5).value
        if not rec_name:
            # thử lấy từ nội dung
            noi_dung = ws.cell(row=row, column=11).value
            rec_name = f"(Bug #{ws.cell(row=row, column=4).value})"
        key = str(rec_name).strip()
        data[key].append({
            "ngay_gio":    _adj_time(ws.cell(row=row, column=2).value, tz_offset),
            "nguoi_ghi":   ws.cell(row=row, column=6).value,
            "loai_msg":    ws.cell(row=row, column=7).value,
            "noi_bo":      ws.cell(row=row, column=9).value,
            "noi_dung":    ws.cell(row=row, column=11).value,
            "ten_truong":  ws.cell(row=row, column=12).value,
            "field_name":  ws.cell(row=row, column=13).value,
            "gia_tri_cu":  ws.cell(row=row, column=15).value,
            "gia_tri_moi": ws.cell(row=row, column=16).value,
        })
    print(f"  📄 Sheet issue (bug): {sum(len(v) for v in data.values())} dòng, {len(data)} lỗi/sự cố")
    return data


# =============================================================
#   BUILD LOOKUP TABLES TỪ DB
# =============================================================

def build_db_lookups(cur):
    """Trả về các dict tra cứu ID theo tên kèm thông tin chi tiết."""

    # --- Projects: th_code -> (project_id, project_name) ---
    cur.execute("SELECT id, th_code, name FROM th_project WHERE th_code IS NOT NULL")
    project_by_code = {}
    for pid, code, name in cur.fetchall():
        code_str = str(code).strip()
        val = (pid, name or "")
        project_by_code[code_str] = val
        project_by_code[code_str.lower()] = val  # case-insensitive fallback

    # --- Tasks: build 3 fallback strategies -> (task_id, task_name, task_code) ---
    cur.execute("""
        SELECT t.id, t.name, t.th_code, p.th_code as pcode
        FROM th_project_task t
        JOIN th_project p ON p.id = t.th_project_id
        WHERE t.name IS NOT NULL
    """)
    task_exact = {}         # (pcode, full_name) -> (tid, tname, tcode)
    task_stripped = {}      # (pcode, stripped_name) -> (tid, tname, tcode)
    task_name_in_proj = {}  # (pcode, name_lower) -> (tid, tname, tcode)
    for tid, tname, tcode, pcode in cur.fetchall():
        pcode = str(pcode or "").strip()
        fname = str(tname).strip()
        tcode_str = str(tcode or "").strip()
        sname = fname[fname.index(']') + 1:].strip() if (fname.startswith('[') and ']' in fname) else fname
        val = (tid, fname, tcode_str)
        task_exact[(pcode, fname)] = val
        task_stripped[(pcode, sname)] = val
        task_name_in_proj[(pcode, sname.lower())] = val

    # --- Issues: match by name -> (issue_id, issue_name) ---
    cur.execute("SELECT id, name FROM th_issue WHERE name IS NOT NULL")
    issue_by_name = {}
    issue_by_name_lower = {}
    for iid, iname in cur.fetchall():
        iname_str = str(iname).strip()
        sname = iname_str[iname_str.index(']') + 1:].strip() if (iname_str.startswith('[') and ']' in iname_str) else iname_str
        val = (iid, iname_str)
        issue_by_name[iname_str] = val
        issue_by_name_lower[sname.lower()] = val
        issue_by_name_lower[iname_str.lower()] = val

    return project_by_code, task_exact, task_stripped, task_name_in_proj, issue_by_name, issue_by_name_lower


def find_task(ma_du_an, ten_task, task_exact, task_stripped, task_name_in_proj):
    """Tìm task ID theo 3 chiến lược fallback. Trả về ((tid, tname, tcode), strat)."""
    # 1. Exact: (project_code, task_name)
    val = task_exact.get((ma_du_an, ten_task))
    if val:
        return val, "exact"
    # 2. Stripped: (project_code, task_name after stripping [CODE])
    val = task_stripped.get((ma_du_an, ten_task))
    if val:
        return val, "stripped"
    # 3. Case-insensitive within same project
    val = task_name_in_proj.get((ma_du_an, ten_task.lower()))
    if val:
        return val, "case_insensitive"
    return None, None


def find_issue(ten_issue, issue_by_name, issue_by_name_lower):
    """Tìm issue ID theo tên. Trả về ((iid, iname), strat)."""
    val = issue_by_name.get(ten_issue)
    if val:
        return val, "exact"
    sname = ten_issue[ten_issue.index(']') + 1:].strip() if (ten_issue.startswith('[') and ']' in ten_issue) else ten_issue
    val = issue_by_name_lower.get(sname.lower())
    if val:
        return val, "stripped"
    val = issue_by_name_lower.get(ten_issue.lower())
    if val:
        return val, "lower"
    return None, None


# =============================================================
#   INSERT LOG NOTE
# =============================================================

def insert_lognote(cur, model, res_id, html_body, subtype_id, author_partner_id, author_user_id):
    now_utc = datetime.utcnow()
    cur.execute("""
        INSERT INTO mail_message (
            model, res_id, message_type, subtype_id,
            author_id, body, date, is_internal,
            create_uid, write_uid, create_date, write_date
        ) VALUES ('""" + model + """', %s, 'comment', %s, %s, %s, %s, true,
                  %s, %s, %s, %s)
        RETURNING id
    """, (
        res_id, subtype_id, author_partner_id, html_body,
        now_utc, author_user_id, author_user_id, now_utc, now_utc,
    ))
    return cur.fetchone()[0]


def already_imported(cur, model, res_id):
    cur.execute(
        "SELECT id FROM mail_message WHERE model=%s AND res_id=%s AND (body LIKE %s OR body LIKE %s OR body LIKE %s) LIMIT 1",
        (model, res_id, f'%{IMPORT_MARKER}%', '%Nhập từ file log Excel%', '%Dữ liệu được nhập từ file log Excel%')
    )
    return cur.fetchone()


# =============================================================
#   XUẤT BÁO CÁO MAPPING EXCEL (CHUẨN FORM)
# =============================================================

# Styles chuẩn
COLOR_HEADER   = PatternFill("solid", fgColor="2F5496")
COLOR_TOTAL    = PatternFill("solid", fgColor="1F3864")
COLOR_GREEN    = PatternFill("solid", fgColor="C6EFCE")
COLOR_RED      = PatternFill("solid", fgColor="FFC7CE")
COLOR_YELLOW   = PatternFill("solid", fgColor="FFEB9C")

FONT_HEADER    = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
FONT_BOLD      = Font(name="Calibri", bold=True, size=10)
FONT_NORMAL    = Font(name="Calibri", size=10)

ALIGN_CENTER   = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT     = Alignment(horizontal="left",   vertical="center", wrap_text=True)

def _thin_border():
    s = Side(style="thin", color="D9D9D9")
    return Border(left=s, right=s, top=s, bottom=s)


def _set_cell(ws, row, col, value, fill=None, font=None, align=ALIGN_LEFT):
    c = ws.cell(row=row, column=col, value=value)
    if fill:
        c.fill = fill
    c.font = font or FONT_NORMAL
    c.alignment = align
    c.border = _thin_border()
    return c


def _set_header(ws, row, col, value, width=None):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = COLOR_HEADER
    c.font = FONT_HEADER
    c.alignment = ALIGN_CENTER
    c.border = _thin_border()
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width


def export_report(report_data: dict, output_path: str):
    """
    Xuất file Excel báo cáo mapping y hệt format d:\\omega\\bao_cao_mapping_lognote.xlsx
    report_data = {
        'projects': list of dicts {excel_code, excel_name, db_id, db_name, status},
        'tasks':    list of dicts {excel_code, excel_proj_name, excel_task, db_id, db_code, db_name},
        'issues':   list of dicts {excel_name, db_id, db_name, status},
    }
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Xóa sheet mặc định

    # ─────────────────────────────────────────────────────────────
    # SHEET 1: ' Tổng quan'
    # ─────────────────────────────────────────────────────────────
    ws1 = wb.create_sheet(' Tổng quan')
    headers1 = ['Module', 'Tổng record Excel', 'Map được (DB)', 'Không map', 'Tỷ lệ map', 'Đã import', 'Chưa import', 'Ghi chú']
    widths1  = [30.0, 18.0, 15.0, 12.0, 13.0, 13.0, 13.0, 30.0]
    for ci, (h, w) in enumerate(zip(headers1, widths1), 1):
        _set_header(ws1, 1, ci, h, w)
    ws1.row_dimensions[1].height = 26

    # Thống kê từng module
    proj_list = report_data.get('projects', [])
    task_list = report_data.get('tasks', [])
    iss_list  = report_data.get('issues', [])

    def calc_stats(lst):
        total = len(lst)
        matched = sum(1 for x in lst if x.get('db_id'))
        no_map = total - matched
        rate_str = f"{matched / total * 100:.1f}%" if total else "0.0%"
        imported = sum(1 for x in lst if 'Đã import' in str(x.get('status', '')))
        not_yet = matched - imported
        return total, matched, no_map, rate_str, imported, not_yet

    p_tot, p_mat, p_nom, p_rat, p_imp, p_not = calc_stats(proj_list)
    t_tot, t_mat, t_nom, t_rat, t_imp, t_not = calc_stats(task_list)
    i_tot, i_mat, i_nom, i_rat, i_imp, i_not = calc_stats(iss_list)

    note_text = "Record trong DB khác hệ thống cũ"
    summary_data = [
        ('Dự án',    p_tot, p_mat, p_nom, p_rat, p_imp, p_not, note_text if p_nom > p_mat else "Phần lớn đã map"),
        ('Đầu việc', t_tot, t_mat, t_nom, t_rat, t_imp, t_not, note_text if t_nom > t_mat else "Phần lớn đã map"),
        ('Lỗi/Sự cố',i_tot, i_mat, i_nom, i_rat, i_imp, i_not, note_text if i_nom > i_mat else "Phần lớn đã map"),
    ]

    for ri, (mod, tot, mat, nom, rat, imp, not_y, nt) in enumerate(summary_data, 2):
        ws1.row_dimensions[ri].height = 20
        _set_cell(ws1, ri, 1, mod,   fill=COLOR_YELLOW, font=FONT_BOLD)
        _set_cell(ws1, ri, 2, tot,   fill=None, align=ALIGN_CENTER)
        _set_cell(ws1, ri, 3, mat,   fill=COLOR_GREEN, align=ALIGN_CENTER)
        _set_cell(ws1, ri, 4, nom,   fill=COLOR_RED if nom > 0 else COLOR_GREEN, align=ALIGN_CENTER)
        _set_cell(ws1, ri, 5, rat,   fill=None, align=ALIGN_CENTER)
        _set_cell(ws1, ri, 6, imp,   fill=COLOR_GREEN, align=ALIGN_CENTER)
        _set_cell(ws1, ri, 7, not_y, fill=COLOR_YELLOW if not_y > 0 else COLOR_GREEN, align=ALIGN_CENTER)
        _set_cell(ws1, ri, 8, nt,    fill=None)

    # Dòng TỔNG CỘNG
    g_tot = p_tot + t_tot + i_tot
    g_mat = p_mat + t_mat + i_mat
    g_nom = p_nom + t_nom + i_nom
    g_rat = f"{g_mat / g_tot * 100:.1f}%" if g_tot else "0.0%"
    g_imp = p_imp + t_imp + i_imp
    g_not = p_not + t_not + i_not
    r_grand = 5
    ws1.row_dimensions[r_grand].height = 22
    _set_cell(ws1, r_grand, 1, 'TỔNG CỘNG', fill=COLOR_TOTAL, font=FONT_HEADER)
    _set_cell(ws1, r_grand, 2, g_tot,       fill=COLOR_TOTAL, font=FONT_HEADER, align=ALIGN_CENTER)
    _set_cell(ws1, r_grand, 3, g_mat,       fill=COLOR_TOTAL, font=FONT_HEADER, align=ALIGN_CENTER)
    _set_cell(ws1, r_grand, 4, g_nom,       fill=COLOR_TOTAL, font=FONT_HEADER, align=ALIGN_CENTER)
    _set_cell(ws1, r_grand, 5, g_rat,       fill=COLOR_TOTAL, font=FONT_HEADER, align=ALIGN_CENTER)
    _set_cell(ws1, r_grand, 6, g_imp,       fill=COLOR_TOTAL, font=FONT_HEADER, align=ALIGN_CENTER)
    _set_cell(ws1, r_grand, 7, g_not,       fill=COLOR_TOTAL, font=FONT_HEADER, align=ALIGN_CENTER)
    _set_cell(ws1, r_grand, 8, '',          fill=COLOR_TOTAL, font=FONT_HEADER)
    ws1.freeze_panes = 'A2'

    # ─────────────────────────────────────────────────────────────
    # SHEET 2: 'Dự án (th.project)'
    # ─────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet('Dự án (th.project)')
    headers2 = ['STT', 'Mã dự án (Excel)', 'Tên dự án (Excel)', 'ID DB', 'Tên DB', 'Trạng thái']
    widths2  = [5.0, 25.0, 45.0, 8.0, 45.0, 18.0]
    for ci, (h, w) in enumerate(zip(headers2, widths2), 1):
        _set_header(ws2, 1, ci, h, w)
    ws2.row_dimensions[1].height = 26

    for ri, rec in enumerate(proj_list, 2):
        matched = bool(rec.get('db_id'))
        fill = COLOR_GREEN if matched else COLOR_RED
        status_val = rec.get('status') or ('✅ Đã import' if matched else '❌ Không tìm thấy trong DB')
        ws2.row_dimensions[ri].height = 18
        _set_cell(ws2, ri, 1, ri - 1,                      fill=fill, align=ALIGN_CENTER)
        _set_cell(ws2, ri, 2, rec.get('excel_code') or '',fill=fill)
        _set_cell(ws2, ri, 3, rec.get('excel_name') or '',fill=fill)
        _set_cell(ws2, ri, 4, rec.get('db_id') or '—',    fill=fill, align=ALIGN_CENTER)
        _set_cell(ws2, ri, 5, rec.get('db_name') or '—',  fill=fill)
        _set_cell(ws2, ri, 6, status_val,                  fill=fill, align=ALIGN_CENTER)
    ws2.freeze_panes = 'A2'

    # ─────────────────────────────────────────────────────────────
    # SHEET 3: 'Đầu việc (th.project.task)'
    # ─────────────────────────────────────────────────────────────
    ws3 = wb.create_sheet('Đầu việc (th.project.task)')
    headers3 = ['STT', 'Mã dự án', 'Tên dự án', 'Tên task (Excel)', 'ID task DB', 'Mã task DB', 'Tên task DB']
    widths3  = [5.0, 20.0, 38.0, 42.0, 9.0, 22.0, 42.0]
    for ci, (h, w) in enumerate(zip(headers3, widths3), 1):
        _set_header(ws3, 1, ci, h, w)
    ws3.row_dimensions[1].height = 26

    for ri, rec in enumerate(task_list, 2):
        matched = bool(rec.get('db_id'))
        fill = COLOR_GREEN if matched else COLOR_RED
        ws3.row_dimensions[ri].height = 18
        _set_cell(ws3, ri, 1, ri - 1,                            fill=fill, align=ALIGN_CENTER)
        _set_cell(ws3, ri, 2, rec.get('excel_code') or '',      fill=fill)
        _set_cell(ws3, ri, 3, rec.get('excel_proj_name') or '', fill=fill)
        _set_cell(ws3, ri, 4, rec.get('excel_task') or '',      fill=fill)
        _set_cell(ws3, ri, 5, rec.get('db_id') or '—',          fill=fill, align=ALIGN_CENTER)
        _set_cell(ws3, ri, 6, rec.get('db_code') or '—',        fill=fill)
        _set_cell(ws3, ri, 7, rec.get('db_name') or '—',        fill=fill)
    ws3.freeze_panes = 'A2'

    # ─────────────────────────────────────────────────────────────
    # SHEET 4: 'Lỗi-Sự cố (th.issue)'
    # ─────────────────────────────────────────────────────────────
    ws4 = wb.create_sheet('Lỗi-Sự cố (th.issue)')
    headers4 = ['STT', 'Tên lỗi (Excel - hệ thống cũ)', 'ID issue DB', 'Tên issue DB', 'Trạng thái']
    widths4  = [5.0, 65.0, 10.0, 65.0, 22.0]
    for ci, (h, w) in enumerate(zip(headers4, widths4), 1):
        _set_header(ws4, 1, ci, h, w)
    ws4.row_dimensions[1].height = 26

    for ri, rec in enumerate(iss_list, 2):
        matched = bool(rec.get('db_id'))
        fill = COLOR_GREEN if matched else COLOR_RED
        status_val = rec.get('status') or ('✅ Đã import' if matched else '❌ Không tìm thấy')
        ws4.row_dimensions[ri].height = 18
        _set_cell(ws4, ri, 1, ri - 1,                     fill=fill, align=ALIGN_CENTER)
        _set_cell(ws4, ri, 2, rec.get('excel_name') or '',fill=fill)
        _set_cell(ws4, ri, 3, rec.get('db_id') or '—',   fill=fill, align=ALIGN_CENTER)
        _set_cell(ws4, ri, 4, rec.get('db_name') or '—', fill=fill)
        _set_cell(ws4, ri, 5, status_val,                 fill=fill, align=ALIGN_CENTER)
    ws4.freeze_panes = 'A2'

    wb.save(output_path)
    print(f"\n📄 Đã xuất file báo cáo mapping: {output_path}")


# =============================================================
#   MAIN IMPORT FUNCTION
# =============================================================

def run_import(env_config: dict, overrides: dict, dry_run: bool, scope: str, reset: bool, report_path: str = None, no_report: bool = False):
    cfg = {**env_config, **overrides}
    excel_path  = cfg["excel_path"]
    tz_offset   = int(cfg.get("timezone_offset_hours", 7))
    admin_login = cfg.get("admin_login", "admin")

    print(f"\n{'='*60}")
    print(f"🌐 Môi trường : {cfg.get('description', cfg['db_name'])}")
    print(f"🗄️  Database   : {cfg['db_name']} @ {cfg['db_host']}:{cfg['db_port']}")
    print(f"📄 Excel       : {excel_path}")
    print(f"⏰ Offset giờ  : +{tz_offset}h  |  Scope: {scope}  |  Dry-run: {dry_run}")
    print(f"{'='*60}")

    if not os.path.exists(excel_path):
        print(f"❌ Không tìm thấy: {excel_path}")
        sys.exit(1)

    # Kết nối DB
    try:
        conn = psycopg2.connect(
            host=cfg["db_host"], port=int(cfg["db_port"]),
            dbname=cfg["db_name"], user=cfg["db_user"], password=cfg["db_password"]
        )
        cur = conn.cursor()
        print(f"✅ Kết nối DB thành công")
    except Exception as e:
        print(f"❌ Lỗi kết nối: {e}"); sys.exit(1)

    # Kiểm tra module
    cur.execute("SELECT EXISTS (SELECT FROM information_schema.tables WHERE table_name='th_project')")
    if not cur.fetchone()[0]:
        print(f"❌ DB chưa cài module omg_project"); conn.close(); sys.exit(1)

    # Reset mode
    if reset:
        for mdl in ['th.project', 'th.project.task', 'th.issue']:
            cur.execute(
                "DELETE FROM mail_message WHERE model=%s AND (body LIKE %s OR body LIKE %s OR body LIKE %s)",
                (mdl, f'%{IMPORT_MARKER}%', '%Nhập từ file log Excel%', '%Dữ liệu được nhập từ file log Excel%')
            )
            print(f"🗑️  Đã xoá log import cũ cho {mdl}: {cur.rowcount} bản ghi")
        conn.commit()
        conn.close()
        return

    # Lấy subtype + author
    cur.execute("SELECT id FROM mail_message_subtype WHERE internal=true ORDER BY id LIMIT 1")
    r = cur.fetchone()
    subtype_id = r[0] if r else 2

    cur.execute("SELECT id, partner_id FROM res_users WHERE login=%s LIMIT 1", (admin_login,))
    r = cur.fetchone()
    if not r:
        cur.execute("SELECT id, partner_id FROM res_users WHERE id=1")
        r = cur.fetchone()
    author_user_id, author_partner_id = r
    print(f"👤 Author: user_id={author_user_id}, partner_id={author_partner_id}, subtype_id={subtype_id}\n")

    # Build DB lookups
    project_by_code, task_exact, task_stripped, task_name_in_proj, issue_by_name, issue_by_name_lower = build_db_lookups(cur)

    # Lấy danh sách record đã có log note từ trước
    cur.execute(
        "SELECT model, res_id FROM mail_message WHERE body LIKE %s OR body LIKE %s OR body LIKE %s",
        (f'%{IMPORT_MARKER}%', '%Nhập từ file log Excel%', '%Dữ liệu được nhập từ file log Excel%')
    )
    imported_set = set(cur.fetchall())  # set of (model, res_id)

    # Đọc Excel
    print("📂 Đọc file Excel...")
    wb = openpyxl.load_workbook(excel_path)

    totals = {"ok": 0, "skip": 0, "no_match": 0}
    report_data = {'projects': [], 'tasks': [], 'issues': []}

    # Luôn đọc đủ cả 3 sheet để lập báo cáo mapping toàn diện
    proj_data  = read_sheet_project(wb, tz_offset)
    task_data  = read_sheet_task(wb, tz_offset)
    issue_data = read_sheet_issue(wb, tz_offset)

    # ---- XỬ LÝ: PROJECT ----
    print(f"\n--- 🏗️  Dự án (th.project) ---")
    for ma_du_an, entries in sorted(proj_data.items()):
        db = project_by_code.get(ma_du_an) or project_by_code.get(ma_du_an.lower())
        entries_sorted = sorted(entries, key=lambda x: x["ngay_gio"] or datetime.min)
        ten_du_an = next((e["ten_du_an"] for e in entries_sorted if e["ten_du_an"]), ma_du_an)

        if not db:
            totals["no_match"] += 1
            report_data['projects'].append({
                'excel_code': ma_du_an,
                'excel_name': ten_du_an,
                'db_id':      None,
                'db_name':    None,
                'status':     '❌ Không tìm thấy trong DB'
            })
            continue

        project_id, project_name = db
        already = ('th.project', project_id) in imported_set

        if scope in ("all", "project"):
            if already:
                print(f"  ⏭️  [PROJ] [{ma_du_an}] skip (đã có log)")
                totals["skip"] += 1
            elif dry_run:
                print(f"  🔍 [PROJ] [{ma_du_an}] -> project_id={project_id} ({len(entries_sorted)} entries) [DRY]")
                totals["ok"] += 1
                already = True
            else:
                html = build_html_project(ma_du_an, ten_du_an, entries_sorted)
                if html:
                    mid = insert_lognote(cur, 'th.project', project_id, html, subtype_id, author_partner_id, author_user_id)
                    print(f"  ✅ [PROJ] [{ma_du_an}] -> project_id={project_id}, msg_id={mid} ({len(entries_sorted)} entries)")
                    totals["ok"] += 1
                    imported_set.add(('th.project', project_id))
                    already = True

        status_text = '✅ Đã import' if already else '🔄 Map được - chưa import'
        report_data['projects'].append({
            'excel_code': ma_du_an,
            'excel_name': ten_du_an,
            'db_id':      project_id,
            'db_name':    project_name,
            'status':     status_text
        })

    # ---- XỬ LÝ: TASK ----
    print(f"\n--- 📋 Đầu việc (th.project.task) ---")
    no_match_tasks = []
    for (ma_du_an, ten_task), entries in sorted(task_data.items()):
        db, strat = find_task(ma_du_an, ten_task, task_exact, task_stripped, task_name_in_proj)
        entries_sorted = sorted(entries, key=lambda x: x["ngay_gio"] or datetime.min)
        ten_du_an = next((e["ten_du_an"] for e in entries_sorted if e["ten_du_an"]), "")

        if not db:
            no_match_tasks.append((ma_du_an, ten_task, len(entries)))
            totals["no_match"] += 1
            report_data['tasks'].append({
                'excel_code':      ma_du_an,
                'excel_proj_name': ten_du_an,
                'excel_task':      ten_task,
                'db_id':           None,
                'db_code':         None,
                'db_name':         None,
                'status':          '❌ Không tìm thấy'
            })
            continue

        task_id, tname_db, tcode_db = db
        already = ('th.project.task', task_id) in imported_set

        if scope in ("all", "task"):
            if already:
                totals["skip"] += 1
            elif dry_run:
                print(f"  🔍 [TASK] [{ma_du_an}] '{ten_task[:45]}' -> task_id={task_id} ({strat}) [DRY]")
                totals["ok"] += 1
                already = True
            else:
                html = build_html_task(ma_du_an, ten_task, entries_sorted)
                if html:
                    mid = insert_lognote(cur, 'th.project.task', task_id, html, subtype_id, author_partner_id, author_user_id)
                    print(f"  ✅ [TASK] [{ma_du_an}] '{ten_task[:45]}' -> task_id={task_id} msg_id={mid} ({strat})")
                    totals["ok"] += 1
                    imported_set.add(('th.project.task', task_id))
                    already = True

        status_text = '✅ Đã import' if already else '🔄 Map được - chưa import'
        report_data['tasks'].append({
            'excel_code':      ma_du_an,
            'excel_proj_name': ten_du_an,
            'excel_task':      ten_task,
            'db_id':           task_id,
            'db_code':         tcode_db,
            'db_name':         tname_db,
            'status':          status_text
        })

    if no_match_tasks and scope in ("all", "task"):
        print(f"\n  ❌ Không khớp task ({len(no_match_tasks)}):")
        for m in no_match_tasks[:15]:
            print(f"      [{m[0]}] '{m[1][:55]}' ({m[2]} entries)")
        if len(no_match_tasks) > 15:
            print(f"      ... và {len(no_match_tasks)-15} task khác")

    # ---- XỬ LÝ: ISSUE ----
    print(f"\n--- 🐛 Lỗi/Sự cố (th.issue) ---")
    no_match_issues = []
    for ten_issue, entries in sorted(issue_data.items()):
        db, strat = find_issue(ten_issue, issue_by_name, issue_by_name_lower)
        entries_sorted = sorted(entries, key=lambda x: x["ngay_gio"] or datetime.min)

        if not db:
            no_match_issues.append((ten_issue, len(entries)))
            totals["no_match"] += 1
            report_data['issues'].append({
                'excel_name': ten_issue,
                'db_id':      None,
                'db_name':    None,
                'status':     '❌ Không tìm thấy'
            })
            continue

        issue_id, iname_db = db
        already = ('th.issue', issue_id) in imported_set

        if scope in ("all", "issue"):
            if already:
                totals["skip"] += 1
            elif dry_run:
                print(f"  🔍 [ISSUE] '{ten_issue[:55]}' -> issue_id={issue_id} ({strat}) [DRY]")
                totals["ok"] += 1
                already = True
            else:
                html = build_html_issue(ten_issue, entries_sorted)
                if html:
                    mid = insert_lognote(cur, 'th.issue', issue_id, html, subtype_id, author_partner_id, author_user_id)
                    print(f"  ✅ [ISSUE] '{ten_issue[:55]}' -> issue_id={issue_id} msg_id={mid} ({strat})")
                    totals["ok"] += 1
                    imported_set.add(('th.issue', issue_id))
                    already = True

        status_text = '✅ Đã import' if already else '🔄 Map được - chưa import'
        report_data['issues'].append({
            'excel_name': ten_issue,
            'db_id':      issue_id,
            'db_name':    iname_db,
            'status':     status_text
        })

    if no_match_issues and scope in ("all", "issue"):
        print(f"\n  ❌ Không khớp issue ({len(no_match_issues)}):")
        for m in no_match_issues[:15]:
            print(f"      '{m[0][:65]}' ({m[1]} entries)")
        if len(no_match_issues) > 15:
            print(f"      ... và {len(no_match_issues)-15} lỗi khác")

    # Commit vào DB
    if not dry_run:
        conn.commit()
        print(f"\n💾 Đã commit vào DB '{cfg['db_name']}'")
    else:
        print(f"\n⚠️  Dry-run: không có gì được lưu vào DB")

    conn.close()

    print(f"\n{'='*60}")
    print(f"📊 KẾT QUẢ IMPORT [{cfg.get('description', cfg['db_name'])}]")
    print(f"  ✅ Đã import mới : {totals['ok']}")
    print(f"  ⏭️  Bỏ qua (đã có): {totals['skip']}")
    print(f"  ❌ Không khớp   : {totals['no_match']}")
    print(f"{'='*60}")

    # Xuất file báo cáo Excel
    if not no_report:
        out_excel = report_path or os.path.join(os.path.dirname(os.path.abspath(excel_path)), "bao_cao_mapping_lognote.xlsx")
        export_report(report_data, out_excel)


# =============================================================
#   CLI
# =============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Import log note từ Excel vào Odoo (th.project / th.project.task / th.issue) và xuất báo cáo mapping",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument("--env",         default="default", help="Tên môi trường (default)")
    parser.add_argument("--scope",       default="all",     choices=["all", "project", "task", "issue"],
                        help="Phạm vi import: all / project / task / issue")
    parser.add_argument("--db",          help="Override DB name")
    parser.add_argument("--host",        help="Override DB host")
    parser.add_argument("--port",        help="Override DB port")
    parser.add_argument("--user",        help="Override DB user")
    parser.add_argument("--password",    help="Override DB password")
    parser.add_argument("--excel",       help="Override đường dẫn file Excel nguồn")
    parser.add_argument("--offset",      type=int, help="Override timezone offset (giờ)")
    parser.add_argument("--report",      help="Đường dẫn file Excel báo cáo mapping xuất ra (mặc định: bao_cao_mapping_lognote.xlsx)")
    parser.add_argument("--no-report",   action="store_true", help="Không xuất file Excel báo cáo")
    parser.add_argument("--dry-run",     action="store_true", help="Xem trước, không ghi DB")
    parser.add_argument("--reset",       action="store_true", help="Xoá toàn bộ log đã import (cẩn thận!)")
    parser.add_argument("--list-envs",   action="store_true", help="Liệt kê môi trường")
    args = parser.parse_args()

    if args.list_envs:
        print("\n📋 Môi trường đã cấu hình:")
        for name, cfg in ENVIRONMENTS.items():
            print(f"  [{name}] {cfg.get('description','')} -> {cfg['db_name']} @ {cfg['db_host']}:{cfg['db_port']}")
        return

    if args.env not in ENVIRONMENTS:
        print(f"❌ Môi trường '{args.env}' không tồn tại. Có: {', '.join(ENVIRONMENTS)}")
        sys.exit(1)

    overrides = {}
    if args.db:       overrides["db_name"]               = args.db
    if args.host:     overrides["db_host"]               = args.host
    if args.port:     overrides["db_port"]               = args.port
    if args.user:     overrides["db_user"]               = args.user
    if args.password: overrides["db_password"]           = args.password
    if args.excel:    overrides["excel_path"]            = args.excel
    if args.offset is not None: overrides["timezone_offset_hours"] = args.offset

    run_import(
        ENVIRONMENTS[args.env],
        overrides,
        dry_run=args.dry_run,
        scope=args.scope,
        reset=args.reset,
        report_path=args.report,
        no_report=args.no_report
    )


if __name__ == "__main__":
    main()

