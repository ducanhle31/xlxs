#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
import_legacy_data.py
======================
Import dữ liệu từ file Excel OpenProject (openproject_json.xlsx)
vào model th.legacy.data.wp và th.legacy.data.nb trong Odoo.

Cách dùng:
    python import_legacy_data.py [--env ENV] [--file FILE] [--dry-run] [--reset]

Tham số:
    --env     : Môi trường (default: local). Xem dict ENVIRONMENTS bên dưới.
    --file    : Đường dẫn file Excel (mặc định: openproject_json.xlsx cùng thư mục)
    --dry-run : Chạy thử, không ghi vào DB
    --reset   : Xóa toàn bộ dữ liệu cũ trước khi import

Yêu cầu:
    pip install psycopg2-binary openpyxl
"""

import argparse
import sys
import os
from datetime import datetime, date

try:
    import psycopg2
    import psycopg2.extras
except ImportError:
    print("❌ Thiếu thư viện psycopg2. Chạy: pip install psycopg2-binary")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("❌ Thiếu thư viện openpyxl. Chạy: pip install openpyxl")
    sys.exit(1)

# ══════════════════════════════════════════════════════════════
#  CẤU HÌNH MÔI TRƯỜNG
# ══════════════════════════════════════════════════════════════
ENVIRONMENTS = {
    "local": {
        "host": "127.0.0.1",
        "port": 5432,
        "dbname": "omega",
        "user": "odoo",
        "password": "1",
    },
    # Thêm môi trường khác tại đây:
    # "staging": {
    #     "host": "192.168.1.100",
    #     "port": 5432,
    #     "dbname": "omega_staging",
    #     "user": "odoo",
    #     "password": "yourpassword",
    # },
    # "production": {
    #     "host": "prod-server.example.com",
    #     "port": 5432,
    #     "dbname": "omega_prod",
    #     "user": "odoo",
    #     "password": "yourpassword",
    # },
}

DEFAULT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "openproject_json.xlsx")

# ══════════════════════════════════════════════════════════════
#  ĐỊNH NGHĨA SHEET VÀ LOẠI
# ══════════════════════════════════════════════════════════════
# Sheet dạng Work Package (WP) - 2 loại cấu trúc cột
# Type "op" = OP Sync (35 cột), type "wms" = WMS/Request (26 cột)
WP_SHEETS = {
    "OP Sync - Work Packages": "op",
    "WMS Task - WP": "wms",
    "WMS Bug - WP": "wms",
    "WMS Test Request - WP": "wms",
    "WMS BA Profile - WP": "wms",
    "Request Ticket - WP": "wms",
}

# Sheet dạng Notebook/Activity (13 cột, cấu trúc giống nhau)
NB_SHEETS = [
    "OP Sync - Activities (Notebook)",
    "WMS Task - Notebook",
    "WMS Bug - Notebook",
    "WMS Test Request - Notebook",
    "WMS BA Profile - Notebook",
    "Request Ticket - Notebook",
]

# ══════════════════════════════════════════════════════════════
#  HÀM HỖ TRỢ
# ══════════════════════════════════════════════════════════════

def safe_str(val, max_len=None):
    """Chuyển giá trị sang string, cắt nếu cần."""
    if val is None:
        return None
    result = str(val).strip()
    if not result:
        return None
    if max_len and len(result) > max_len:
        result = result[:max_len]
    return result


def safe_int(val):
    """Chuyển sang int, trả về None nếu lỗi."""
    if val is None:
        return None
    try:
        return int(val)
    except (ValueError, TypeError):
        return None


def safe_date(val):
    """Chuyển sang date string (YYYY-MM-DD), trả về None nếu lỗi."""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        if isinstance(val, datetime):
            return val.date().isoformat()
        return val.isoformat()
    s = str(val).strip()
    if not s:
        return None
    # Thử parse các định dạng phổ biến
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    return s  # Trả về nguyên bản nếu không parse được


def safe_datetime(val):
    """Chuyển sang datetime string (YYYY-MM-DD HH:MM:SS), trả về None nếu lỗi."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d %H:%M:%S")
    return safe_str(val)


def is_empty_row(vals):
    """Kiểm tra hàng có trống không."""
    return all(v is None or (isinstance(v, str) and not v.strip()) for v in vals)


def get_next_id(cur, table_name):
    """Lấy sequence ID tiếp theo cho table."""
    cur.execute(f"SELECT COALESCE(MAX(id), 0) + 1 FROM {table_name}")
    return cur.fetchone()[0]


# ══════════════════════════════════════════════════════════════
#  PARSE WP ROWS
# ══════════════════════════════════════════════════════════════

def parse_wp_op_row(vals, sheet_name):
    """
    Parse 1 row từ sheet OP Sync (35 cột):
    1:DB ID, 2:OP ID(DB), 3:Subject(DB), 4:Project(DB), 5:Status(DB), 6:Type(DB),
    7:Odoo Model, 8:Odoo Reference, 9:Is Processed, 10:Sync Date,
    11:OP ID(JSON), 12:Subject(JSON), 13:Description, 14:Status, 15:Type,
    16:Priority, 17:Project, 18:Author, 19:Assignee, 20:Parent,
    21:Ngày bắt đầu, 22:Ngày kết thúc, 23:Thời gian ước tính,
    24:Thời gian đã dùng, 25:% Hoàn thành, 26:Tự lên lịch,
    27:Tạo lúc, 28:Cập nhật lúc, 29:customField19, 30:customField20,
    31:customField31, 32:Chi phí nhân công, 33:Chi phí vật liệu,
    34:Tổng chi phí, 35:Raw JSON
    """
    return {
        "source_sheet": sheet_name,
        "db_id": safe_int(vals[0]),
        "record_name": safe_str(vals[2], 500),  # Subject(DB)
        "openproject_id": safe_str(vals[1], 100),
        "sync_status": None,
        "sync_data_id": None,
        "odoo_model": safe_str(vals[6], 200),
        "odoo_reference": safe_str(vals[7], 200),
        "is_processed": safe_str(vals[8], 100),
        "sync_date": safe_datetime(vals[9]),
        "op_id_json": safe_str(vals[10], 100),
        "subject": safe_str(vals[11], 500),
        "description": safe_str(vals[12]),
        "status": safe_str(vals[13], 200),
        "wp_type": safe_str(vals[14], 200),
        "priority": safe_str(vals[15], 100),
        "project": safe_str(vals[16], 500),
        "author": safe_str(vals[17], 200),
        "assignee": safe_str(vals[18], 200),
        "parent": safe_str(vals[19], 200),
        "date_start": safe_date(vals[20]),
        "date_end": safe_date(vals[21]),
        "estimated_time": safe_str(vals[22], 100),
        "spent_time": safe_str(vals[23], 100),
        "percent_done": safe_str(vals[24], 50),
        "auto_schedule": safe_str(vals[25], 50),
        "created_at": safe_str(vals[26], 100),
        "updated_at": safe_str(vals[27], 100),
        "custom_field_19": safe_str(vals[28], 500),
        "custom_field_20": safe_str(vals[29], 500),
        "custom_field_31": safe_str(vals[30], 500),
        "labor_cost": safe_str(vals[31], 100),
        "material_cost": safe_str(vals[32], 100),
        "total_cost": safe_str(vals[33], 100),
        "raw_json": safe_str(vals[34]),
    }


def parse_wp_wms_row(vals, sheet_name):
    """
    Parse 1 row từ sheet WMS/Request WP (27 cột):
    1:DB ID, 2:Tên record, 3:OpenProject ID, 4:Sync Status, 5:Sync Data ID,
    6:Trạng thái (odoo status), 7:OP ID(JSON), 8:Subject, 9:Description,
    10:Status, 11:Type, 12:Priority, 13:Project, 14:Author, 15:Assignee,
    16:Parent, 17:Ngày bắt đầu, 18:Ngày kết thúc, 19:Thời gian ước tính,
    20:Thời gian đã dùng, 21:% Hoàn thành, 22:Tạo lúc, 23:Cập nhật lúc,
    24:customField19, 25:customField20, 26:customField31, 27:Raw JSON
    """
    return {
        "source_sheet": sheet_name,
        "db_id": safe_int(vals[0]),
        "record_name": safe_str(vals[1], 500),
        "openproject_id": safe_str(vals[2], 100),
        "sync_status": safe_str(vals[3], 100),
        "sync_data_id": safe_str(vals[4], 100),
        "odoo_model": safe_str(vals[5], 200),   # Trạng thái (odoo)
        "odoo_reference": None,
        "is_processed": None,
        "sync_date": None,
        "op_id_json": safe_str(vals[6], 100),
        "subject": safe_str(vals[7], 500),
        "description": safe_str(vals[8]),
        "status": safe_str(vals[9], 200),
        "wp_type": safe_str(vals[10], 200),
        "priority": safe_str(vals[11], 100),
        "project": safe_str(vals[12], 500),
        "author": safe_str(vals[13], 200),
        "assignee": safe_str(vals[14], 200),
        "parent": safe_str(vals[15], 200),
        "date_start": safe_date(vals[16]),
        "date_end": safe_date(vals[17]),
        "estimated_time": safe_str(vals[18], 100),
        "spent_time": safe_str(vals[19], 100),
        "percent_done": safe_str(vals[20], 50),
        "auto_schedule": None,
        "created_at": safe_str(vals[21], 100),
        "updated_at": safe_str(vals[22], 100),
        "custom_field_19": safe_str(vals[23], 500),
        "custom_field_20": safe_str(vals[24], 500),
        "custom_field_31": safe_str(vals[25], 500),
        "labor_cost": None,
        "material_cost": None,
        "total_cost": None,
        "raw_json": safe_str(vals[26]) if len(vals) > 26 else None,
    }


def parse_nb_row(vals, sheet_name):
    """
    Parse 1 row từ sheet Notebook (13 cột):
    1:DB ID, 2:Tên record, 3:OpenProject ID, 4:Sync Status,
    5:Activity ID, 6:Activity Type, 7:Người thực hiện,
    8:Ghi chú(Comment), 9:Chi tiết thay đổi, 10:Version,
    11:Tạo lúc, 12:Cập nhật lúc, 13:Work Package
    """
    return {
        "source_sheet": sheet_name,
        "db_id": safe_int(vals[0]),
        "record_name": safe_str(vals[1], 500),
        "openproject_id": safe_str(vals[2], 100),
        "sync_status": safe_str(vals[3], 100),
        "activity_id": safe_str(vals[4], 100),
        "activity_type": safe_str(vals[5], 200),
        "actor": safe_str(vals[6], 200),
        "comment": safe_str(vals[7]),
        "change_detail": safe_str(vals[8]),
        "version": safe_str(vals[9], 50),
        "created_at": safe_str(vals[10], 100),
        "updated_at": safe_str(vals[11], 100),
        "work_package_name": safe_str(vals[12], 500),
    }


# ══════════════════════════════════════════════════════════════
#  IMPORT VÀO DB
# ══════════════════════════════════════════════════════════════

WP_INSERT_SQL = """
INSERT INTO th_legacy_data_wp (
    source_sheet, db_id, record_name, openproject_id, sync_status, sync_data_id,
    odoo_model, odoo_reference, is_processed, sync_date,
    op_id_json, subject, description, status, wp_type, priority,
    project, author, assignee, parent,
    date_start, date_end, estimated_time, spent_time, percent_done, auto_schedule,
    created_at, updated_at,
    custom_field_19, custom_field_20, custom_field_31,
    labor_cost, material_cost, total_cost, raw_json,
    create_uid, write_uid, create_date, write_date
) VALUES (
    %(source_sheet)s, %(db_id)s, %(record_name)s, %(openproject_id)s,
    %(sync_status)s, %(sync_data_id)s,
    %(odoo_model)s, %(odoo_reference)s, %(is_processed)s, %(sync_date)s,
    %(op_id_json)s, %(subject)s, %(description)s, %(status)s, %(wp_type)s, %(priority)s,
    %(project)s, %(author)s, %(assignee)s, %(parent)s,
    %(date_start)s, %(date_end)s, %(estimated_time)s, %(spent_time)s,
    %(percent_done)s, %(auto_schedule)s,
    %(created_at)s, %(updated_at)s,
    %(custom_field_19)s, %(custom_field_20)s, %(custom_field_31)s,
    %(labor_cost)s, %(material_cost)s, %(total_cost)s, %(raw_json)s,
    1, 1, NOW(), NOW()
)
"""

NB_INSERT_SQL = """
INSERT INTO th_legacy_data_nb (
    source_sheet, db_id, record_name, openproject_id, sync_status,
    activity_id, activity_type, actor, comment, change_detail,
    version, created_at, updated_at, work_package_name,
    create_uid, write_uid, create_date, write_date
) VALUES (
    %(source_sheet)s, %(db_id)s, %(record_name)s, %(openproject_id)s, %(sync_status)s,
    %(activity_id)s, %(activity_type)s, %(actor)s, %(comment)s, %(change_detail)s,
    %(version)s, %(created_at)s, %(updated_at)s, %(work_package_name)s,
    1, 1, NOW(), NOW()
)
"""


def import_wp_sheet(ws, sheet_name, sheet_type, cur, dry_run, batch_size=500):
    """Import dữ liệu Work Package từ 1 sheet (dùng iter_rows để nhanh hơn)."""
    inserted = 0
    skipped = 0
    batch = []
    row_num = 0

    # iter_rows(values_only=True) nhanh hơn nhiều so với ws.cell() từng ô
    rows_iter = ws.iter_rows(min_row=2, values_only=True)

    for vals in rows_iter:
        row_num += 1
        if is_empty_row(vals):
            continue

        try:
            if sheet_type == "op":
                record = parse_wp_op_row(vals, sheet_name)
            else:
                record = parse_wp_wms_row(vals, sheet_name)
            batch.append(record)
        except Exception as e:
            print(f"  ⚠️  Lỗi parse row {row_num + 1}: {e}")
            skipped += 1
            continue

        if len(batch) >= batch_size:
            if not dry_run:
                psycopg2.extras.execute_batch(cur, WP_INSERT_SQL, batch)
            inserted += len(batch)
            batch = []
            print(f"    ... đã xử lý {inserted:,} records (sheet: {sheet_name})")

    if batch:
        if not dry_run:
            psycopg2.extras.execute_batch(cur, WP_INSERT_SQL, batch)
        inserted += len(batch)

    return inserted, skipped


def import_nb_sheet(ws, sheet_name, cur, dry_run, batch_size=500):
    """Import dữ liệu Notebook từ 1 sheet (dùng iter_rows để nhanh hơn)."""
    inserted = 0
    skipped = 0
    batch = []
    row_num = 0

    rows_iter = ws.iter_rows(min_row=2, values_only=True)

    for vals in rows_iter:
        row_num += 1
        if is_empty_row(vals):
            continue

        try:
            record = parse_nb_row(vals, sheet_name)
            batch.append(record)
        except Exception as e:
            print(f"  ⚠️  Lỗi parse row {row_num + 1}: {e}")
            skipped += 1
            continue

        if len(batch) >= batch_size:
            if not dry_run:
                psycopg2.extras.execute_batch(cur, NB_INSERT_SQL, batch)
            inserted += len(batch)
            batch = []
            print(f"    ... đã xử lý {inserted:,} records (sheet: {sheet_name})")

    if batch:
        if not dry_run:
            psycopg2.extras.execute_batch(cur, NB_INSERT_SQL, batch)
        inserted += len(batch)

    return inserted, skipped


# ══════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Import dữ liệu OpenProject từ Excel vào Odoo (th.legacy.data.wp / th.legacy.data.nb)"
    )
    parser.add_argument("--env", default="local",
                        help=f"Môi trường DB (mặc định: local). Các lựa chọn: {list(ENVIRONMENTS.keys())}")
    parser.add_argument("--file", default=DEFAULT_FILE,
                        help=f"Đường dẫn file Excel (mặc định: {DEFAULT_FILE})")
    parser.add_argument("--dry-run", action="store_true",
                        help="Chạy thử, không ghi vào DB")
    parser.add_argument("--reset", action="store_true",
                        help="Xóa toàn bộ dữ liệu cũ trong bảng trước khi import")
    parser.add_argument("--sheet", default=None,
                        help="Chỉ import 1 sheet cụ thể (tên sheet, mặc định: tất cả)")
    args = parser.parse_args()

    # Validate môi trường
    if args.env not in ENVIRONMENTS:
        print(f"❌ Môi trường '{args.env}' không tồn tại. Các lựa chọn: {list(ENVIRONMENTS.keys())}")
        sys.exit(1)

    # Validate file
    if not os.path.exists(args.file):
        print(f"❌ Không tìm thấy file: {args.file}")
        sys.exit(1)

    db_cfg = ENVIRONMENTS[args.env]
    dry = args.dry_run

    print("=" * 60)
    print(f"  IMPORT DỮ LIỆU CŨ OPENPROJECT → ODOO")
    print("=" * 60)
    print(f"  Môi trường : {args.env} ({db_cfg['host']}:{db_cfg['port']}/{db_cfg['dbname']})")
    print(f"  File Excel  : {args.file}")
    print(f"  Dry-run     : {dry}")
    print(f"  Reset       : {args.reset}")
    print(f"  Sheet filter: {args.sheet or 'Tất cả'}")
    print("=" * 60)

    # Mở file Excel
    print("\n📂 Đang mở file Excel (có thể mất vài giây với file lớn)...")
    try:
        # read_only=True + iter_rows = nhanh nhất cho file lớn
        wb = openpyxl.load_workbook(args.file, read_only=True, data_only=True)
    except Exception as e:
        print(f"❌ Không mở được file Excel: {e}")
        sys.exit(1)
    print(f"   ✅ Mở thành công! Sheets: {wb.sheetnames}")

    # Kết nối DB
    print("\n🔌 Đang kết nối DB...")
    try:
        conn = psycopg2.connect(**db_cfg)
        conn.autocommit = False
        cur = conn.cursor()
        print("   ✅ Kết nối thành công!")
    except Exception as e:
        print(f"❌ Không kết nối được DB: {e}")
        sys.exit(1)

    total_wp = 0
    total_nb = 0
    total_skip = 0

    try:
        # Reset nếu cần
        if args.reset and not dry:
            print("\n🗑️  Đang xóa dữ liệu cũ...")
            cur.execute("DELETE FROM th_legacy_data_wp")
            cur.execute("DELETE FROM th_legacy_data_nb")
            conn.commit()
            print("   ✅ Đã xóa xong.")

        # Import Work Package sheets
        print("\n📊 Import Work Package sheets...")
        for sheet_name, sheet_type in WP_SHEETS.items():
            if args.sheet and sheet_name != args.sheet:
                continue
            if sheet_name not in wb.sheetnames:
                print(f"  ⚠️  Sheet '{sheet_name}' không có trong file, bỏ qua.")
                continue
            print(f"\n  ➡️  Sheet: {sheet_name} (type={sheet_type})")
            ws = wb[sheet_name]
            ins, skp = import_wp_sheet(ws, sheet_name, sheet_type, cur, dry)
            print(f"     ✅ Inserted: {ins} | Skipped: {skp}")
            total_wp += ins
            total_skip += skp
            if not dry:
                conn.commit()

        # Import Notebook sheets
        print("\n📝 Import Notebook/Activity sheets...")
        for sheet_name in NB_SHEETS:
            if args.sheet and sheet_name != args.sheet:
                continue
            if sheet_name not in wb.sheetnames:
                print(f"  ⚠️  Sheet '{sheet_name}' không có trong file, bỏ qua.")
                continue
            print(f"\n  ➡️  Sheet: {sheet_name}")
            ws = wb[sheet_name]
            ins, skp = import_nb_sheet(ws, sheet_name, cur, dry)
            print(f"     ✅ Inserted: {ins} | Skipped: {skp}")
            total_nb += ins
            total_skip += skp
            if not dry:
                conn.commit()

        # Kết quả
        print("\n" + "=" * 60)
        print("  KẾT QUẢ IMPORT")
        print("=" * 60)
        print(f"  Work Package records : {total_wp:,}")
        print(f"  Notebook records     : {total_nb:,}")
        print(f"  Tổng cộng            : {total_wp + total_nb:,}")
        print(f"  Bỏ qua (lỗi)        : {total_skip:,}")
        if dry:
            print("\n  ⚠️  DRY-RUN: Không có gì được ghi vào DB!")
        else:
            print("\n  ✅ Import hoàn thành!")
        print("=" * 60)

    except Exception as e:
        print(f"\n❌ Lỗi: {e}")
        if not dry:
            conn.rollback()
            print("   ↩️  Đã rollback.")
        raise
    finally:
        cur.close()
        conn.close()
        wb.close()


if __name__ == "__main__":
    main()

