import pandas as pd
import openpyxl
from datetime import datetime
import os

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(PROJECT_ROOT, "data", "excel_config.xlsx")

def format_excel_value(value):
    if pd.isna(value):
        return ""
    if hasattr(value, "strftime"):
        if getattr(value, "hour", 0) or getattr(value, "minute", 0) or getattr(value, "second", 0):
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return value.strftime("%Y-%m-%d")
    return str(value).strip()

def read_excel_data():
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"ไม่พบไฟล์ Config ที่: {EXCEL_PATH}")
    
    # 1. อ่าน Sheet Main เพื่อเอา Path หลัก (Row 2, Column B)
    df_main = pd.read_excel(EXCEL_PATH, sheet_name="Main", header=None)
    main_path = ""
    if len(df_main) >= 2 and df_main.shape[1] >= 2:
        main_path = str(df_main.iloc[1, 1]).strip()

    # 2. อ่านข้อมูลโดยไม่กำหนด header แล้วข้าม Row 1 และ 2 เองแบบ manual
    df_raw = pd.read_excel(EXCEL_PATH, sheet_name="Manage Permission", header=None)
    
    data_list = []
    numeric_ids = []
    data_row_count = 0
    for row_index in range(2, len(df_raw)):
        row = df_raw.iloc[row_index]
        path_value = str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else ""
        if not path_value or path_value.lower() in ["nan", "none", "path_ย่อย"]:
            continue
        data_row_count += 1
        raw_value = str(row[0]).strip() if pd.notna(row[0]) else ""
        if raw_value.replace(".", "", 1).isdigit():
            numeric_ids.append(int(float(raw_value)))

    next_id = max(max(numeric_ids, default=0), data_row_count) + 1
    ids_changed = False
    statuses_changed = False
    # วนลูปเริ่มตั้งแต่ Row Index 2 (ซึ่งก็คือ Row ที่ 3 ใน Excel)
    for idx in range(2, len(df_raw)):
        row = df_raw.iloc[idx]
        
        # คอลัมน์ A (Col 0) = ลำดับ
        # คอลัมน์ B-F (Col 1-5) = Path, Email, Role และช่วงเวลาให้สิทธิ์
        # คอลัมน์ G-I (Col 6-8) = สถานะ, Logs_Date, Remark
        
        raw_id = str(row[0]).strip() if pd.notna(row[0]) else ""
        path_sub = str(row[1]).strip() if pd.notna(row[1]) else ""
        email = str(row[2]).strip() if len(row) > 2 and pd.notna(row[2]) else ""
        role = str(row[3]).strip() if len(row) > 3 and pd.notna(row[3]) else ""
        start_date = format_excel_value(row[4]) if len(row) > 4 else ""
        end_date = format_excel_value(row[5]) if len(row) > 5 else ""
        status = str(row[6]).strip() if len(row) > 6 and pd.notna(row[6]) else ""
        if status.lower() == "error":
            status = ""
            statuses_changed = True
        logs_date = format_excel_value(row[7]) if len(row) > 7 else ""
        remark = str(row[8]).strip() if len(row) > 8 and pd.notna(row[8]) else ""

        # ถ้าไม่มีข้อมูล Path_ย่อย หรือ Email ให้ข้ามแถวนั้น
        if not path_sub or path_sub.lower() in ["nan", "none", "path_ย่อย"]:
            continue
            
        if raw_id.replace(".", "", 1).isdigit():
            row_id = int(float(raw_id))
        else:
            row_id = next_id
            next_id += 1
            ids_changed = True

        data_list.append({
            "id": row_id,
            "_excel_row": idx + 1,
            "path_sub": path_sub,
            "email": email,
            "role": role,
            "start_date": start_date,
            "end_date": end_date,
            "status": status,
            "logs_date": logs_date,
            "remark": remark
        })
        
    if ids_changed or statuses_changed:
        try:
            wb = openpyxl.load_workbook(EXCEL_PATH)
            ws = wb["Manage Permission"]
            for item in data_list:
                excel_row = item["_excel_row"]
                if ids_changed:
                    ws.cell(row=excel_row, column=1, value=item["id"])
                if statuses_changed and str(ws.cell(row=excel_row, column=7).value or "").strip().lower() == "error":
                    ws.cell(row=excel_row, column=7, value="")
            wb.save(EXCEL_PATH)
            wb.close()
        except PermissionError:
            raise Exception("กรุณาปิดไฟล์ excel_config.xlsx ก่อนทำรายการ")

    for item in data_list:
        item.pop("_excel_row", None)

    return {"main_path": main_path, "permissions": data_list}

def update_excel_logs(results):
    if not os.path.exists(EXCEL_PATH):
        return

    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        if "Manage Permission" not in wb.sheetnames:
            return

        ws = wb["Manage Permission"]
        today_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # แปลง ID ใน res_dict ให้เป็น string ทั้งหมดเพื่อป้องกัน Mismatch
        res_dict = {str(item["id"]): item for item in results}

        # วนลูปเริ่มตั้งแต่ Row ที่ 3 (ตรงกับข้อมูลจริงใน Excel)
        for row_idx in range(3, ws.max_row + 1):
            id_val = ws.cell(row=row_idx, column=1).value
            if id_val is not None:
                str_id = str(id_val).strip()
                if str_id.endswith('.0'):
                    str_id = str_id[:-2]
                    
                if str_id in res_dict:
                    res = res_dict[str_id]
                    if str(res.get("status", "")).strip().lower() == "error":
                        ws.cell(row=row_idx, column=9, value=res.get("remark", "")) # Col I: Remark เท่านั้น
                    else:
                        ws.cell(row=row_idx, column=7, value=res["status"])  # Col G: สถานะล่าสุด
                        ws.cell(row=row_idx, column=8, value=today_str)       # Col H: Logs_Date
                        ws.cell(row=row_idx, column=9, value=res.get("remark", "")) # Col I: Remark

        wb.save(EXCEL_PATH)
        wb.close()
    except PermissionError:
        raise Exception("กรุณาปิดไฟล์ excel_config.xlsx ก่อนทำรายการ")

def update_permission(permission_id, values):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["Manage Permission"]
    row_idx = find_permission_row(ws, permission_id)
    if row_idx is None:
        wb.close()
        raise ValueError(f"ไม่พบรายการ ID {permission_id}")

    for column, key in [(2, "path_sub"), (3, "email"), (4, "role"), (5, "start_date"), (6, "end_date")]:
        if key in values:
            ws.cell(row=row_idx, column=column, value=values[key])
    wb.save(EXCEL_PATH)
    wb.close()

def copy_permission(permission_id):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["Manage Permission"]
    source_row = find_permission_row(ws, permission_id)
    if source_row is None:
        wb.close()
        raise ValueError(f"ไม่พบรายการ ID {permission_id}")

    data_row_count = 0
    numeric_ids = []
    for row in range(3, ws.max_row + 1):
        path_value = ws.cell(row=row, column=2).value
        if path_value is not None and str(path_value).strip():
            data_row_count += 1
        value = ws.cell(row=row, column=1).value
        if value is not None and str(value).replace(".", "", 1).isdigit():
            numeric_ids.append(int(float(value)))

    new_id = max(max(numeric_ids, default=0), data_row_count) + 1
    target_row = ws.max_row + 1
    ws.cell(row=target_row, column=1, value=new_id)
    for column in range(2, 7):
        ws.cell(row=target_row, column=column, value=ws.cell(row=source_row, column=column).value)
    for column in range(7, 10):
        ws.cell(row=target_row, column=column, value="")
    wb.save(EXCEL_PATH)
    wb.close()
    return new_id

def delete_permission(permission_id):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["Manage Permission"]
    row_idx = find_permission_row(ws, permission_id)
    if row_idx is None:
        wb.close()
        raise ValueError(f"ไม่พบรายการ ID {permission_id}")

    ws.delete_rows(row_idx, 1)
    wb.save(EXCEL_PATH)
    wb.close()

def find_permission_row(ws, permission_id):
    expected = str(permission_id).strip()
    for row_idx in range(3, ws.max_row + 1):
        value = str(ws.cell(row=row_idx, column=1).value).strip()
        if value.endswith(".0"):
            value = value[:-2]
        if value == expected:
            return row_idx
    return None