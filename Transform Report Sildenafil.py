import os
import re
import sys
import shutil
import base64
from openpyxl import load_workbook
import pandas as pd
from datetime import datetime
import json
import subprocess

# Mengatur BASE_DIR agar kompatibel saat dijalankan sebagai script .py maupun file .exe (PyInstaller)
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DIR_SOURCE = os.path.join(BASE_DIR, "source")
DIR_IMAGE = os.path.join(BASE_DIR, "image")
DIR_OUTPUT = os.path.join(BASE_DIR, "output")
DIR_LOGO = os.path.join(BASE_DIR, "Logo Image")
FILE_MAPPING = os.path.join(BASE_DIR, "mapping_surveyor.xlsx")
FILE_LOG = os.path.join(BASE_DIR, "log_proses_survey.xlsx")
FILE_HTML_PETA = os.path.join(BASE_DIR, "index.html")

for direktori in [DIR_SOURCE, DIR_IMAGE, DIR_OUTPUT]:
    if not os.path.exists(direktori):
        os.makedirs(direktori)

# Load Logo Images as Base64 (if available)
logo_50_base64 = ""
logo_100_base64 = ""
logo_50_path = os.path.join(DIR_LOGO, "Sildenafil 50 mg - Front View.png")
logo_100_path = os.path.join(DIR_LOGO, "Sildenafil 100 mg - Front View.png")

if os.path.exists(logo_50_path):
    try:
        with open(logo_50_path, "rb") as f_logo:
            logo_50_base64 = base64.b64encode(f_logo.read()).decode("utf-8")
    except Exception as e:
        print(f"[WARNING] Gagal membaca logo 50 mg: {e}")

if os.path.exists(logo_100_path):
    try:
        with open(logo_100_path, "rb") as f_logo:
            logo_100_base64 = base64.b64encode(f_logo.read()).decode("utf-8")
    except Exception as e:
        print(f"[WARNING] Gagal membaca logo 100 mg: {e}")

if not os.path.exists(FILE_MAPPING):
    df_template = pd.DataFrame(columns=["Surveyor ID", "Sales Name", "Position", "Region", "ASM", "STATUS"])
    df_template.to_excel(FILE_MAPPING, index=False)
    print(f"[INFO] File '{os.path.basename(FILE_MAPPING)}' baru saja dibuat.")
    print("-> Silakan isi data Surveyor kamu di file Excel tersebut dahulu, lalu jalankan ulang script ini.")
    exit()

print("-> Memuat data master dari Excel mapping...")
df_mapping_excel = pd.read_excel(FILE_MAPPING)

# Cari nama kolom Status secara case-insensitive (misal: STATUS, Status, status)
status_col_name = None
for col in df_mapping_excel.columns:
    if str(col).strip().lower() == "status":
        status_col_name = col
        break

MASTER_SURVEYOR = {}
for idx, row in df_mapping_excel.iterrows():
    if pd.isna(row["Surveyor ID"]):
        continue
    sid = str(row["Surveyor ID"]).strip()
    raw_status = str(row[status_col_name]).strip() if (status_col_name and not pd.isna(row[status_col_name])) else "Active"
    norm_status = "Inactive" if raw_status.lower() == "inactive" else "Active"
    MASTER_SURVEYOR[sid] = {
        "Sales Name": str(row["Sales Name"]).strip() if not pd.isna(row["Sales Name"]) else "Unknown",
        "Position": str(row["Position"]).strip() if not pd.isna(row["Position"]) else "-",
        "Region": str(row["Region"]).strip() if not pd.isna(row["Region"]) else "-",
        "ASM": str(row["ASM"]).strip() if not pd.isna(row["ASM"]) else "-",
        "Status": norm_status
    }

files_in_source = [f for f in os.listdir(DIR_SOURCE) if f.endswith('.xlsx') and not f.startswith('~$')]

if not files_in_source:
    print(f"[ERROR] Tidak ditemukan file .xlsx di: {DIR_SOURCE}")
    exit()

print(f"Ditemukan {len(files_in_source)} file Excel untuk diproses.\n")

outlet_aggregated_data = {}
kolom_headers = []
list_log_detail = []

total_foto_diekstrak = 0
total_baris_digabung = 0
total_foto_gagal = 0
total_tanpa_foto = 0
total_duplikat_dibuang = 0

seen_exact_rows = set()

for nama_file_excel in files_in_source:
    path_excel_input = os.path.join(DIR_SOURCE, nama_file_excel)
    print(f"==================================================")
    print(f"-> Membaca file: {nama_file_excel}")
    print(f"==================================================")

    wb_image = load_workbook(path_excel_input, data_only=True)
    ws_image = wb_image.active

    if not kolom_headers:
        kolom_headers = [ws_image.cell(row=2, column=col).value for col in range(1, ws_image.max_column + 1)]

    lokasi_col_idx = None
    for col_idx, header in enumerate(kolom_headers, start=1):
        if header:
            h_str = str(header).lower().strip()
            if ('survey' in h_str and ('loc' in h_str or 'coord' in h_str or 'koordinat' in h_str or 'lat' in h_str)) or (h_str in ['survey location', 'lokasi survey', 'survey_location']):
                lokasi_col_idx = col_idx
                break
    
    if lokasi_col_idx is None:
        for col_idx, header in enumerate(kolom_headers, start=1):
            if header:
                h_str = str(header).lower().strip()
                if 'survey' in h_str and 'visit' not in h_str:
                    lokasi_col_idx = col_idx
                    break

    gambar_per_baris = {}
    for img in ws_image._images:
        baris_gambar = img.anchor._from.row + 1
        gambar_per_baris[baris_gambar] = img

    for row in range(3, ws_image.max_row + 1):
        timestamp_raw = ws_image.cell(row=row, column=1).value
        surveyor_id_raw = ws_image.cell(row=row, column=2).value
        surveyor = ws_image.cell(row=row, column=3).value
        surveyee_id = ws_image.cell(row=row, column=4).value
        surveyee = ws_image.cell(row=row, column=5).value

        if not surveyor or not surveyee_id or not surveyee:
            continue
            
        timestamp_str = str(timestamp_raw).strip() if timestamp_raw else f"Tanpa_Timestamp_{row}"
        surveyor_id_str = str(surveyor_id_raw).strip() if surveyor_id_raw else ""
        surveyor_str = str(surveyor).strip()
        surveyee_id_str = str(surveyee_id).strip()
        surveyee_str = str(surveyee).strip()

        exact_row_signature = (timestamp_str, surveyor_id_str, surveyor_str, surveyee_id_str, surveyee_str)
        if exact_row_signature in seen_exact_rows:
            total_duplikat_dibuang += 1
            list_log_detail.append({
                "Waktu Eksekusi": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Nama File Source": nama_file_excel,
                "Nomor Baris": row,
                "Timestamp Data": timestamp_str,
                "Surveyor ID": surveyor_id_str,
                "Surveyor": surveyor_str,
                "Surveyee ID": surveyee_id_str,
                "Surveyee / Outlet": surveyee_str,
                "Status Baris": "DIHAPUS (Duplikat Identik)",
                "Status Foto": "-",
                "Path / Keterangan": "Baris duplikat murni dihapus"
            })
            continue
        
        seen_exact_rows.add(exact_row_signature)

        row_values = [ws_image.cell(row=row, column=col).value for col in range(1, ws_image.max_column + 1)]
        
        lat, lon = None, None
        def parse_and_fix_latlon(val_str):
            if not val_str:
                return None, None
            parts = re.split(r'[,\s]+', str(val_str).strip())
            floats = []
            for p in parts:
                try:
                    floats.append(float(p))
                except ValueError:
                    continue
            if len(floats) >= 2:
                v1, v2 = floats[0], floats[1]
                if (95.0 <= abs(v1) <= 141.0) and (-11.0 <= v2 <= 6.0):
                    return v2, v1
                elif (-11.0 <= v1 <= 6.0) and (95.0 <= abs(v2) <= 141.0):
                    return v1, v2
            return None, None

        if lokasi_col_idx:
            loc_val = ws_image.cell(row=row, column=lokasi_col_idx).value
            lat, lon = parse_and_fix_latlon(loc_val)

        row_qty = 0
        for idx_col, val_col in enumerate(row_values):
            header_name = str(kolom_headers[idx_col]).lower() if idx_col < len(kolom_headers) else ""
            if any(k in header_name for k in ['qty', 'jumlah', 'total', 'pack', 'pcs', 'quantity']):
                if pd.notna(val_col):
                    nums = re.findall(r'\d+', str(val_col).replace(',', '.'))
                    if nums:
                        try:
                            row_qty += float(nums[0])
                        except:
                            pass

        path_simpan_relatif_list = []
        status_foto_log = "-"
        keterangan_foto_log = "-"

        if row in gambar_per_baris:
            surveyor_clean = re.sub(r'[\/*?:"<>|]', '', surveyor_str).strip(' .')
            surveyee_clean = re.sub(r'[\/*?:"<>|]', '', surveyee_str).strip(' .')
            surveyee_id_clean = re.sub(r'[\/*?:"<>|]', '', surveyee_id_str).strip(' .')
            timestamp_clean = re.sub(r'[\s:]', '_', timestamp_str).strip(' .')

            surveyee_clean_short = surveyee_clean[:35].strip()

            folder_tujuan_foto = os.path.join(DIR_IMAGE, surveyor_clean, surveyee_clean_short)
            if not os.path.exists(folder_tujuan_foto):
                os.makedirs(folder_tujuan_foto)

            nama_file_foto = f"{surveyee_id_clean}_{surveyee_clean_short}_{timestamp_clean}_r{row}.jpg"
            path_simpan_foto = os.path.join(folder_tujuan_foto, nama_file_foto)
            path_simpan_relatif = f"image/{surveyor_clean}/{surveyee_clean_short}/{nama_file_foto}".replace(os.sep, "/")

            img_obj = gambar_per_baris[row]
            try:
                if hasattr(img_obj, 'image'):
                    img_obj.image.save(path_simpan_foto)
                elif hasattr(img_obj, 'ref') and hasattr(img_obj.ref, 'read'):
                    img_obj.ref.seek(0)
                    with open(path_simpan_foto, 'wb') as f:
                        f.write(img_obj.ref.read())
                elif hasattr(img_obj, 'fp') and hasattr(img_obj.fp, 'read'):
                    img_obj.fp.seek(0)
                    with open(path_simpan_foto, 'wb') as f:
                        f.write(img_obj.fp.read())
                
                total_foto_diekstrak += 1
                path_simpan_relatif_list.append(path_simpan_relatif)
                status_foto_log = "SUKSES"
                keterangan_foto_log = path_simpan_foto
            except Exception as e:
                total_foto_gagal += 1
                status_foto_log = "GAGAL"
                keterangan_foto_log = str(e)
        else:
            total_tanpa_foto += 1
            status_foto_log = "TIDAK ADA FOTO"
            keterangan_foto_log = "Tidak ada objek gambar pada baris ini"

        outlet_key = (surveyor_id_str, surveyor_str, surveyee_id_str, surveyee_str)

        if outlet_key in outlet_aggregated_data:
            total_baris_digabung += 1
            existing_qty = outlet_aggregated_data[outlet_key]["qty_numeric"]
            new_qty = existing_qty + row_qty
            outlet_aggregated_data[outlet_key]["qty_numeric"] = new_qty
            outlet_aggregated_data[outlet_key]["qty"] = int(new_qty) if new_qty.is_integer() else new_qty
            
            if path_simpan_relatif_list:
                for p in path_simpan_relatif_list:
                    if p not in outlet_aggregated_data[outlet_key]["foto_list"]:
                        outlet_aggregated_data[outlet_key]["foto_list"].append(p)
            
            if timestamp_str not in outlet_aggregated_data[outlet_key]["timestamp"]:
                outlet_aggregated_data[outlet_key]["timestamp"] = f"{outlet_aggregated_data[outlet_key]['timestamp']}, {timestamp_str}"

            list_log_detail.append({
                "Waktu Eksekusi": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Nama File Source": nama_file_excel,
                "Nomor Baris": row,
                "Timestamp Data": timestamp_str,
                "Surveyor ID": surveyor_id_str,
                "Surveyor": surveyor_str,
                "Surveyee ID": surveyee_id_str,
                "Surveyee / Outlet": surveyee_str,
                "Status Baris": "DIGABUNG (Akumulasi Qty Waktu Berbeda)",
                "Status Foto": status_foto_log,
                "Path / Keterangan": f"Qty digabung jadi {outlet_aggregated_data[outlet_key]['qty']}"
            })
        else:
            outlet_aggregated_data[outlet_key] = {
                "surveyor_id": surveyor_id_str,
                "surveyor": surveyor_str,
                "surveyee_id": surveyee_id_str,
                "surveyee": surveyee_str,
                "timestamp": timestamp_str,
                "lat": lat,
                "lon": lon,
                "qty_numeric": row_qty,
                "qty": int(row_qty) if row_qty.is_integer() else (row_qty if row_qty > 0 else "-"),
                "foto_list": path_simpan_relatif_list,
                "row_values": row_values
            }
            list_log_detail.append({
                "Waktu Eksekusi": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Nama File Source": nama_file_excel,
                "Nomor Baris": row,
                "Timestamp Data": timestamp_str,
                "Surveyor ID": surveyor_id_str,
                "Surveyor": surveyor_str,
                "Surveyee ID": surveyee_id_str,
                "Surveyee / Outlet": surveyee_str,
                "Status Baris": "BARU (Tersimpan)",
                "Status Foto": status_foto_log,
                "Path / Keterangan": keterangan_foto_log
            })

print(f"Total outlet unik setelah agregasi: {len(outlet_aggregated_data)} outlet")

qty_col_indices = []
for idx_col, header in enumerate(kolom_headers):
    header_name = str(header).lower() if header else ""
    if any(k in header_name for k in ['qty', 'jumlah', 'total', 'pack', 'pcs', 'quantity']):
        qty_col_indices.append(idx_col)

semua_data_gabungan = []
list_data_peta = []

for k, v in outlet_aggregated_data.items():
    current_row_values = list(v["row_values"])
    for q_idx in qty_col_indices:
        if q_idx < len(current_row_values):
            current_row_values[q_idx] = v["qty"]
            
    semua_data_gabungan.append(current_row_values)
    
    sid = v["surveyor_id"]
    reg = MASTER_SURVEYOR.get(sid, {}).get("Region", "-")
    asm = MASTER_SURVEYOR.get(sid, {}).get("ASM", "-")
    pos = MASTER_SURVEYOR.get(sid, {}).get("Position", "-")
    st = MASTER_SURVEYOR.get(sid, {}).get("Status", "Active")

    list_data_peta.append({
        "lat": v["lat"],
        "lon": v["lon"],
        "timestamp": v["timestamp"],
        "surveyor": v["surveyor"],
        "surveyor_id": v["surveyor_id"],
        "surveyee": v["surveyee"],
        "surveyee_id": v["surveyee_id"],
        "region": reg,
        "asm": asm,
        "position": pos,
        "status": st,
        "qty": v["qty_numeric"],
        "foto_list": v["foto_list"]
    })

if semua_data_gabungan:
    df_clean = pd.DataFrame(semua_data_gabungan, columns=kolom_headers)

    list_lat_excel = []
    list_lon_excel = []
    def parse_and_fix_latlon(val_str):
        if not val_str:
            return None, None
        parts = re.split(r'[,\s]+', str(val_str).strip())
        floats = []
        for p in parts:
            try:
                floats.append(float(p))
            except ValueError:
                continue
        if len(floats) >= 2:
            v1, v2 = floats[0], floats[1]
            if (95.0 <= abs(v1) <= 141.0) and (-11.0 <= v2 <= 6.0):
                return v2, v1
            elif (-11.0 <= v1 <= 6.0) and (95.0 <= abs(v2) <= 141.0):
                return v1, v2
        return None, None

    for idx, row in df_clean.iterrows():
        loc_val = row.iloc[lokasi_col_idx - 1] if lokasi_col_idx and lokasi_col_idx <= len(row) else None
        lt, ln = parse_and_fix_latlon(loc_val)
        list_lat_excel.append(lt if lt is not None else "-")
        list_lon_excel.append(ln if ln is not None else "-")

    df_clean["Latitude"] = list_lat_excel
    df_clean["Longitude"] = list_lon_excel

    list_position = []
    list_region = []
    list_asm = []
    list_status = []

    for idx, row in df_clean.iterrows():
        sid = str(row['Surveyor ID']).strip() if 'Surveyor ID' in df_clean.columns else ""
        if sid in MASTER_SURVEYOR:
            list_position.append(MASTER_SURVEYOR[sid]["Position"])
            list_region.append(MASTER_SURVEYOR[sid]["Region"])
            list_asm.append(MASTER_SURVEYOR[sid]["ASM"])
            list_status.append(MASTER_SURVEYOR[sid]["Status"])
        else:
            list_position.append("-")
            list_region.append("-")
            list_asm.append("-")
            list_status.append("Active")

    df_clean["Position"] = list_position
    df_clean["Region"] = list_region
    df_clean["ASM"] = list_asm
    df_clean["Status"] = list_status

    path_excel_output = os.path.join(DIR_OUTPUT, "all_data_combined_clean.xlsx")
    
    try:
        df_clean.to_excel(path_excel_output, index=False)
        df_log = pd.DataFrame(list_log_detail)
        df_log.to_excel(FILE_LOG, index=False)
        
        master_list_json = []
        for sid, info in MASTER_SURVEYOR.items():
            master_list_json.append({
                "surveyor_id": sid,
                "surveyor": info["Sales Name"],
                "position": info["Position"],
                "region": info["Region"],
                "asm": info["ASM"],
                "status": info["Status"]
            })

        latest_dt = None
        for fn_src in files_in_source:
            p_src = os.path.join(DIR_SOURCE, fn_src)
            try:
                wb_tmp = load_workbook(p_src, data_only=True)
                ws_tmp = wb_tmp.active
                for r_idx in range(3, ws_tmp.max_row + 1):
                    val_t = ws_tmp.cell(row=r_idx, column=1).value
                    if val_t is not None:
                        if isinstance(val_t, datetime):
                            if latest_dt is None or val_t > latest_dt:
                                latest_dt = val_t
                        else:
                            ts_str = str(val_t).strip()
                            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S.%f"):
                                try:
                                    parsed = datetime.strptime(ts_str, fmt)
                                    if latest_dt is None or parsed > latest_dt:
                                        latest_dt = parsed
                                    break
                                except ValueError:
                                    continue
            except Exception as e_ts:
                print(f"[WARNING] Gagal membaca timestamp dari {fn_src}: {e_ts}")

        if latest_dt:
            months_en = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            m_str = months_en[latest_dt.month - 1]
            last_update_str = f"{latest_dt.strftime('%d')} {m_str} {latest_dt.strftime('%Y - %H:%M')}"
        else:
            now_dt = datetime.now()
            months_en = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            last_update_str = f"{now_dt.strftime('%d')} {months_en[now_dt.month - 1]} {now_dt.strftime('%Y - %H:%M')}"

        payload_data = {
            "surveyData": list_data_peta,
            "masterSurveyor": master_list_json,
            "lastUpdate": last_update_str
        }
        data_json = json.dumps(payload_data)

        html_template = """<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Dashboard Sildenafil Package - Performance & GIS Map</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        *, *::before, *::after { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
        html, body { height: 100%; margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f1f5f9; color: #1e293b; overflow-x: hidden; }
        
        /* Navigation / Header */
        header {
            min-height: 70px;
            background-color: #ffffff;
            color: #0f172a;
            padding: 10px 18px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px solid #cbd5e1;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 3000;
            flex-wrap: wrap;
            gap: 8px 12px;
        }
        header h1 {
            margin: 0;
            font-size: clamp(13px, 1.5vw, 16px);
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 8px;
            flex-wrap: wrap;
        }
        .header-logos {
            display: inline-flex;
            gap: 6px;
            align-items: center;
            margin-left: 4px;
            vertical-align: middle;
        }
        .header-logos img {
            height: clamp(38px, 5vw, 56px);
            width: auto;
            object-fit: contain;
            border-radius: 4px;
        }
        .header-right {
            display: flex;
            align-items: center;
            gap: 10px;
            flex-wrap: wrap;
        }
        .last-update-pill {
            font-size: 11px;
            font-weight: normal;
            color: #64748b;
            background: #f8fafc;
            padding: 4px 8px;
            border: 1px solid #cbd5e1;
            border-radius: 4px;
            white-space: nowrap;
        }
        .nav-buttons {
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            align-items: center;
        }
        .nav-btn {
            background: #e2e8f0;
            color: #334155;
            border: none;
            padding: 7px 12px;
            border-radius: 6px;
            font-size: 11px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            white-space: nowrap;
        }
        .nav-btn.active { background: #2563eb; color: white; }
        .nav-btn:hover { opacity: 0.92; }

        /* Lang Switcher Buttons */
        .lang-switch { display: flex; gap: 4px; background: #e2e8f0; padding: 2px; border-radius: 6px; }
        .lang-btn { background: transparent; border: none; padding: 4px 8px; font-size: 10px; font-weight: bold; cursor: pointer; border-radius: 4px; color: #475569; }
        .lang-btn.active { background: #ffffff; color: #2563eb; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }

        /* Views Container - offset dinamis mengikuti tinggi header aktual */
        .view-container {
            --header-offset: 88px;
            margin-top: var(--header-offset);
            height: calc(100dvh - var(--header-offset));
            width: 100%;
            position: relative;
            overflow: hidden;
        }

        .view-page {
            position: absolute;
            inset: 0;
            width: 100%;
            height: 100%;
            display: none;
            background: #f1f5f9;
            overflow-y: auto;
            overflow-x: hidden;
        }
        .view-page.active { display: flex; flex-direction: column; }

        /* Summary Table View Styles */
        .table-content {
            padding: 18px 16px 28px 16px;
            max-width: 1400px;
            margin: 0 auto;
            width: 100%;
        }
        .filter-bar {
            background: white;
            padding: 14px;
            border-radius: 8px;
            border: 1px solid #cbd5e1;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(175px, 1fr));
            gap: 12px;
            align-items: end;
            margin-bottom: 18px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        }
        .filter-field { display: flex; flex-direction: column; gap: 4px; min-width: 0; }
        .filter-field label { font-size: 11px; font-weight: 600; color: #64748b; }
        .filter-field select {
            width: 100%;
            padding: 8px 10px;
            border-radius: 6px;
            border: 1px solid #cbd5e1;
            font-size: 12px;
            background: #f8fafc;
            outline: none;
        }
        .filter-actions {
            display: flex;
            align-items: stretch;
            gap: 8px;
            flex-wrap: wrap;
            grid-column: 1 / -1;
        }
        @media (min-width: 900px) {
            .filter-actions {
                grid-column: auto;
                align-items: flex-end;
                justify-content: flex-end;
            }
        }
        .btn-reset {
            padding: 8px 14px;
            background: #cbd5e1;
            color: #0f172a;
            border: none;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            white-space: nowrap;
        }
        .btn-excel {
            padding: 8px 14px;
            background: #059669;
            color: white;
            border: none;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            white-space: nowrap;
        }
        
        .summary-cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(210px, 1fr));
            gap: 14px;
            margin-bottom: 18px;
        }
        .scard {
            background: white;
            padding: 14px 16px;
            border-radius: 8px;
            border: 1px solid #cbd5e1;
            box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        }
        .scard h4 { margin: 0 0 6px 0; font-size: 11px; color: #64748b; text-transform: uppercase; line-height: 1.35; }
        .scard .sval { font-size: clamp(17px, 2vw, 22px); font-weight: bold; color: #0f172a; word-break: break-word; }

        /* Toggle Summary Buttons */
        .summary-toggle-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; flex-wrap: wrap; gap: 10px; }
        .summary-title { font-weight: 600; font-size: 14px; color: #0f172a; }
        .toggle-group { display: flex; gap: 6px; background: #e2e8f0; padding: 3px; border-radius: 6px; }
        .toggle-btn { background: transparent; border: none; padding: 5px 12px; font-size: 11px; font-weight: 600; cursor: pointer; border-radius: 4px; color: #475569; transition: all 0.2s; white-space: nowrap; }
        .toggle-btn.active { background: #2563eb; color: #ffffff; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }

        .table-responsive {
            background: white;
            border-radius: 8px;
            border: 1px solid #cbd5e1;
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
            box-shadow: 0 2px 4px rgba(0,0,0,0.02);
            margin-bottom: 22px;
        }
        table { width: 100%; border-collapse: collapse; text-align: left; font-size: 12px; min-width: 520px; }
        th, td { padding: 10px 12px; border-bottom: 1px solid #e2e8f0; }
        th { background: #f8fafc; font-weight: 600; color: #475569; position: sticky; top: 0; z-index: 10; }
        tr:hover { background: #f8fafc; }
        .total-row { background: #fefce8 !important; font-weight: bold; border-top: 2px solid #fde047; border-bottom: 2px solid #fde047; }
        .badge { padding: 3px 8px; border-radius: 4px; font-size: 10px; font-weight: 600; display: inline-block; }
        .badge-sbd { background: #dbeafe; color: #1e40af; }
        .badge-msd { background: #fef3c7; color: #92400e; }
        .badge-asm { background: #d1fae5; color: #065f46; }
        .badge-tss { background: #fae8ff; color: #86198f; }

        /* Label Status persis gaya badge Role: hijau Active, merah Inactive */
        .status-badge-active { background: #dcfce7; color: #15803d; padding: 3px 8px; border-radius: 4px; font-size: 10px; font-weight: 600; margin-left: 6px; display: inline-block; cursor: help; }
        .status-badge-inactive { background: #fee2e2; color: #b91c1c; padding: 3px 8px; border-radius: 4px; font-size: 10px; font-weight: 600; margin-left: 6px; display: inline-block; cursor: help; }

        .bottom-cta { margin-top: 10px; display: flex; justify-content: flex-end; }
        .btn-open-map {
            padding: 10px 18px;
            background: #2563eb;
            color: white;
            border: none;
            border-radius: 6px;
            font-weight: 600;
            cursor: pointer;
            font-size: 12px;
        }

        /* Map View Styles */
        .map-layout { display: flex; height: 100%; width: 100%; position: relative; overflow: hidden; }
        .sidebar {
            width: clamp(280px, 28vw, 390px);
            flex-shrink: 0;
            background-color: #ffffff;
            border-right: 1px solid #cbd5e1;
            display: flex;
            flex-direction: column;
            padding: 14px;
            gap: 10px;
            z-index: 1000;
            box-shadow: 4px 0 15px rgba(0,0,0,0.05);
            overflow-y: auto;
            -webkit-overflow-scrolling: touch;
        }
        .sidebar-filter { display: flex; flex-direction: column; gap: 3px; }
        .sidebar-filter label { font-size: 11px; font-weight: 600; color: #64748b; }
        .sidebar-filter select {
            width: 100%;
            padding: 7px 10px;
            border-radius: 6px;
            border: 1px solid #cbd5e1;
            background-color: #f8fafc;
            color: #1e293b;
            font-size: 12px;
            outline: none;
        }
        .outlet-list {
            flex: 1;
            min-height: 120px;
            overflow-y: auto;
            -webkit-overflow-scrolling: touch;
            display: flex;
            flex-direction: column;
            gap: 6px;
            padding-right: 2px;
        }
        .outlet-item { background-color: #f8fafc; border: 1px solid #cbd5e1; padding: 8px 10px; border-radius: 6px; cursor: pointer; transition: background-color 0.15s ease, border-color 0.15s ease; }
        .outlet-item:hover { border-color: #2563eb; background-color: #eff6ff; }
        .outlet-item.active { border-color: #2563eb; background-color: #dbeafe; }
        .outlet-name { font-weight: 600; font-size: 12px; color: #0f172a; margin-bottom: 2px; word-break: break-word; }
        .outlet-info { font-size: 11px; color: #64748b; display: flex; justify-content: space-between; gap: 6px; }
        .map-canvas-wrap { position: relative; flex: 1; height: 100%; min-width: 0; min-height: 220px; }
        #map { height: 100%; width: 100%; z-index: 1; }

        /* Map Loading Overlay Bar */
        .map-loading-overlay {
            position: absolute;
            inset: 0;
            background: rgba(241, 245, 249, 0.88);
            backdrop-filter: blur(2px);
            z-index: 2000;
            display: none;
            align-items: center;
            justify-content: center;
            padding: 16px;
        }
        .map-loading-overlay.active { display: flex; }
        .map-loading-box {
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 10px;
            padding: 16px 20px;
            width: 100%;
            max-width: 360px;
            box-shadow: 0 10px 25px rgba(15, 23, 42, 0.12);
        }
        .map-loading-title { font-size: 13px; font-weight: 700; color: #0f172a; margin-bottom: 6px; }
        .map-loading-sub { font-size: 11px; color: #64748b; margin-bottom: 10px; display: flex; justify-content: space-between; gap: 8px; }
        .progress-track { width: 100%; height: 10px; background: #e2e8f0; border-radius: 999px; overflow: hidden; }
        .progress-fill { width: 0%; height: 100%; background: linear-gradient(90deg, #2563eb, #0ea5e9); border-radius: 999px; transition: width 0.1s ease-out; }

        /* Tablet & Mobile Responsive Breakpoints */
        @media (max-width: 900px) {
            .map-layout { flex-direction: column-reverse; }
            .sidebar {
                width: 100%;
                height: 44%;
                max-height: none;
                border-right: none;
                border-top: 1px solid #cbd5e1;
                padding: 12px;
                gap: 8px;
            }
            .map-canvas-wrap {
                width: 100%;
                height: 56%;
                flex: none;
            }
        }

        @media (max-width: 640px) {
            header { padding: 8px 12px; }
            .header-right { width: 100%; justify-content: space-between; }
            .last-update-pill { font-size: 10px; padding: 3px 6px; white-space: normal; }
            .nav-buttons { width: 100%; gap: 6px; }
            .nav-btn { flex: 1 1 140px; text-align: center; padding: 7px 8px; font-size: 11px; }
            .table-content { padding: 14px 10px 24px 10px; }
            .filter-bar { padding: 10px; gap: 10px; grid-template-columns: 1fr; }
            .filter-actions { flex-direction: column; width: 100%; }
            .btn-reset, .btn-excel { width: 100%; }
            .summary-cards { grid-template-columns: 1fr 1fr; gap: 10px; }
            .scard { padding: 12px; }
            .bottom-cta { justify-content: stretch; }
            .btn-open-map { width: 100%; text-align: center; }
            .sidebar { height: 48%; padding: 10px; }
        }

        @media (max-width: 400px) {
            .summary-cards { grid-template-columns: 1fr; }
        }

        ::-webkit-scrollbar { width: 5px; height: 5px; }
        ::-webkit-scrollbar-track { background: #f1f5f9; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
    </style>
</head>
<body>
    <header>
        <h1>
            <span data-i18n="mainTitle">🗺️ Sildenafil Package Dashboard</span>
            <span class="header-logos">
                __LOGO_50_TAG__
                __LOGO_100_TAG__
            </span>
        </h1>
        <div class="header-right">
            <span class="last-update-pill">
                <span data-i18n="lastUpdatedLabel">🕒 Terakhir Diperbarui:</span> <span id="lastUpdateSpan"></span>
            </span>
            <div class="lang-switch">
                <button id="btnLangId" class="lang-btn active" onclick="switchLanguage('id')">ID</button>
                <button id="btnLangEn" class="lang-btn" onclick="switchLanguage('en')">EN</button>
            </div>
        </div>
        <div class="nav-buttons">
            <button id="btnTableView" class="nav-btn active" onclick="switchView('table')" data-i18n="btnTable">📊 Tabel Target vs Actual</button>
            <button id="btnMapView" class="nav-btn" onclick="switchView('map')" data-i18n="btnMap">📍 Visualisasi Sebaran Titik Distribusi</button>
        </div>
    </header>

    <div class="view-container">
        <!-- VIEW 1: TABEL RINGKASAN TARGET VS ACTUAL (CASCADING FILTER) -->
        <div id="viewTable" class="view-page active">
            <div class="table-content">
                <div class="filter-bar">
                    <div class="filter-field">
                        <label for="tableRegionFilter" data-i18n="filterRegion">Filter Region:</label>
                        <select id="tableRegionFilter" onchange="onTableRegionChange()">
                            <option value="ALL" data-i18n="optAllRegions">-- Semua Region --</option>
                        </select>
                    </div>
                    <div class="filter-field">
                        <label for="tableAsmFilter" data-i18n="filterAsm">Filter ASM:</label>
                        <select id="tableAsmFilter" onchange="onTableAsmChange()">
                            <option value="ALL" data-i18n="optAllAsms">-- Semua ASM --</option>
                        </select>
                    </div>
                    <div class="filter-field">
                        <label for="tablePositionFilter" data-i18n="filterRole">Filter Role:</label>
                        <select id="tablePositionFilter" onchange="onTablePositionChange()">
                            <option value="ALL" data-i18n="optAllRoles">-- Semua Role --</option>
                            <option value="SBD">SBD</option>
                            <option value="MSD">MSD</option>
                            <option value="ASM">ASM</option>
                            <option value="TSS">TSS</option>
                        </select>
                    </div>
                    <div class="filter-field">
                        <label for="tableSurveyorFilter" data-i18n="filterSurveyor">Filter Surveyor:</label>
                        <select id="tableSurveyorFilter" onchange="renderTableData()">
                            <option value="ALL" data-i18n="optAllSurveyors">-- Semua Surveyor --</option>
                        </select>
                    </div>
                    <div class="filter-actions">
                        <button onclick="resetTableFilters()" class="btn-reset" data-i18n="btnResetFilters">Atur Ulang Filter</button>
                        <button onclick="downloadFilteredExcel()" class="btn-excel" data-i18n="btnDownloadExcel">📥 Unduh Excel Sesuai Filter</button>
                    </div>
                </div>

                <div class="summary-cards">
                    <div class="scard" style="background-color: #fef2f2; border-color: #fca5a5;">
                        <h4 style="color: #991b1b;" data-i18n="cardTargetPkg">Total Target Distribusi Sildenafil Package</h4>
                        <div id="sumTargetPkg" class="sval" style="color: #b91c1c;">0</div>
                    </div>
                    <div class="scard" style="background-color: #fef2f2; border-color: #fca5a5;">
                        <h4 style="color: #991b1b;" data-i18n="cardActualPkg">Total Actual Distribusi Sildenafil Package</h4>
                        <div id="sumActualPkg" class="sval" style="color: #b91c1c;">0 <small id="sumPkgPct" style="font-size:12px; font-weight:normal; color:#b91c1c;">(0.0% vs Target)</small></div>
                    </div>
                    <div class="scard" style="background-color: #eff6ff; border-color: #93c5fd;">
                        <h4 style="color: #1e40af;" data-i18n="cardTargetCust">Total Target Customer</h4>
                        <div id="sumTargetCust" class="sval" style="color: #1d4ed8;">0</div>
                    </div>
                    <div class="scard" style="background-color: #eff6ff; border-color: #93c5fd;">
                        <h4 style="color: #1e40af;" data-i18n="cardActualCust">Total Actual Customer</h4>
                        <div id="sumActualCust" class="sval" style="color: #1d4ed8;">0 <small id="sumCustPct" style="font-size:12px; font-weight:normal; color:#1e40af;">(0.0% vs Target)</small></div>
                    </div>
                </div>

                <!-- TABEL SUMMARY DENGAN PILIHAN TOGGLE (REGION / ASM) -->
                <div class="summary-toggle-bar">
                    <div id="summaryTableHeaderTitle" class="summary-title" data-i18n="regionSummaryTitle">📋 Ringkasan Performa Per Region (RSM)</div>
                    <div class="toggle-group">
                        <button id="toggleRegionBtn" class="toggle-btn active" onclick="setSummaryMode('region')" data-i18n="toggleRegion">Region (RSM)</button>
                        <button id="toggleAsmBtn" class="toggle-btn" onclick="setSummaryMode('asm')" data-i18n="toggleAsm">ASM</button>
                    </div>
                </div>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr id="summaryTableHeadRow">
                                <!-- Dinamis via JS tergantung mode region / asm -->
                            </tr>
                        </thead>
                        <tbody id="summaryTableHeaderBody">
                            <!-- Diisi via JavaScript -->
                        </tbody>
                    </table>
                </div>

                <!-- TABEL UTAMA DETAIL SURVEYOR -->
                <div style="margin-bottom: 8px; font-weight: 600; font-size: 14px; color: #0f172a;" data-i18n="detailTableTitle">📊 Detail Performa Per Surveyor</div>
                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th data-i18n="thRegion">Region</th>
                                <th data-i18n="thAsm">ASM</th>
                                <th data-i18n="thRole">Role</th>
                                <th data-i18n="thSurveyorName">Nama Surveyor</th>
                                <th style="text-align: center;"><span data-i18n="thDistPkg">Distribusi Sildenafil Package</span><br><small style="font-weight: normal; color: #64748b;" data-i18n="thActualVsTarget">(Aktual vs Target)</small></th>
                                <th style="text-align: center;"><span data-i18n="thCustPkg">Customer Sildenafil Package</span><br><small style="font-weight: normal; color: #64748b;" data-i18n="thActualVsTarget">(Aktual vs Target)</small></th>
                            </tr>
                        </thead>
                        <tbody id="summaryTableBody">
                            <!-- Diisi via JavaScript -->
                        </tbody>
                    </table>
                </div>
                <div class="bottom-cta">
                    <button onclick="switchView('map')" class="btn-open-map" data-i18n="btnOpenMap">Buka Visualisasi Sebaran Titik Distribusi ➡️</button>
                </div>
            </div>
        </div>

        <!-- VIEW 2: PETA GIS & SEBARAN -->
        <div id="viewMap" class="view-page">
            <div class="map-layout">
                <div class="sidebar">
                    <div class="sidebar-filter">
                        <label for="mapRegionFilter" data-i18n="filterRegion">Filter Region:</label>
                        <select id="mapRegionFilter" onchange="onMapRegionChange()">
                            <option value="ALL" data-i18n="optAllRegions">-- Semua Region --</option>
                        </select>
                    </div>
                    <div class="sidebar-filter">
                        <label for="mapAsmFilter" data-i18n="filterAsm">Filter ASM:</label>
                        <select id="mapAsmFilter" onchange="onMapAsmChange()">
                            <option value="ALL" data-i18n="optAllAsms">-- Semua ASM --</option>
                        </select>
                    </div>
                    <div class="sidebar-filter">
                        <label for="mapSurveyorFilter" data-i18n="filterSurveyor">Filter Surveyor:</label>
                        <select id="mapSurveyorFilter" onchange="onMapSurveyorChange()">
                            <option value="ALL" data-i18n="optAllSurveyors">-- Semua Surveyor --</option>
                        </select>
                    </div>
                    <div class="sidebar-filter">
                        <label for="mapOutletFilter" data-i18n="filterOutlet">Filter Outlet (Surveyee):</label>
                        <select id="mapOutletFilter" onchange="onMapOutletChange()">
                            <option value="ALL" data-i18n="optAllOutlets">-- Semua Outlet --</option>
                        </select>
                    </div>
                    <div>
                        <button onclick="resetMapFilters()" style="width: 100%; padding: 7px 10px; background: #cbd5e1; color: #0f172a; border: none; border-radius: 6px; font-size: 11px; font-weight: 600; cursor: pointer;" data-i18n="btnResetMapFilters">Atur Ulang Filter Peta</button>
                    </div>
                    <div style="background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%); padding: 10px; border-radius: 8px; color: white;">
                        <h3 style="margin: 0 0 2px 0; font-size: 10px; text-transform: uppercase;" data-i18n="mapTotalStatTitle">Total Titik Sebaran Peta</h3>
                        <div id="mapTotalStat" class="val" style="font-size: 18px; font-weight: bold;">0</div>
                    </div>
                    <div style="font-size: 11px; font-weight: 600; color: #64748b;" data-i18n="mapOutletListTitle">Daftar Outlet Terfilter:</div>
                    <div id="outletListContainer" class="outlet-list"></div>
                </div>
                <div class="map-canvas-wrap">
                    <div id="mapLoadingOverlay" class="map-loading-overlay">
                        <div class="map-loading-box">
                            <div id="mapLoadingTitle" class="map-loading-title" data-i18n="loadingPointsTitle">Memproses titik sebaran distribusi...</div>
                            <div class="map-loading-sub">
                                <span id="mapLoadingDetail">0 / 0 titik</span>
                                <span id="mapLoadingPct">0%</span>
                            </div>
                            <div class="progress-track">
                                <div id="mapLoadingBar" class="progress-fill"></div>
                            </div>
                        </div>
                    </div>
                    <div id="map"></div>
                </div>
            </div>
        </div>
    </div>

    <script>
        const rawPayload = __PAYLOAD_JSON_PLACEHOLDER__;
        const surveyData = rawPayload.surveyData || [];
        const masterSurveyor = rawPayload.masterSurveyor || [];

        // Pra-indeks surveyData per surveyor_id agar render tabel tidak melakukan filter berulang
        const surveyBySurveyorId = {};
        surveyData.forEach(item => {
            const sid = item.surveyor_id;
            if (!surveyBySurveyorId[sid]) surveyBySurveyorId[sid] = [];
            surveyBySurveyorId[sid].push(item);
        });

        const masterStatusById = {};
        const masterStatusByName = {};
        masterSurveyor.forEach(m => {
            const st = String(m.status || 'Active').trim().toLowerCase() === 'inactive' ? 'Inactive' : 'Active';
            if (m.surveyor_id) masterStatusById[String(m.surveyor_id).trim()] = st;
            if (m.surveyor) masterStatusByName[String(m.surveyor).trim()] = st;
        });

        function resolveStatus(statusVal, surveyorId, surveyorName) {
            if (statusVal && String(statusVal).trim() !== '') {
                return String(statusVal).trim().toLowerCase() === 'inactive' ? 'Inactive' : 'Active';
            }
            const sid = surveyorId ? String(surveyorId).trim() : '';
            if (sid && masterStatusById[sid]) return masterStatusById[sid];
            const sname = surveyorName ? String(surveyorName).trim() : '';
            if (sname && masterStatusByName[sname]) return masterStatusByName[sname];
            return 'Active';
        }

        function formatSurveyorTooltip(role, status) {
            const r = role || '-';
            if (status === 'Inactive') {
                return `${r} status Inactive - Vacant`;
            }
            return `${r} status Active`;
        }

        function escapeAttr(str) {
            return String(str || '')
                .replace(/&/g, '&amp;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;');
        }

        // Nama surveyor tetap normal; hanya bagian (Active) / (Inactive) yang jadi label badge
        function renderSurveyorLabelHtml(name, role, status) {
            const safeName = escapeAttr(name || 'Unknown');
            const tooltipText = escapeAttr(formatSurveyorTooltip(role, status));
            const badgeClass = status === 'Inactive' ? 'status-badge-inactive' : 'status-badge-active';
            return `<b>${safeName}</b><span class="${badgeClass}" title="${tooltipText}">(${status})</span>`;
        }

        // --- Multi-Language Dictionary ---
        const translations = {
            id: {
                mainTitle: "🗺️ Sildenafil Package Dashboard",
                lastUpdatedLabel: "🕒 Terakhir Diperbarui:",
                btnTable: "📊 Tabel Target vs Actual",
                btnMap: "📍 Visualisasi Sebaran Titik Distribusi",
                filterRegion: "Filter Region:",
                filterAsm: "Filter ASM:",
                filterRole: "Filter Role:",
                filterSurveyor: "Filter Surveyor:",
                filterOutlet: "Filter Outlet (Surveyee):",
                optAllRegions: "-- Semua Region --",
                optAllAsms: "-- Semua ASM --",
                optAllRoles: "-- Semua Role --",
                optAllSurveyors: "-- Semua Surveyor --",
                optAllOutlets: "-- Semua Outlet --",
                btnResetFilters: "Atur Ulang Filter",
                btnDownloadExcel: "📥 Unduh Excel Sesuai Filter",
                cardTargetPkg: "Total Target Distribusi Sildenafil Package",
                cardActualPkg: "Total Actual Distribusi Sildenafil Package",
                cardTargetCust: "Total Target Customer",
                cardActualCust: "Total Actual Customer",
                thRegion: "Region",
                thAsm: "ASM",
                thRole: "Role",
                thSurveyorName: "Nama Surveyor",
                thDistPkg: "Distribusi Sildenafil Package",
                thCustPkg: "Customer Sildenafil Package",
                thActualVsTarget: "(Aktual vs Target)",
                btnOpenMap: "Buka Visualisasi Sebaran Titik Distribusi ➡️",
                btnResetMapFilters: "Atur Ulang Filter Peta",
                mapTotalStatTitle: "Total Titik Sebaran Peta",
                mapOutletListTitle: "Daftar Outlet Terfilter:",
                mapPopupMapsBtn: "🗺️ Buka Rute ke Google Maps",
                mapPopupNoPhoto: "Tidak ada foto.",
                mapPopupPhotoLabel: "Foto",
                regionSummaryTitle: "📋 Ringkasan Performa Per Region (RSM)",
                asmSummaryTitle: "📋 Ringkasan Performa Per ASM",
                detailTableTitle: "📊 Detail Performa Per Surveyor",
                toggleRegion: "Region (RSM)",
                toggleAsm: "ASM",
                totalNational: "TOTAL NASIONAL",
                loadingPointsTitle: "Memproses seluruh titik sebaran distribusi..."
            },
            en: {
                mainTitle: "🗺️ Sildenafil Package Dashboard",
                lastUpdatedLabel: "🕒 Last Updated:",
                btnTable: "📊 Target vs Actual Table",
                btnMap: "📍 Distribution Map Visualization",
                filterRegion: "Filter Region:",
                filterAsm: "Filter ASM:",
                filterRole: "Filter Role:",
                filterSurveyor: "Filter Surveyor:",
                filterOutlet: "Filter Outlet (Surveyee):",
                optAllRegions: "-- All Regions --",
                optAllAsms: "-- All ASMs --",
                optAllRoles: "-- All Roles --",
                optAllSurveyors: "-- All Surveyors --",
                optAllOutlets: "-- All Outlets --",
                btnResetFilters: "Reset Filters",
                btnDownloadExcel: "📥 Download Filtered Excel",
                cardTargetPkg: "Total Sildenafil Package Distribution Target",
                cardActualPkg: "Total Actual Sildenafil Package Distribution",
                cardTargetCust: "Total Customer Target",
                cardActualCust: "Total Actual Customer",
                thRegion: "Region",
                thAsm: "ASM",
                thRole: "Role",
                thSurveyorName: "Surveyor Name",
                thDistPkg: "Sildenafil Package Distribution",
                thCustPkg: "Sildenafil Package Customers",
                thActualVsTarget: "(Actual vs Target)",
                btnOpenMap: "Open Distribution Map Visualization ➡️",
                btnResetMapFilters: "Reset Map Filters",
                mapTotalStatTitle: "Total Map Distribution Points",
                mapOutletListTitle: "Filtered Outlet List:",
                mapPopupMapsBtn: "🗺️ Open Route in Google Maps",
                mapPopupNoPhoto: "No photos available.",
                mapPopupPhotoLabel: "Photo",
                regionSummaryTitle: "📋 Performance Summary Per Region (RSM)",
                asmSummaryTitle: "📋 Performance Summary Per ASM",
                detailTableTitle: "📊 Detailed Performance Per Surveyor",
                toggleRegion: "Region (RSM)",
                toggleAsm: "ASM",
                totalNational: "NATIONAL TOTAL",
                loadingPointsTitle: "Processing all distribution points..."
            }
        };

        let currentLang = 'id';
        let currentSummaryMode = 'region'; // 'region' or 'asm'

        function syncHeaderOffset() {
            const hdr = document.querySelector('header');
            const container = document.querySelector('.view-container');
            if (hdr && container) {
                const h = Math.ceil(hdr.getBoundingClientRect().height) + 8;
                container.style.setProperty('--header-offset', `${h}px`);
            }
            if (map && document.getElementById('viewMap').classList.contains('active')) {
                setTimeout(() => { if (map) map.invalidateSize(); }, 80);
            }
        }

        function setSummaryMode(mode) {
            currentSummaryMode = mode;
            document.getElementById('toggleRegionBtn').classList.toggle('active', mode === 'region');
            document.getElementById('toggleAsmBtn').classList.toggle('active', mode === 'asm');
            
            const titleEl = document.getElementById('summaryTableHeaderTitle');
            if (mode === 'region') {
                titleEl.setAttribute('data-i18n', 'regionSummaryTitle');
                titleEl.textContent = translations[currentLang].regionSummaryTitle;
            } else {
                titleEl.setAttribute('data-i18n', 'asmSummaryTitle');
                titleEl.textContent = translations[currentLang].asmSummaryTitle;
            }
            renderTableData();
        }

        function switchLanguage(lang) {
            currentLang = lang;
            document.getElementById('btnLangId').classList.toggle('active', lang === 'id');
            document.getElementById('btnLangEn').classList.toggle('active', lang === 'en');
            
            document.querySelectorAll('[data-i18n]').forEach(el => {
                const key = el.getAttribute('data-i18n');
                if (translations[currentLang][key]) {
                    el.textContent = translations[currentLang][key];
                }
            });

            populateTableRegionDropdown();
            populateTableAsmDropdown(document.getElementById('tableRegionFilter').value);
            populateTableSurveyorDropdown(
                document.getElementById('tableRegionFilter').value,
                document.getElementById('tableAsmFilter').value,
                document.getElementById('tablePositionFilter').value
            );

            renderTableData();

            if (map && document.getElementById('viewMap').classList.contains('active')) {
                populateMapAsmDropdown(surveyData);
                populateMapSurveyorDropdown(surveyData);
                populateMapOutletDropdown(surveyData);
                renderMapDashboard(
                    document.getElementById('mapRegionFilter').value,
                    document.getElementById('mapAsmFilter').value,
                    document.getElementById('mapSurveyorFilter').value,
                    document.getElementById('mapOutletFilter').value
                );
            }
            syncHeaderOffset();
        }

        const targetPerPerson = {
            "SBD": { pkg: 600, cust: 300 },
            "MSD": { pkg: 200, cust: 100 },
            "ASM": { pkg: 40, cust: 4 },
            "TSS": { pkg: 20, cust: 2 }
        };

        const fmtNum = (num) => Number(num || 0).toLocaleString('en-US');

        let surveyorColorMap = {};
        let map, markersLayer;
        let mapRenderToken = 0;

        function switchView(viewName) {
            // Batalkan proses chunked peta yang sedang berjalan agar pindah tab instan
            mapRenderToken++;
            hideMapLoadingOverlay();

            document.getElementById('viewTable').classList.remove('active');
            document.getElementById('viewMap').classList.remove('active');
            document.getElementById('btnTableView').classList.remove('active');
            document.getElementById('btnMapView').classList.remove('active');

            if (viewName === 'table') {
                document.getElementById('viewTable').classList.add('active');
                document.getElementById('btnTableView').classList.add('active');
            } else {
                document.getElementById('viewMap').classList.add('active');
                document.getElementById('btnMapView').classList.add('active');
                if (!map) {
                    initMap();
                } else {
                    map.invalidateSize();
                }
                renderMapDashboard(
                    document.getElementById('mapRegionFilter').value,
                    document.getElementById('mapAsmFilter').value,
                    document.getElementById('mapSurveyorFilter').value,
                    document.getElementById('mapOutletFilter').value
                );
            }
            syncHeaderOffset();
        }

        // --- CASCADING TABLE FILTERS ---
        function getTableFilteredMasters(reg, asm, pos, srv) {
            return masterSurveyor.filter(m => {
                if (reg && reg !== 'ALL' && m.region !== reg) return false;
                if (asm && asm !== 'ALL' && m.asm !== asm) return false;
                if (pos && pos !== 'ALL' && m.position.toUpperCase() !== pos.toUpperCase()) return false;
                if (srv && srv !== 'ALL' && m.surveyor !== srv) return false;
                return true;
            });
        }

        function initTableFilters() {
            populateTableRegionDropdown();
            populateTableAsmDropdown('ALL');
            populateTableSurveyorDropdown('ALL', 'ALL', 'ALL');
            renderTableData();
        }

        function populateTableRegionDropdown() {
            const regSelect = document.getElementById('tableRegionFilter');
            const currentVal = regSelect.value;
            const defaultText = translations[currentLang].optAllRegions;
            regSelect.innerHTML = `<option value="ALL">${defaultText}</option>`;
            const regions = [...new Set(masterSurveyor.map(m => m.region))].filter(r => r && r !== '-').sort();
            regions.forEach(r => {
                const opt = document.createElement('option');
                opt.value = r; opt.textContent = r;
                regSelect.appendChild(opt);
            });
            regSelect.value = regions.includes(currentVal) ? currentVal : 'ALL';
        }

        function populateTableAsmDropdown(region) {
            const asmSelect = document.getElementById('tableAsmFilter');
            const currentVal = asmSelect.value;
            const defaultText = translations[currentLang].optAllAsms;
            asmSelect.innerHTML = `<option value="ALL">${defaultText}</option>`;
            
            let subset = masterSurveyor;
            if (region && region !== 'ALL') {
                subset = subset.filter(m => m.region === region);
            }
            const asms = [...new Set(subset.map(m => m.asm))].filter(a => a && a !== '-').sort();
            asms.forEach(a => {
                const opt = document.createElement('option');
                opt.value = a; opt.textContent = a;
                asmSelect.appendChild(opt);
            });
            asmSelect.value = asms.includes(currentVal) ? currentVal : 'ALL';
        }

        function populateTableSurveyorDropdown(region, asm, position) {
            const srvSelect = document.getElementById('tableSurveyorFilter');
            const currentVal = srvSelect.value;
            const defaultText = translations[currentLang].optAllSurveyors;
            srvSelect.innerHTML = `<option value="ALL">${defaultText}</option>`;

            let subset = masterSurveyor;
            if (region && region !== 'ALL') subset = subset.filter(m => m.region === region);
            if (asm && asm !== 'ALL') subset = subset.filter(m => m.asm === asm);
            if (position && position !== 'ALL') subset = subset.filter(m => m.position.toUpperCase() === position.toUpperCase());

            const surveyors = [...new Set(subset.map(m => m.surveyor))].filter(s => s && s !== '-').sort();
            surveyors.forEach(s => {
                const st = resolveStatus(null, null, s);
                const opt = document.createElement('option');
                opt.value = s;
                opt.textContent = `${s} (${st})`;
                srvSelect.appendChild(opt);
            });
            srvSelect.value = surveyors.includes(currentVal) ? currentVal : 'ALL';
        }

        function onTableRegionChange() {
            const reg = document.getElementById('tableRegionFilter').value;
            const asm = document.getElementById('tableAsmFilter').value;
            const pos = document.getElementById('tablePositionFilter').value;

            populateTableAsmDropdown(reg);
            populateTableSurveyorDropdown(reg, asm, pos);
            renderTableData();
        }

        function onTableAsmChange() {
            const reg = document.getElementById('tableRegionFilter').value;
            const asm = document.getElementById('tableAsmFilter').value;
            const pos = document.getElementById('tablePositionFilter').value;

            populateTableSurveyorDropdown(reg, asm, pos);
            renderTableData();
        }

        function onTablePositionChange() {
            const reg = document.getElementById('tableRegionFilter').value;
            const asm = document.getElementById('tableAsmFilter').value;
            const pos = document.getElementById('tablePositionFilter').value;

            populateTableSurveyorDropdown(reg, asm, pos);
            renderTableData();
        }

        function resetTableFilters() {
            document.getElementById('tableRegionFilter').value = 'ALL';
            document.getElementById('tableAsmFilter').value = 'ALL';
            document.getElementById('tablePositionFilter').value = 'ALL';
            document.getElementById('tableSurveyorFilter').value = 'ALL';
            populateTableAsmDropdown('ALL');
            populateTableSurveyorDropdown('ALL', 'ALL', 'ALL');
            renderTableData();
        }

        function renderTableData() {
            const regVal = document.getElementById('tableRegionFilter').value;
            const asmVal = document.getElementById('tableAsmFilter').value;
            const posVal = document.getElementById('tablePositionFilter').value;
            const srvVal = document.getElementById('tableSurveyorFilter').value;

            let filteredMasters = getTableFilteredMasters(regVal, asmVal, posVal, srvVal);

            const roleRank = { 'ASM': 1, 'TSS': 2, 'SBD': 3, 'MSD': 4 };
            filteredMasters.sort((a, b) => {
                let regA = a.region || '';
                let regB = b.region || '';
                if (regA !== regB) return regA.localeCompare(regB);

                let asmA = a.asm || '';
                let asmB = b.asm || '';
                if (asmA !== asmB) return asmA.localeCompare(asmB);

                let rA = roleRank[a.position.toUpperCase().trim()] || 99;
                let rB = roleRank[b.position.toUpperCase().trim()] || 99;
                if (rA !== rB) return rA - rB;

                return (a.surveyor || '').localeCompare(b.surveyor || '');
            });

            let actualPkgMap = {};
            let actualCustMap = {};

            surveyData.forEach(item => {
                let sid = item.surveyor_id;
                let q = typeof item.qty === 'number' ? item.qty : parseFloat(item.qty) || 0;
                actualPkgMap[sid] = (actualPkgMap[sid] || 0) + q;

                if (!actualCustMap[sid]) actualCustMap[sid] = new Set();
                if (item.surveyee_id) actualCustMap[sid].add(item.surveyee_id);
            });

            const headRow = document.getElementById('summaryTableHeadRow');
            const summaryTbody = document.getElementById('summaryTableHeaderBody');
            summaryTbody.innerHTML = '';

            const thRegText = translations[currentLang].thRegion;
            const thAsmText = translations[currentLang].thAsm;
            const thDistText = translations[currentLang].thDistPkg;
            const thCustText = translations[currentLang].thCustPkg;
            const thActualTargetText = translations[currentLang].thActualVsTarget;
            const totalNatText = translations[currentLang].totalNational;

            const isRegionFiltered = (regVal && regVal !== 'ALL');
            const isAsmFiltered = (asmVal && asmVal !== 'ALL');

            if (currentSummaryMode === 'region') {
                headRow.innerHTML = `
                    <th>${thRegText}</th>
                    <th style="text-align: center;"><span>${thDistText}</span><br><small style="font-weight: normal; color: #64748b;">${thActualTargetText}</small></th>
                    <th style="text-align: center;"><span>${thCustText}</span><br><small style="font-weight: normal; color: #64748b;">${thActualTargetText}</small></th>
                `;

                const regionGroupMap = {};
                filteredMasters.forEach(m => {
                    let regName = m.region || '-';
                    if (!regionGroupMap[regName]) {
                        regionGroupMap[regName] = { region: regName, tgtPkg: 0, actPkg: 0, tgtCust: 0, actCustSet: new Set() };
                    }
                    let pos = m.position.toUpperCase().trim();
                    let tPkg = (targetPerPerson[pos] ? targetPerPerson[pos].pkg : 0);
                    let tCust = (targetPerPerson[pos] ? targetPerPerson[pos].cust : 0);
                    let aPkg = actualPkgMap[m.surveyor_id] || 0;

                    regionGroupMap[regName].tgtPkg += tPkg;
                    regionGroupMap[regName].actPkg += aPkg;
                    regionGroupMap[regName].tgtCust += tCust;

                    const rowsForSid = surveyBySurveyorId[m.surveyor_id] || [];
                    rowsForSid.forEach(d => {
                        if (d.surveyee_id) regionGroupMap[regName].actCustSet.add(d.surveyee_id);
                    });
                });

                let regionList = Object.values(regionGroupMap);

                regionList.sort((a, b) => {
                    let pctA = a.tgtPkg > 0 ? (a.actPkg / a.tgtPkg) : 0;
                    let pctB = b.tgtPkg > 0 ? (b.actPkg / b.tgtPkg) : 0;
                    return pctB - pctA;
                });

                if (!isRegionFiltered && !isAsmFiltered) {
                    let natTgtPkg = 0, natActPkg = 0, natTgtCust = 0;
                    let natActCustSet = new Set();
                    regionList.forEach(g => {
                        natTgtPkg += g.tgtPkg;
                        natActPkg += g.actPkg;
                        natTgtCust += g.tgtCust;
                        g.actCustSet.forEach(cid => natActCustSet.add(cid));
                    });

                    let natPctPkg = natTgtPkg > 0 ? ((natActPkg / natTgtPkg) * 100).toFixed(1) : "0.0";
                    let natActCustCount = natActCustSet.size;
                    let natPctCust = natTgtCust > 0 ? ((natActCustCount / natTgtCust) * 100).toFixed(1) : "0.0";

                    const trTotal = document.createElement('tr');
                    trTotal.className = 'total-row';
                    trTotal.innerHTML = `
                        <td>⭐ <b>${totalNatText}</b></td>
                        <td style="text-align: center;">
                            <b>${fmtNum(natActPkg)}</b> / ${fmtNum(natTgtPkg)} <br>
                            <small style="color: ${natActPkg >= natTgtPkg ? '#059669' : '#dc2626'}; font-weight: 600;">(${natPctPkg}%)</small>
                        </td>
                        <td style="text-align: center;">
                            <b>${fmtNum(natActCustCount)}</b> / ${fmtNum(natTgtCust)} <br>
                            <small style="color: ${natActCustCount >= natTgtCust ? '#059669' : '#dc2626'}; font-weight: 600;">(${natPctCust}%)</small>
                        </td>
                    `;
                    summaryTbody.appendChild(trTotal);
                }

                const regFrag = document.createDocumentFragment();
                regionList.forEach(group => {
                    let aCust = group.actCustSet.size;
                    let pPkg = group.tgtPkg > 0 ? ((group.actPkg / group.tgtPkg) * 100).toFixed(1) : "0.0";
                    let pCust = group.tgtCust > 0 ? ((aCust / group.tgtCust) * 100).toFixed(1) : "0.0";

                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td><b>${group.region}</b></td>
                        <td style="text-align: center;">
                            <b>${fmtNum(group.actPkg)}</b> / ${fmtNum(group.tgtPkg)} <br>
                            <small style="color: ${group.actPkg >= group.tgtPkg ? '#059669' : '#dc2626'}; font-weight: 600;">(${pPkg}%)</small>
                        </td>
                        <td style="text-align: center;">
                            <b>${fmtNum(aCust)}</b> / ${fmtNum(group.tgtCust)} <br>
                            <small style="color: ${aCust >= group.tgtCust ? '#059669' : '#dc2626'}; font-weight: 600;">(${pCust}%)</small>
                        </td>
                    `;
                    regFrag.appendChild(tr);
                });
                summaryTbody.appendChild(regFrag);

                if (regionList.length === 0) {
                    summaryTbody.innerHTML = `<tr><td colspan="3" style="text-align: center; color: #64748b; font-style: italic;">Tidak ada data Region yang sesuai filter.</td></tr>`;
                }

            } else {
                headRow.innerHTML = `
                    <th>${thRegText}</th>
                    <th>${thAsmText}</th>
                    <th style="text-align: center;"><span>${thDistText}</span><br><small style="font-weight: normal; color: #64748b;">${thActualTargetText}</small></th>
                    <th style="text-align: center;"><span>${thCustText}</span><br><small style="font-weight: normal; color: #64748b;">${thActualTargetText}</small></th>
                `;

                const asmGroupMap = {};
                filteredMasters.forEach(m => {
                    let asmName = m.asm || '-';
                    let regName = m.region || '-';
                    let asmKey = `${regName}___${asmName}`;
                    
                    if (!asmGroupMap[asmKey]) {
                        asmGroupMap[asmKey] = { region: regName, asm: asmName, tgtPkg: 0, actPkg: 0, tgtCust: 0, actCustSet: new Set() };
                    }
                    let pos = m.position.toUpperCase().trim();
                    let tPkg = (targetPerPerson[pos] ? targetPerPerson[pos].pkg : 0);
                    let tCust = (targetPerPerson[pos] ? targetPerPerson[pos].cust : 0);
                    let aPkg = actualPkgMap[m.surveyor_id] || 0;

                    asmGroupMap[asmKey].tgtPkg += tPkg;
                    asmGroupMap[asmKey].actPkg += aPkg;
                    asmGroupMap[asmKey].tgtCust += tCust;

                    const rowsForSid = surveyBySurveyorId[m.surveyor_id] || [];
                    rowsForSid.forEach(d => {
                        if (d.surveyee_id) asmGroupMap[asmKey].actCustSet.add(d.surveyee_id);
                    });
                });

                let asmList = Object.values(asmGroupMap);

                asmList.sort((a, b) => {
                    let pctA = a.tgtPkg > 0 ? (a.actPkg / a.tgtPkg) : 0;
                    let pctB = b.tgtPkg > 0 ? (b.actPkg / b.tgtPkg) : 0;
                    return pctB - pctA;
                });

                if (!isRegionFiltered && !isAsmFiltered) {
                    let natTgtPkgAsm = 0, natActPkgAsm = 0, natTgtCustAsm = 0;
                    let natActCustSetAsm = new Set();
                    asmList.forEach(g => {
                        natTgtPkgAsm += g.tgtPkg;
                        natActPkgAsm += g.actPkg;
                        natTgtCustAsm += g.tgtCust;
                        g.actCustSet.forEach(cid => natActCustSetAsm.add(cid));
                    });

                    let natPctPkgAsm = natTgtPkgAsm > 0 ? ((natActPkgAsm / natTgtPkgAsm) * 100).toFixed(1) : "0.0";
                    let natActCustCountAsm = natActCustSetAsm.size;
                    let natPctCustAsm = natTgtCustAsm > 0 ? ((natActCustCountAsm / natTgtCustAsm) * 100).toFixed(1) : "0.0";

                    const trTotalAsm = document.createElement('tr');
                    trTotalAsm.className = 'total-row';
                    trTotalAsm.innerHTML = `
                        <td>⭐ <b>${totalNatText}</b></td>
                        <td>-</td>
                        <td style="text-align: center;">
                            <b>${fmtNum(natActPkgAsm)}</b> / ${fmtNum(natTgtPkgAsm)} <br>
                            <small style="color: ${natActPkgAsm >= natTgtPkgAsm ? '#059669' : '#dc2626'}; font-weight: 600;">(${natPctPkgAsm}%)</small>
                        </td>
                        <td style="text-align: center;">
                            <b>${fmtNum(natActCustCountAsm)}</b> / ${fmtNum(natTgtCustAsm)} <br>
                            <small style="color: ${natActCustCountAsm >= natTgtCustAsm ? '#059669' : '#dc2626'}; font-weight: 600;">(${natPctCustAsm}%)</small>
                        </td>
                    `;
                    summaryTbody.appendChild(trTotalAsm);
                }

                const asmFrag = document.createDocumentFragment();
                asmList.forEach(group => {
                    let aCust = group.actCustSet.size;
                    let pPkg = group.tgtPkg > 0 ? ((group.actPkg / group.tgtPkg) * 100).toFixed(1) : "0.0";
                    let pCust = group.tgtCust > 0 ? ((aCust / group.tgtCust) * 100).toFixed(1) : "0.0";

                    const tr = document.createElement('tr');
                    tr.innerHTML = `
                        <td><b>${group.region}</b></td>
                        <td><b>${group.asm}</b></td>
                        <td style="text-align: center;">
                            <b>${fmtNum(group.actPkg)}</b> / ${fmtNum(group.tgtPkg)} <br>
                            <small style="color: ${group.actPkg >= group.tgtPkg ? '#059669' : '#dc2626'}; font-weight: 600;">(${pPkg}%)</small>
                        </td>
                        <td style="text-align: center;">
                            <b>${fmtNum(aCust)}</b> / ${fmtNum(group.tgtCust)} <br>
                            <small style="color: ${aCust >= group.tgtCust ? '#059669' : '#dc2626'}; font-weight: 600;">(${pCust}%)</small>
                        </td>
                    `;
                    asmFrag.appendChild(tr);
                });
                summaryTbody.appendChild(asmFrag);

                if (asmList.length === 0) {
                    summaryTbody.innerHTML = `<tr><td colspan="4" style="text-align: center; color: #64748b; font-style: italic;">Tidak ada data ASM yang sesuai filter.</td></tr>`;
                }
            }

            const tbody = document.getElementById('summaryTableBody');
            tbody.innerHTML = '';

            let totalTgtPkg = 0;
            let totalActPkg = 0;
            let totalTgtCust = 0;
            let totalActCustSet = new Set();
            const detailFrag = document.createDocumentFragment();

            filteredMasters.forEach(m => {
                let pos = m.position.toUpperCase().trim();
                let st = resolveStatus(m.status, m.surveyor_id, m.surveyor);
                let tPkg = (targetPerPerson[pos] ? targetPerPerson[pos].pkg : 0);
                let tCust = (targetPerPerson[pos] ? targetPerPerson[pos].cust : 0);

                let aPkg = actualPkgMap[m.surveyor_id] || 0;
                let aCustSet = actualCustMap[m.surveyor_id] || new Set();
                let aCust = aCustSet.size;

                totalTgtPkg += tPkg;
                totalActPkg += aPkg;
                totalTgtCust += tCust;
                
                aCustSet.forEach(cid => totalActCustSet.add(cid));

                let pPkg = tPkg > 0 ? ((aPkg / tPkg) * 100).toFixed(1) : "0.0";
                let pCust = tCust > 0 ? ((aCust / tCust) * 100).toFixed(1) : "0.0";

                let badgeClass = 'badge-sbd';
                if (pos === 'MSD') badgeClass = 'badge-msd';
                else if (pos === 'ASM') badgeClass = 'badge-asm';
                else if (pos === 'TSS') badgeClass = 'badge-tss';

                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${m.region}</td>
                    <td>${m.asm}</td>
                    <td><span class="badge ${badgeClass}">${m.position}</span></td>
                    <td>${renderSurveyorLabelHtml(m.surveyor, m.position, st)}</td>
                    <td style="text-align: center;">
                        <b>${fmtNum(aPkg)}</b> / ${fmtNum(tPkg)} <br>
                        <small style="color: ${aPkg >= tPkg ? '#059669' : '#dc2626'}; font-weight: 600;">(${pPkg}%)</small>
                    </td>
                    <td style="text-align: center;">
                        <b>${fmtNum(aCust)}</b> / ${fmtNum(tCust)} <br>
                        <small style="color: ${aCust >= tCust ? '#059669' : '#dc2626'}; font-weight: 600;">(${pCust}%)</small>
                    </td>
                `;
                detailFrag.appendChild(tr);
            });
            tbody.appendChild(detailFrag);

            let pctPkgOverall = totalTgtPkg > 0 ? ((totalActPkg / totalTgtPkg) * 100).toFixed(1) : "0.0";
            let actualCustCount = totalActCustSet.size;
            let pctCustOverall = totalTgtCust > 0 ? ((actualCustCount / totalTgtCust) * 100).toFixed(1) : "0.0";

            document.getElementById('sumTargetPkg').innerText = fmtNum(totalTgtPkg);
            document.getElementById('sumActualPkg').innerHTML = `${fmtNum(totalActPkg)} <small id="sumPkgPct" style="font-size:12px; font-weight:normal; color:#b91c1c;">(${pctPkgOverall}% vs Target)</small>`;
            document.getElementById('sumTargetCust').innerText = fmtNum(totalTgtCust);
            document.getElementById('sumActualCust').innerHTML = `${fmtNum(actualCustCount)} <small id="sumCustPct" style="font-size:12px; font-weight:normal; color:#1e40af;">(${pctCustOverall}% vs Target)</small>`;
        }

        // --- MAP VIEW ---
        function initMapFilters() {
            const regSelect = document.getElementById('mapRegionFilter');
            const regions = [...new Set(surveyData.map(item => item.region))].sort();
            const defaultRegText = translations[currentLang].optAllRegions;
            regSelect.innerHTML = `<option value="ALL">${defaultRegText}</option>`;
            regions.forEach(r => {
                if(r && r !== '-') {
                    const opt = document.createElement('option');
                    opt.value = r; opt.textContent = r;
                    regSelect.appendChild(opt);
                }
            });
            populateMapAsmDropdown(surveyData);
            populateMapSurveyorDropdown(surveyData);
            populateMapOutletDropdown(surveyData);
        }

        function populateMapAsmDropdown(dataList) {
            const select = document.getElementById('mapAsmFilter');
            const currentVal = select.value;
            const defaultText = translations[currentLang].optAllAsms;
            select.innerHTML = `<option value="ALL">${defaultText}</option>`;
            const asms = [...new Set(dataList.map(item => item.asm))].sort();
            asms.forEach(a => {
                if(a && a !== '-') {
                    const opt = document.createElement('option');
                    opt.value = a; opt.textContent = a;
                    select.appendChild(opt);
                }
            });
            select.value = asms.includes(currentVal) ? currentVal : 'ALL';
        }

        function populateMapSurveyorDropdown(dataList) {
            const select = document.getElementById('mapSurveyorFilter');
            const currentVal = select.value;
            const defaultText = translations[currentLang].optAllSurveyors;
            select.innerHTML = `<option value="ALL">${defaultText}</option>`;
            const surveyors = [...new Set(dataList.map(item => item.surveyor))].sort();
            surveyors.forEach(s => {
                if(s && s !== '-') {
                    const st = resolveStatus(null, null, s);
                    const opt = document.createElement('option');
                    opt.value = s;
                    opt.textContent = `${s} (${st})`;
                    select.appendChild(opt);
                }
            });
            select.value = surveyors.includes(currentVal) ? currentVal : 'ALL';

            const palette = ['#2563eb', '#dc2626', '#059669', '#d97706', '#7c3aed', '#db2777', '#0891b2', '#4f46e5', '#ca8a04', '#65a30d'];
            surveyors.forEach((s, idx) => { surveyorColorMap[s] = palette[idx % palette.length]; });
        }

        function populateMapOutletDropdown(dataList) {
            const select = document.getElementById('mapOutletFilter');
            const currentVal = select.value;
            const defaultText = translations[currentLang].optAllOutlets;
            select.innerHTML = `<option value="ALL">${defaultText}</option>`;
            const outlets = [...new Set(dataList.map(item => item.surveyee))].sort();
            outlets.forEach(o => {
                if(o && o !== '-') {
                    const opt = document.createElement('option');
                    opt.value = o; opt.textContent = o;
                    select.appendChild(opt);
                }
            });
            select.value = outlets.includes(currentVal) ? currentVal : 'ALL';
        }

        function initMap() {
            // preferCanvas: true menggambar semua marker di satu layer canvas agar tetap 100% muncul tanpa berat DOM
            map = L.map('map', { center: [-2.5489, 118.0149], zoom: 5, tap: true, preferCanvas: true });
            const osmStandard = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: '&copy; OpenStreetMap' });
            osmStandard.addTo(map);
            markersLayer = L.layerGroup().addTo(map);
            initMapFilters();
        }

        function getMapFilteredSubset(region, asm, surveyor) {
            let subset = surveyData;
            if (region && region !== 'ALL') subset = subset.filter(item => item.region === region);
            if (asm && asm !== 'ALL') subset = subset.filter(item => item.asm === asm);
            if (surveyor && surveyor !== 'ALL') subset = subset.filter(item => item.surveyor === surveyor);
            return subset;
        }

        function onMapRegionChange() {
            const reg = document.getElementById('mapRegionFilter').value;
            const asm = document.getElementById('mapAsmFilter').value;
            const srv = document.getElementById('mapSurveyorFilter').value;

            populateMapAsmDropdown(getMapFilteredSubset(reg, 'ALL', 'ALL'));
            populateMapSurveyorDropdown(getMapFilteredSubset(reg, asm, 'ALL'));
            populateMapOutletDropdown(getMapFilteredSubset(reg, asm, srv));

            renderMapDashboard(reg, document.getElementById('mapAsmFilter').value, document.getElementById('mapSurveyorFilter').value, 'ALL');
        }

        function onMapAsmChange() {
            const reg = document.getElementById('mapRegionFilter').value;
            const asm = document.getElementById('mapAsmFilter').value;
            const srv = document.getElementById('mapSurveyorFilter').value;

            populateMapSurveyorDropdown(getMapFilteredSubset(reg, asm, 'ALL'));
            populateMapOutletDropdown(getMapFilteredSubset(reg, asm, srv));

            renderMapDashboard(reg, asm, document.getElementById('mapSurveyorFilter').value, 'ALL');
        }

        function onMapSurveyorChange() {
            const reg = document.getElementById('mapRegionFilter').value;
            const asm = document.getElementById('mapAsmFilter').value;
            const srv = document.getElementById('mapSurveyorFilter').value;

            populateMapOutletDropdown(getMapFilteredSubset(reg, asm, srv));
            renderMapDashboard(reg, asm, srv, 'ALL');
        }

        function onMapOutletChange() {
            renderMapDashboard(
                document.getElementById('mapRegionFilter').value,
                document.getElementById('mapAsmFilter').value,
                document.getElementById('mapSurveyorFilter').value,
                document.getElementById('mapOutletFilter').value
            );
        }

        function resetMapFilters() {
            document.getElementById('mapRegionFilter').value = 'ALL';
            document.getElementById('mapAsmFilter').value = 'ALL';
            document.getElementById('mapSurveyorFilter').value = 'ALL';
            document.getElementById('mapOutletFilter').value = 'ALL';

            populateMapAsmDropdown(surveyData);
            populateMapSurveyorDropdown(surveyData);
            populateMapOutletDropdown(surveyData);
            renderMapDashboard('ALL', 'ALL', 'ALL', 'ALL');
        }

        function setMapLoadingProgress(processed, total) {
            const overlay = document.getElementById('mapLoadingOverlay');
            const bar = document.getElementById('mapLoadingBar');
            const detail = document.getElementById('mapLoadingDetail');
            const pctEl = document.getElementById('mapLoadingPct');
            const titleEl = document.getElementById('mapLoadingTitle');

            if (!overlay) return;
            titleEl.textContent = translations[currentLang].loadingPointsTitle;
            const pct = total > 0 ? Math.min(100, Math.round((processed / total) * 100)) : 100;
            bar.style.width = `${pct}%`;
            detail.textContent = `${fmtNum(processed)} / ${fmtNum(total)} ${currentLang === 'en' ? 'points' : 'titik'}`;
            pctEl.textContent = `${pct}%`;
            overlay.classList.add('active');
        }

        function hideMapLoadingOverlay() {
            const overlay = document.getElementById('mapLoadingOverlay');
            if (overlay) overlay.classList.remove('active');
        }

        function renderMapDashboard(reg, asm, srv, out) {
            const currentToken = ++mapRenderToken;
            const listContainer = document.getElementById('outletListContainer');
            listContainer.innerHTML = '';
            markersLayer.clearLayers();
            
            let filteredData = surveyData;
            if (reg !== 'ALL') filteredData = filteredData.filter(item => item.region === reg);
            if (asm !== 'ALL') filteredData = filteredData.filter(item => item.asm === asm);
            if (srv !== 'ALL') filteredData = filteredData.filter(item => item.surveyor === srv);
            if (out !== 'ALL') filteredData = filteredData.filter(item => item.surveyee === out);
            
            document.getElementById('mapTotalStat').innerText = filteredData.length;

            // Seluruh titik koordinat valid tetap ditampilkan 100%
            const validPoints = filteredData.filter(item => item.lat !== null && item.lon !== null);
            const totalPoints = validPoints.length;

            if (totalPoints === 0) {
                hideMapLoadingOverlay();
                setTimeout(() => { if (map) map.invalidateSize(); }, 100);
                return;
            }

            setMapLoadingProgress(0, totalPoints);

            const bounds = [];
            const chunkSize = 200;
            let idx = 0;

            function processNextChunk() {
                if (currentToken !== mapRenderToken) return;

                const end = Math.min(idx + chunkSize, totalPoints);
                const listFrag = document.createDocumentFragment();
                const batchMarkers = [];

                for (; idx < end; idx++) {
                    const item = validPoints[idx];
                    bounds.push([item.lat, item.lon]);
                    const st = resolveStatus(item.status, item.surveyor_id, item.surveyor);

                    const itemDiv = document.createElement('div');
                    itemDiv.className = 'outlet-item';
                    itemDiv.innerHTML = `<div class="outlet-name">${item.surveyee}</div><div class="outlet-info"><span>ID: ${item.surveyee_id}</span><span style="color: #2563eb; font-weight: 600;">Qty: ${fmtNum(item.qty)}</span></div><div style="font-size: 10px; color: #64748b; margin-top: 2px;">${renderSurveyorLabelHtml(item.surveyor, item.position, st)}</div><div style="font-size: 10px; color: #94a3b8; margin-top: 2px;">📍 ${item.lat}, ${item.lon}</div>`;
                    
                    const marker = buildLeafletMarker(item);
                    batchMarkers.push(marker);

                    itemDiv.onclick = function() {
                        document.querySelectorAll('.outlet-item').forEach(el => el.classList.remove('active'));
                        itemDiv.classList.add('active');
                        map.setView([item.lat, item.lon], 16);
                        marker.fire('click');
                    };
                    listFrag.appendChild(itemDiv);
                }

                listContainer.appendChild(listFrag);
                batchMarkers.forEach(m => markersLayer.addLayer(m));
                setMapLoadingProgress(idx, totalPoints);

                if (idx < totalPoints) {
                    requestAnimationFrame(processNextChunk);
                } else {
                    if (bounds.length > 0 && currentToken === mapRenderToken) {
                        map.fitBounds(bounds, { padding: [30, 30], maxZoom: 15 });
                    }
                    setTimeout(() => {
                        if (currentToken !== mapRenderToken) return;
                        if (map) map.invalidateSize();
                        hideMapLoadingOverlay();
                    }, 80);
                }
            }

            requestAnimationFrame(processNextChunk);
        }

        function buildLeafletMarker(item) {
            const color = surveyorColorMap[item.surveyor] || '#2563eb';
            const marker = L.circleMarker([item.lat, item.lon], {
                radius: 6,
                fillColor: color,
                color: '#ffffff',
                weight: 2,
                opacity: 1,
                fillOpacity: 0.9
            });

            // Isi popup dirakit saat marker diklik agar render awal seluruh titik tetap ringan
            marker.on('click', function() {
                if (!marker.getPopup()) {
                    const st = resolveStatus(item.status, item.surveyor_id, item.surveyor);
                    const surveyorDisplay = renderSurveyorLabelHtml(item.surveyor, item.position, st);
                    const t = translations[currentLang];

                    let fotoHtml = '';
                    if (item.foto_list && item.foto_list.length > 0) {
                        fotoHtml = '<div style="margin-top: 6px; text-align: center; max-height: 200px; overflow-y: auto;">';
                        item.foto_list.forEach((fPath, idx) => {
                            fotoHtml += `<div style="margin-bottom: 4px;"><a href="${fPath}" target="_blank"><img src="${fPath}" style="max-width: 100%; width: 200px; height: auto; border-radius: 4px; border: 1px solid #cbd5e1;"></a><small style="display:block; font-size:9px; color:#64748b;">${t.mapPopupPhotoLabel} ${idx+1}</small></div>`;
                        });
                        fotoHtml += '</div>';
                    } else {
                        fotoHtml = `<div style="margin-top: 6px; color: #64748b; font-style: italic; font-size: 10px;">${t.mapPopupNoPhoto}</div>`;
                    }

                    const mapsUrl = `https://www.google.com/maps/search/?api=1&query=${item.lat},${item.lon}`;
                    const popupContent = `<div style="max-width: 240px;"><b>${item.surveyee}</b><br>ID: <b>${item.surveyee_id}</b><hr style="margin: 4px 0;"><small>Surveyor: ${surveyorDisplay}<br>Region/ASM: ${item.region} / ${item.asm}<br><b>Latitude, Longitude:</b> ${item.lat}, ${item.lon}<br><b>Qty:</b> ${fmtNum(item.qty)}<br><b>Timestamp:</b> ${item.timestamp}</small><div style="margin-top: 6px;"><a href="${mapsUrl}" target="_blank" style="display: inline-block; padding: 4px 8px; background: #2563eb; color: white; text-decoration: none; border-radius: 4px; font-size: 11px; font-weight: 600;">${t.mapPopupMapsBtn}</a></div>${fotoHtml}</div>`;
                    marker.bindPopup(popupContent);
                }
                marker.openPopup();
            });

            return marker;
        }

        function downloadFilteredExcel() {
            const regVal = document.getElementById('tableRegionFilter').value;
            const asmVal = document.getElementById('tableAsmFilter').value;
            const posVal = document.getElementById('tablePositionFilter').value;
            const srvVal = document.getElementById('tableSurveyorFilter').value;

            let filteredMasters = masterSurveyor.filter(m => {
                if (regVal && regVal !== 'ALL' && m.region !== regVal) return false;
                if (asmVal && asmVal !== 'ALL' && m.asm !== asmVal) return false;
                if (posVal && posVal !== 'ALL' && m.position.toUpperCase() !== posVal.toUpperCase()) return false;
                if (srvVal && srvVal !== 'ALL' && m.surveyor !== srvVal) return false;
                return true;
            });

            const filteredSurveyorIds = new Set(filteredMasters.map(m => m.surveyor_id));
            let exportData = surveyData.filter(item => filteredSurveyorIds.has(item.surveyor_id));

            let rowsToExport = exportData.map(item => {
                let mInfo = masterSurveyor.find(m => m.surveyor_id === item.surveyor_id) || {};
                let st = resolveStatus(mInfo.status || item.status, item.surveyor_id, item.surveyor);
                return {
                    "Timestamp": item.timestamp,
                    "Surveyor ID": item.surveyor_id,
                    "Surveyor Name": `${item.surveyor} (${st})`,
                    "Status": st,
                    "Surveyee ID": item.surveyee_id,
                    "Surveyee Name": item.surveyee,
                    "Position": mInfo.position || "-",
                    "Region": mInfo.region || "-",
                    "ASM": mInfo.asm || "-",
                    "Latitude": item.lat !== null ? item.lat : "-",
                    "Longitude": item.lon !== null ? item.lon : "-",
                    "Qty": item.qty,
                    "Foto List": item.foto_list ? item.foto_list.join(", ") : "-"
                };
            });

            if (rowsToExport.length === 0) {
                alert(currentLang === 'en' ? "No data matches the selected filters." : "Tidak ada data yang sesuai dengan filter yang dipilih.");
                return;
            }

            let now = new Date();
            let yyyy = now.getFullYear();
            let mm = String(now.getMonth() + 1).padStart(2, '0');
            let dd = String(now.getDate()).padStart(2, '0');
            let hh = String(now.getHours()).padStart(2, '0');
            let min = String(now.getMinutes()).padStart(2, '0');
            let ss = String(now.getSeconds()).padStart(2, '0');
            let timestampStr = `${yyyy}${mm}${dd}_${hh}${min}${ss}`;

            let regionPart = (regVal && regVal !== 'ALL') ? regVal : (currentLang === 'en' ? 'National' : 'Nasional');
            let asmPart = (asmVal && asmVal !== 'ALL') ? asmVal : (currentLang === 'en' ? 'AllASM' : 'SemuaASM');
            let srvPart = (srvVal && srvVal !== 'ALL') ? srvVal : (currentLang === 'en' ? 'AllSurveyor' : 'SemuaSurveyor');

            let filename = `${regionPart} - ${asmPart} - ${srvPart} - ${timestampStr}.xlsx`;

            const worksheet = XLSX.utils.json_to_sheet(rowsToExport);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Filtered Report");

            XLSX.writeFile(workbook, filename);
        }

        window.onload = function() {
            initTableFilters();
            const luEl = document.getElementById('lastUpdateSpan');
            if (luEl) luEl.innerText = rawPayload.lastUpdate || '-';
            syncHeaderOffset();
            window.addEventListener('resize', syncHeaderOffset);
            if (typeof ResizeObserver !== 'undefined') {
                const hdr = document.querySelector('header');
                if (hdr) new ResizeObserver(syncHeaderOffset).observe(hdr);
            }
        };
    </script>
</body>
</html>
"""

        # Inject Logos & Payload into HTML Template
        logo_50_tag = f'<img src="data:image/png;base64,{logo_50_base64}" alt="Sildenafil 50 mg" />' if logo_50_base64 else ''
        logo_100_tag = f'<img src="data:image/png;base64,{logo_100_base64}" alt="Sildenafil 100 mg" />' if logo_100_base64 else ''

        final_html = html_template.replace("__LOGO_50_TAG__", logo_50_tag)
        final_html = final_html.replace("__LOGO_100_TAG__", logo_100_tag)
        final_html = final_html.replace("__PAYLOAD_JSON_PLACEHOLDER__", data_json)

        with open(FILE_HTML_PETA, 'w', encoding='utf-8') as f_html:
            f_html.write(final_html)

        print(f"\n[RINGKASAN PROSES FINAL]")
        print(f"1. Total baris data diproses                : {len(list_log_detail)} baris")
        print(f"2. Total baris duplikat identik dihapus       : {total_duplikat_dibuang} baris")
        print(f"3. Total baris digabung (waktu/kunjungan beda): {total_baris_digabung} baris")
        print(f"4. Total outlet unik tersimpan di Excel Utama: {len(df_clean)} baris")
        print(f"5. Total foto unik yang berhasil didownload    : {total_foto_diekstrak} foto")
        print(f"6. Total foto gagal didownload                : {total_foto_gagal} foto")
        print(f"7. Total outlet/baris tanpa foto              : {total_tanpa_foto} baris")
        print(f"8. Total titik GPS valid di Peta GIS          : {len(list_data_peta)} titik")
        print(f"9. File Excel Utama disimpan di               : '{path_excel_output}'")
        print(f"10. File Log Detail disimpan di               : '{FILE_LOG}'")
        print(f"11. File Peta HTML (Visualisasi Titik)        : '{FILE_HTML_PETA}'")
        
        # --- AUTO-PUSH KE GITHUB (AMAN JIKA GIT TIDAK ADA ATAU ADA LOCK) ---
        print("\n-> Sedang melakukan Auto-Sync & Push ke GitHub...")
        if shutil.which("git") is None:
            print("[INFO] Program Git tidak terdeteksi di perangkat ini.")
            print("[INFO] Proses Excel & Dashboard GIS Selesai dengan sukses! Fitur auto-push ke GitHub dilewati.")
        else:
            try:
                waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                gitignore_path = os.path.join(BASE_DIR, ".gitignore")
                ignore_content = ""
                if os.path.exists(gitignore_path):
                    with open(gitignore_path, "r", encoding="utf-8") as gf:
                        ignore_content = gf.read()
                if "source/" not in ignore_content:
                    with open(gitignore_path, "a", encoding="utf-8") as gf:
                        gf.write("\nsource/\n")

                git_path = r"C:\Program Files\Git\cmd\git.exe"
                if not os.path.exists(git_path):
                    git_path = "git"

                # Nonaktifkan auto gc agar tidak memicu penghapusan folder .git/objects/xx saat file sedang terkunci
                subprocess.run(f'"{git_path}" config gc.auto 0', cwd=BASE_DIR, shell=True, check=False, timeout=15)
                subprocess.run(f'"{git_path}" add .', cwd=BASE_DIR, shell=True, check=True, timeout=60, input="n\n", text=True)
                
                commit_res = subprocess.run(
                    f'"{git_path}" commit -m "Update dashboard & data on {waktu_sekarang}"',
                    cwd=BASE_DIR, shell=True, capture_output=True, text=True, timeout=30, input="n\n"
                )
                if commit_res.returncode != 0 and "nothing to commit" not in (commit_res.stdout + commit_res.stderr).lower():
                    raise subprocess.CalledProcessError(commit_res.returncode, commit_res.args, output=commit_res.stdout, stderr=commit_res.stderr)

                subprocess.run(f'"{git_path}" push', cwd=BASE_DIR, shell=True, check=True, timeout=120, input="n\n", text=True)
                print("[SUKSES] Perubahan berhasil di-push otomatis ke GitHub! 🎉")
            except subprocess.CalledProcessError as git_err:
                print(f"[WARNING] Gagal auto-push ke GitHub: {git_err}")
                print("[INFO] Tips jika muncul 'Deletion of directory .git/objects/... failed': tutup aplikasi/editor/antivirus yang mengunci folder .git, lalu jalankan 'git gc --prune=now' secara manual.")
            except subprocess.TimeoutExpired:
                print("[WARNING] Proses Git melebihi batas waktu (timeout).")
        
    except PermissionError as e:
        print(f"\n[ERROR] Gagal menyimpan file karena sedang dibuka di Microsoft Excel! ({e})")
else:
    print("[WARNING] Total baris gabungan kosong.")

print("\n=== [SELESAI ALL PROCESS] == Dashboard Updated Successfully ==")