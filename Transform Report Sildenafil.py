import os
import re
from openpyxl import load_workbook
import pandas as pd
from datetime import datetime
import json
import subprocess

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DIR_SOURCE = os.path.join(BASE_DIR, "source")
DIR_IMAGE = os.path.join(BASE_DIR, "image")
DIR_OUTPUT = os.path.join(BASE_DIR, "output")
FILE_MAPPING = os.path.join(BASE_DIR, "mapping_surveyor.xlsx")
FILE_LOG = os.path.join(BASE_DIR, "log_proses_survey.xlsx")
FILE_HTML_PETA = os.path.join(BASE_DIR, "index.html")

for direktori in [DIR_SOURCE, DIR_IMAGE, DIR_OUTPUT]:
    if not os.path.exists(direktori):
        os.makedirs(direktori)

if not os.path.exists(FILE_MAPPING):
    df_template = pd.DataFrame(columns=["Surveyor ID", "Sales Name", "Position", "Region", "ASM"])
    df_template.to_excel(FILE_MAPPING, index=False)
    print(f"[INFO] File '{os.path.basename(FILE_MAPPING)}' baru saja dibuat.")
    print("-> Silakan isi data Surveyor kamu di file Excel tersebut dahulu, lalu jalankan ulang script ini.")
    exit()

print("-> Memuat data master dari Excel mapping...")
df_mapping_excel = pd.read_excel(FILE_MAPPING)

MASTER_SURVEYOR = {}
for idx, row in df_mapping_excel.iterrows():
    if pd.isna(row["Surveyor ID"]):
        continue
    sid = str(row["Surveyor ID"]).strip()
    MASTER_SURVEYOR[sid] = {
        "Sales Name": str(row["Sales Name"]).strip() if not pd.isna(row["Sales Name"]) else "Unknown",
        "Position": str(row["Position"]).strip() if not pd.isna(row["Position"]) else "-",
        "Region": str(row["Region"]).strip() if not pd.isna(row["Region"]) else "-",
        "ASM": str(row["ASM"]).strip() if not pd.isna(row["ASM"]) else "-"
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

            # Batasi panjang nama folder & file untuk mencegah error "Filename too long" di Git/Windows
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

    for idx, row in df_clean.iterrows():
        sid = str(row['Surveyor ID']).strip() if 'Surveyor ID' in df_clean.columns else ""
        if sid in MASTER_SURVEYOR:
            list_position.append(MASTER_SURVEYOR[sid]["Position"])
            list_region.append(MASTER_SURVEYOR[sid]["Region"])
            list_asm.append(MASTER_SURVEYOR[sid]["ASM"])
        else:
            list_position.append("-")
            list_region.append("-")
            list_asm.append("-")

    df_clean["Position"] = list_position
    df_clean["Region"] = list_region
    df_clean["ASM"] = list_asm

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
                "asm": info["ASM"]
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
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
    <title>Dashboard Sildenafil Package - Performance & GIS Map</title>
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
        html, body { height: 100%; margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f1f5f9; color: #1e293b; }
        
        /* Navigation / Header */
        header { min-height: 60px; background-color: #ffffff; color: #0f172a; padding: 10px 20px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #cbd5e1; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); position: fixed; top: 0; left: 0; right: 0; z-index: 3000; flex-wrap: wrap; gap: 10px; }
        header h1 { margin: 0; font-size: 15px; font-weight: 600; display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
        .nav-buttons { display: flex; gap: 8px; flex-wrap: wrap; }
        .nav-btn { background: #e2e8f0; color: #334155; border: none; padding: 6px 12px; border-radius: 6px; font-size: 11px; font-weight: 600; cursor: pointer; transition: all 0.2s; white-space: nowrap; }
        .nav-btn.active { background: #2563eb; color: white; }
        .nav-btn:hover { opacity: 0.9; }

        /* Views Container */
        .view-container { margin-top: 70px; height: calc(100vh - 70px); width: 100vw; position: relative; overflow: hidden; }

        @media (max-width: 768px) {
            header { padding: 10px 12px; flex-direction: column; align-items: flex-start; gap: 8px; }
            .nav-buttons { width: 100%; justify-content: space-between; }
            .nav-btn { flex: 1; text-align: center; padding: 8px 6px; font-size: 11px; }
            .view-container { margin-top: 115px; height: calc(100vh - 115px); }
        }
        .view-page { position: absolute; top: 0; left: 0; width: 100%; height: 100%; display: none; background: #f1f5f9; overflow-y: auto; }
        .view-page.active { display: flex; flex-direction: column; }

        /* Summary Table View Styles */
        .table-content { padding: 20px; max-width: 1400px; margin: 0 auto; width: 100%; }
        .filter-bar { background: white; padding: 15px; border-radius: 8px; border: 1px solid #cbd5e1; display: flex; gap: 15px; align-items: center; margin-bottom: 20px; flex-wrap: wrap; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
        .filter-field { display: flex; flex-direction: column; gap: 4px; flex: 1; min-width: 180px; }
        .filter-field label { font-size: 11px; font-weight: 600; color: #64748b; }
        .filter-field select { padding: 8px 12px; border-radius: 6px; border: 1px solid #cbd5e1; font-size: 12px; background: #f8fafc; outline: none; }
        
        .summary-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 15px; margin-bottom: 20px; }
        .scard { background: white; padding: 15px; border-radius: 8px; border: 1px solid #cbd5e1; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
        .scard h4 { margin: 0 0 5px 0; font-size: 11px; color: #64748b; text-transform: uppercase; }
        .scard .sval { font-size: 20px; font-weight: bold; color: #0f172a; }

        .table-responsive { background: white; border-radius: 8px; border: 1px solid #cbd5e1; overflow-x: auto; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
        table { width: 100%; border-collapse: collapse; text-align: left; font-size: 12px; }
        th, td { padding: 10px 12px; border-bottom: 1px solid #e2e8f0; }
        th { background: #f8fafc; font-weight: 600; color: #475569; position: sticky; top: 0; z-index: 10; }
        tr:hover { background: #f8fafc; }
        .badge { padding: 3px 8px; border-radius: 4px; font-size: 10px; font-weight: 600; }
        .badge-sbd { background: #dbeafe; color: #1e40af; }
        .badge-msd { background: #fef3c7; color: #92400e; }
        .badge-asm { background: #d1fae5; color: #065f46; }
        .badge-tss { background: #fae8ff; color: #86198f; }

        /* Map View Styles */
        .map-layout { display: flex; height: 100%; width: 100%; position: relative; }
        .sidebar { width: 380px; background-color: #ffffff; border-right: 1px solid #cbd5e1; display: flex; flex-direction: column; padding: 15px; gap: 10px; z-index: 1000; box-shadow: 4px 0 15px rgba(0,0,0,0.05); overflow-y: auto; }
        .sidebar-filter { display: flex; flex-direction: column; gap: 3px; }
        .sidebar-filter label { font-size: 11px; font-weight: 600; color: #64748b; }
        .sidebar-filter select { padding: 7px 10px; border-radius: 6px; border: 1px solid #cbd5e1; background-color: #f8fafc; color: #1e293b; font-size: 12px; outline: none; }
        .outlet-list { flex: 1; min-height: 150px; overflow-y: auto; display: flex; flex-direction: column; gap: 6px; padding-right: 2px; }
        .outlet-item { background-color: #f8fafc; border: 1px solid #cbd5e1; padding: 8px 10px; border-radius: 6px; cursor: pointer; transition: all 0.2s ease; }
        .outlet-item:hover { border-color: #2563eb; background-color: #eff6ff; }
        .outlet-item.active { border-color: #2563eb; background-color: #dbeafe; }
        .outlet-name { font-weight: 600; font-size: 12px; color: #0f172a; margin-bottom: 2px; }
        .outlet-info { font-size: 11px; color: #64748b; display: flex; justify-content: space-between; }
        #map { flex: 1; height: 100%; width: 100%; z-index: 1; }

        @media (max-width: 768px) {
            .map-layout { flex-direction: column-reverse; }
            .sidebar { width: 100%; height: 45%; max-height: 250px; }
            #map { height: 55% !important; }
        }
        ::-webkit-scrollbar { width: 5px; height: 5px; }
        ::-webkit-scrollbar-track { background: #f1f5f9; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
    </style>
</head>
<body>
    <header>
        <h1>🗺️ Sildenafil Package Dashboard <span style="font-size: 11px; font-weight: normal; color: #64748b; margin-left: 12px; background: #f8fafc; padding: 4px 8px; border: 1px solid #cbd5e1; border-radius: 4px;">🕒 Terakhir Diperbarui: <span id="lastUpdateSpan"></span></span></h1>
        <div class="nav-buttons">
            <button id="btnTableView" class="nav-btn active" onclick="switchView('table')">📊 Tabel Target vs Actual</button>
            <button id="btnMapView" class="nav-btn" onclick="switchView('map')">📍 Visualisasi Sebaran Titik Distribusi</button>
        </div>
    </header>

    <div class="view-container">
        <!-- VIEW 1: TABEL RINGKASAN TARGET VS ACTUAL (CASCADING FILTER) -->
        <div id="viewTable" class="view-page active">
            <div class="table-content">
                <div class="filter-bar">
                    <div class="filter-field">
                        <label for="tableRegionFilter">Filter Region:</label>
                        <select id="tableRegionFilter" onchange="onTableRegionChange()">
                            <option value="ALL">-- Semua Region --</option>
                        </select>
                    </div>
                    <div class="filter-field">
                        <label for="tableAsmFilter">Filter ASM:</label>
                        <select id="tableAsmFilter" onchange="onTableAsmChange()">
                            <option value="ALL">-- Semua ASM --</option>
                        </select>
                    </div>
                    <div class="filter-field">
                        <label for="tablePositionFilter">Filter Role:</label>
                        <select id="tablePositionFilter" onchange="onTablePositionChange()">
                            <option value="ALL">-- Semua Role --</option>
                            <option value="SBD">SBD</option>
                            <option value="MSD">MSD</option>
                            <option value="ASM">ASM</option>
                            <option value="TSS">TSS</option>
                        </select>
                    </div>
                    <div class="filter-field">
                        <label for="tableSurveyorFilter">Filter Surveyor:</label>
                        <select id="tableSurveyorFilter" onchange="renderTableData()">
                            <option value="ALL">-- Semua Surveyor --</option>
                        </select>
                    </div>
                    <div style="display: flex; align-items: flex-end; gap: 8px;">
                        <button onclick="resetTableFilters()" style="padding: 8px 14px; background: #cbd5e1; border: none; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer;">Atur Ulang Filter</button>
                        <button onclick="downloadFilteredExcel()" style="padding: 8px 14px; background: #059669; color: white; border: none; border-radius: 6px; font-size: 12px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; gap: 6px;">📥 Unduh Excel Sesuai Filter</button>
                    </div>
                </div>

                <div class="summary-cards">
                    <div class="scard" style="background-color: #fef2f2; border-color: #fca5a5;">
                        <h4 style="color: #991b1b;">Total Target Distribusi Sildenafil Package</h4>
                        <div id="sumTargetPkg" class="sval" style="color: #b91c1c;">0</div>
                    </div>
                    <div class="scard" style="background-color: #fef2f2; border-color: #fca5a5;">
                        <h4 style="color: #991b1b;">Total Actual Distribusi Sildenafil Package</h4>
                        <div id="sumActualPkg" class="sval" style="color: #b91c1c;">0 <small id="sumPkgPct" style="font-size:12px; font-weight:normal; color:#b91c1c;">(0.0% vs Target)</small></div>
                    </div>
                    <div class="scard" style="background-color: #eff6ff; border-color: #93c5fd;">
                        <h4 style="color: #1e40af;">Total Target Customer</h4>
                        <div id="sumTargetCust" class="sval" style="color: #1d4ed8;">0</div>
                    </div>
                    <div class="scard" style="background-color: #eff6ff; border-color: #93c5fd;">
                        <h4 style="color: #1e40af;">Total Actual Customer</h4>
                        <div id="sumActualCust" class="sval" style="color: #1d4ed8;">0 <small id="sumCustPct" style="font-size:12px; font-weight:normal; color:#1e40af;">(0.0% vs Target)</small></div>
                    </div>
                </div>

                <div class="table-responsive">
                    <table>
                        <thead>
                            <tr>
                                <th>Region</th>
                                <th>ASM</th>
                                <th>Role</th>
                                <th>Nama Surveyor</th>
                                <th style="text-align: center;">Distribusi Sildenafil Package<br><small style="font-weight: normal; color: #64748b;">(Aktual vs Target)</small></th>
                                <th style="text-align: center;">Customer Sildenafil Package<br><small style="font-weight: normal; color: #64748b;">(Aktual vs Target)</small></th>
                            </tr>
                        </thead>
                        <tbody id="summaryTableBody">
                            <!-- Diisi via JavaScript -->
                        </tbody>
                    </table>
                </div>
                <div style="margin-top: 15px; text-align: right;">
                    <button onclick="switchView('map')" style="padding: 10px 20px; background: #2563eb; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; font-size: 13px;">Buka Visualisasi Sebaran Titik Distribusi ➡️</button>
                </div>
            </div>
        </div>

        <!-- VIEW 2: PETA GIS & SEBARAN -->
        <div id="viewMap" class="view-page">
            <div class="map-layout">
                <div class="sidebar">
                    <div class="sidebar-filter">
                        <label for="mapRegionFilter">Filter Region:</label>
                        <select id="mapRegionFilter" onchange="onMapRegionChange()">
                            <option value="ALL">-- Semua Region --</option>
                        </select>
                    </div>
                    <div class="sidebar-filter">
                        <label for="mapAsmFilter">Filter ASM:</label>
                        <select id="mapAsmFilter" onchange="onMapAsmChange()">
                            <option value="ALL">-- Semua ASM --</option>
                        </select>
                    </div>
                    <div class="sidebar-filter">
                        <label for="mapSurveyorFilter">Filter Surveyor:</label>
                        <select id="mapSurveyorFilter" onchange="onMapSurveyorChange()">
                            <option value="ALL">-- Semua Surveyor --</option>
                        </select>
                    </div>
                    <div class="sidebar-filter">
                        <label for="mapOutletFilter">Filter Outlet (Surveyee):</label>
                        <select id="mapOutletFilter" onchange="onMapOutletChange()">
                            <option value="ALL">-- Semua Outlet --</option>
                        </select>
                    </div>
                    <div>
                        <button onclick="resetMapFilters()" style="width: 100%; padding: 7px 10px; background: #cbd5e1; border: none; border-radius: 6px; font-size: 11px; font-weight: 600; cursor: pointer;">Atur Ulang Filter Peta</button>
                    </div>
                    <div style="background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%); padding: 10px; border-radius: 8px; color: white;">
                        <h3 style="margin: 0 0 2px 0; font-size: 10px; text-transform: uppercase;">Total Titik Sebaran Peta</h3>
                        <div id="mapTotalStat" class="val" style="font-size: 18px; font-weight: bold;">0</div>
                    </div>
                    <div style="font-size: 11px; font-weight: 600; color: #64748b;">Daftar Outlet Terfilter:</div>
                    <div id="outletListContainer" class="outlet-list"></div>
                </div>
                <div id="map"></div>
            </div>
        </div>
    </div>

    <script>
        const rawPayload = __PAYLOAD_JSON_PLACEHOLDER__;
        const surveyData = rawPayload.surveyData || [];
        const masterSurveyor = rawPayload.masterSurveyor || [];

        const targetPerPerson = {
            "SBD": { pkg: 600, cust: 300 },
            "MSD": { pkg: 200, cust: 100 },
            "ASM": { pkg: 40, cust: 4 },
            "TSS": { pkg: 20, cust: 2 }
        };

        const fmtNum = (num) => Number(num || 0).toLocaleString('en-US');

        let surveyorColorMap = {};
        let map, markersLayer;

        function switchView(viewName) {
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
                    setTimeout(() => { map.invalidateSize(); }, 200);
                }
                renderMapDashboard('ALL', 'ALL', 'ALL', 'ALL');
            }
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
            regSelect.innerHTML = '<option value="ALL">-- Semua Region --</option>';
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
            asmSelect.innerHTML = '<option value="ALL">-- Semua ASM --</option>';
            
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
            srvSelect.innerHTML = '<option value="ALL">-- Semua Surveyor --</option>';

            let subset = masterSurveyor;
            if (region && region !== 'ALL') subset = subset.filter(m => m.region === region);
            if (asm && asm !== 'ALL') subset = subset.filter(m => m.asm === asm);
            if (position && position !== 'ALL') subset = subset.filter(m => m.position.toUpperCase() === position.toUpperCase());

            const surveyors = [...new Set(subset.map(m => m.surveyor))].filter(s => s && s !== '-').sort();
            surveyors.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s; opt.textContent = s;
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

            const tbody = document.getElementById('summaryTableBody');
            tbody.innerHTML = '';

            let totalTgtPkg = 0;
            let totalActPkg = 0;
            let totalTgtCust = 0;
            let totalActCustSet = new Set();

            filteredMasters.forEach(m => {
                let pos = m.position.toUpperCase().trim();
                let tPkg = (targetPerPerson[pos] ? targetPerPerson[pos].pkg : 0);
                let tCust = (targetPerPerson[pos] ? targetPerPerson[pos].cust : 0);

                let aPkg = actualPkgMap[m.surveyor_id] || 0;
                let aCustSet = actualCustMap[m.surveyor_id] || new Set();
                let aCust = aCustSet.size;

                totalTgtPkg += tPkg;
                totalActPkg += aPkg;
                totalTgtCust += tCust;
                
                surveyData.filter(d => d.surveyor_id === m.surveyor_id).forEach(d => {
                    if (d.surveyee_id) totalActCustSet.add(d.surveyee_id);
                });

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
                    <td><b>${m.surveyor}</b></td>
                    <td style="text-align: center;">
                        <b>${fmtNum(aPkg)}</b> / ${fmtNum(tPkg)} <br>
                        <small style="color: ${aPkg >= tPkg ? '#059669' : '#dc2626'}; font-weight: 600;">(${pPkg}%)</small>
                    </td>
                    <td style="text-align: center;">
                        <b>${fmtNum(aCust)}</b> / ${fmtNum(tCust)} <br>
                        <small style="color: ${aCust >= tCust ? '#059669' : '#dc2626'}; font-weight: 600;">(${pCust}%)</small>
                    </td>
                `;
                tbody.appendChild(tr);
            });

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
            select.innerHTML = '<option value="ALL">-- Semua ASM --</option>';
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
            select.innerHTML = '<option value="ALL">-- Semua Surveyor --</option>';
            const surveyors = [...new Set(dataList.map(item => item.surveyor))].sort();
            surveyors.forEach(s => {
                if(s && s !== '-') {
                    const opt = document.createElement('option');
                    opt.value = s; opt.textContent = s;
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
            select.innerHTML = '<option value="ALL">-- Semua Outlet --</option>';
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
            map = L.map('map', { center: [-2.5489, 118.0149], zoom: 5, tap: true });
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

        function renderMapDashboard(reg, asm, srv, out) {
            const listContainer = document.getElementById('outletListContainer');
            listContainer.innerHTML = '';
            markersLayer.clearLayers();
            
            let filteredData = surveyData;
            if (reg !== 'ALL') filteredData = filteredData.filter(item => item.region === reg);
            if (asm !== 'ALL') filteredData = filteredData.filter(item => item.asm === asm);
            if (srv !== 'ALL') filteredData = filteredData.filter(item => item.surveyor === srv);
            if (out !== 'ALL') filteredData = filteredData.filter(item => item.surveyee === out);
            
            document.getElementById('mapTotalStat').innerText = filteredData.length;

            let bounds = [];
            filteredData.forEach((item) => {
                if (item.lat !== null && item.lon !== null) {
                    bounds.push([item.lat, item.lon]);
                    const itemDiv = document.createElement('div');
                    itemDiv.className = 'outlet-item';
                    itemDiv.innerHTML = `<div class="outlet-name">${item.surveyee}</div><div class="outlet-info"><span>ID: ${item.surveyee_id}</span><span style="color: #2563eb; font-weight: 600;">Qty: ${fmtNum(item.qty)}</span></div><div style="font-size: 10px; color: #94a3b8; margin-top: 2px;">📍 ${item.lat}, ${item.lon}</div>`;
                    
                    let marker = createLeafletMarker(item);

                    itemDiv.onclick = function() {
                        document.querySelectorAll('.outlet-item').forEach(el => el.classList.remove('active'));
                        itemDiv.classList.add('active');
                        map.setView([item.lat, item.lon], 16);
                        marker.openPopup();
                    };
                    listContainer.appendChild(itemDiv);
                }
            });

            if (bounds.length > 0) {
                map.fitBounds(bounds, { padding: [30, 30], maxZoom: 15 });
            }
            setTimeout(() => { map.invalidateSize(); }, 200);
        }

        function createLeafletMarker(item) {
            const color = surveyorColorMap[item.surveyor] || '#2563eb';
            const customIcon = L.divIcon({ className: 'custom-div-icon', html: `<div style="background-color: ${color}; width: 14px; height: 14px; border-radius: 50%; border: 2px solid white; box-shadow: 0 0 6px rgba(0,0,0,0.4);"></div>`, iconSize: [14, 14], iconAnchor: [7, 7] });
            
            let fotoHtml = '';
            if (item.foto_list && item.foto_list.length > 0) {
                fotoHtml = '<div style="margin-top: 6px; text-align: center; max-height: 200px; overflow-y: auto;">';
                item.foto_list.forEach((fPath, idx) => {
                    fotoHtml += `<div style="margin-bottom: 4px;"><a href="${fPath}" target="_blank"><img src="${fPath}" width="200" style="border-radius: 4px; border: 1px solid #cbd5e1;"></a><small style="display:block; font-size:9px; color:#64748b;">Foto ${idx+1}</small></div>`;
                });
                fotoHtml += '</div>';
            } else {
                fotoHtml = '<div style="margin-top: 6px; color: #64748b; font-style: italic; font-size: 10px;">Tidak ada foto.</div>';
            }
            
            const mapsUrl = `https://www.google.com/maps/search/?api=1&query=${item.lat},${item.lon}`;
            const popupContent = `<b>${item.surveyee}</b><br>ID: <b>${item.surveyee_id}</b><br><hr style="margin: 4px 0;"><small>Surveyor: ${item.surveyor}<br>Region/ASM: ${item.region} / ${item.asm}<br><b>Latitude, Longitude:</b> ${item.lat}, ${item.lon}<br><b>Qty:</b> ${fmtNum(item.qty)}<br><b>Timestamp:</b> ${item.timestamp}</small><div style="margin-top: 6px;"><a href="${mapsUrl}" target="_blank" style="display: inline-block; padding: 4px 8px; background: #2563eb; color: white; text-decoration: none; border-radius: 4px; font-size: 11px; font-weight: 600;">🗺️ Buka Rute ke Google Maps</a></div>${fotoHtml}`;
            const marker = L.marker([item.lat, item.lon], { icon: customIcon }).bindPopup(popupContent);
            markersLayer.addLayer(marker);
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
                return {
                    "Timestamp": item.timestamp,
                    "Surveyor ID": item.surveyor_id,
                    "Surveyor Name": item.surveyor,
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
                alert("Tidak ada data yang sesuai dengan filter yang dipilih.");
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

            let regionPart = (regVal && regVal !== 'ALL') ? regVal : 'Nasional';
            let asmPart = (asmVal && asmVal !== 'ALL') ? asmVal : 'SemuaASM';
            let srvPart = (srvVal && srvVal !== 'ALL') ? srvVal : 'SemuaSurveyor';

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
        };
    </script>
</body>
</html>
"""

        final_html = html_template.replace("__PAYLOAD_JSON_PLACEHOLDER__", data_json)

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
        
        # --- AUTO-PUSH KE GITHUB ---
        print("\n-> Sedang melakukan Auto-Sync & Push ke GitHub...")
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

            subprocess.run(f'"{git_path}" add .', cwd=BASE_DIR, shell=True, check=True)
            subprocess.run(f'"{git_path}" commit -m "Auto-aggregate quantity on Excel output and update dashboard pada {waktu_sekarang}"', cwd=BASE_DIR, shell=True, check=True)
            subprocess.run(f'"{git_path}" push', cwd=BASE_DIR, shell=True, check=True)
            print("[SUKSES] Perubahan berhasil di-push otomatis ke GitHub! 🎉")
        except subprocess.CalledProcessError as git_err:
            print(f"[WARNING] Gagal auto-push ke GitHub: {git_err}")
        
    except PermissionError as e:
        print(f"\n[ERROR] Gagal menyimpan file karena sedang dibuka di Microsoft Excel! ({e})")
else:
    print("[WARNING] Total baris gabungan kosong.")

print("\n=== [SELESAI ALL PROCESS] == Dashboard Updated Successfully ==")