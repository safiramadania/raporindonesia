from __future__ import annotations

import csv
import hmac
import io
import json
import os
import sqlite3
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from flask import Flask, jsonify, render_template, request, send_file

ROOT = Path(__file__).resolve().parent
BUNDLE = ROOT / "model_bundle"
DB_PATH = ROOT / "rapor_local.sqlite3"
MAX_UPLOAD_BYTES = 16 * 1024 * 1024
MAX_ROWS = 20_000

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES
APP_ENV = os.getenv("APP_ENV", "local").casefold()
APP_ACCESS_PASSWORD = os.getenv("APP_ACCESS_PASSWORD", "")
if APP_ENV == "production" and not APP_ACCESS_PASSWORD:
    raise RuntimeError("APP_ACCESS_PASSWORD wajib diatur untuk menjalankan aplikasi online.")


@app.before_request
def require_deployment_password():
    if APP_ENV != "production" or request.path == "/healthz":
        return None
    auth = request.authorization
    supplied = auth.password if auth and auth.password else ""
    if hmac.compare_digest(supplied, APP_ACCESS_PASSWORD):
        return None
    response = jsonify({"error": "Masukkan kata sandi aplikasi untuk melanjutkan."})
    response.status_code = 401
    response.headers["WWW-Authenticate"] = 'Basic realm="Ruang Sekolah"'
    return response

sys.path.insert(0, str(BUNDLE))
from inference import load_predictor  # noqa: E402

predict, manifest = load_predictor(BUNDLE, device="cpu")
FEATURES = manifest["feature_columns"]
ID_COLUMN = manifest["school_id_column"]
TARGETS = manifest["targets"]
FIELD_SPECS = json.loads((ROOT / "input_dictionary.json").read_text(encoding="utf-8"))
CODE_TO_LABEL = {item["code"]: item["label"] for item in FIELD_SPECS}
LABEL_TO_CODE = {item["label"].strip().casefold(): item["code"] for item in FIELD_SPECS}
UI_FIELD_SPECS = [item for item in FIELD_SPECS if item["code"] != "kd_kokab"]
UI_FEATURES = [c for c in FEATURES if c != "kd_kokab"]
PREPROCESSOR = next(v.cell_contents for v in predict.__closure__
                    if hasattr(v.cell_contents, "num_cols_") and hasattr(v.cell_contents, "cat_cols_"))
NUM_FEATURES = set(PREPROCESSOR.num_cols_)
CAT_FEATURES = {
    "jenis_sek", "sts_sek", "daerah_khusus", "kd_kokab",
    "wilayah_bagian", "jenis_wilayah", "status_wilayah",
    *(f"bg{i}" for i in [1,2,3,4,5,6,7,8,9,10,11,12,15,16,17,18,19]),
}


def db():
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def init_db():
    with db() as con:
        con.execute("""CREATE TABLE IF NOT EXISTS prediction_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT NOT NULL,
            school_name TEXT NOT NULL, period TEXT NOT NULL, n_participants INTEGER NOT NULL,
            lit_mean REAL NOT NULL, num_mean REAL NOT NULL, lit_median REAL NOT NULL,
            num_median REAL NOT NULL, filename TEXT NOT NULL)""")
        con.execute("""CREATE TABLE IF NOT EXISTS action_plans (
            id INTEGER PRIMARY KEY AUTOINCREMENT, created_at TEXT NOT NULL,
            school_name TEXT NOT NULL, focus TEXT NOT NULL, action TEXT NOT NULL,
            owner TEXT NOT NULL, due_date TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'Direncanakan')""")


init_db()


@app.get("/")
def home():
    return render_template("index.html", feature_count=len(FEATURES), model=manifest["model"],
                           public_mode=APP_ENV == "public")


@app.get("/healthz")
def healthz():
    return jsonify({"status": "ok"})


@app.get("/api/meta")
def meta():
    return jsonify({"model": manifest["model"], "feature_count": len(FEATURES),
                    "fields": [{k: item[k] for k in ["code", "label", "group", "kind", "choices"]}
                               for item in UI_FIELD_SPECS if item["code"] in FEATURES],
                    "metrics": manifest.get("test_metrics", {})})


@app.get("/template.csv")
def template_csv():
    columns = ["Nomor Induk Peserta", "Nama sekolah", "Nama kabupaten/kota",
               *[CODE_TO_LABEL[c] for c in UI_FEATURES]]
    data = pd.DataFrame([{c: "" for c in columns}])
    raw = data.to_csv(index=False).encode("utf-8-sig")
    return send_file(io.BytesIO(raw), mimetype="text/csv", as_attachment=True,
                     download_name="template_input_rapor_diagnostik.csv")


@app.get("/template.xlsx")
def template_xlsx():
    wb = Workbook()
    start = wb.active
    start.title = "Mulai di sini"
    instructions = [
        ["PANDUAN INPUT DATA PESERTA"],
        ["Satu baris di lembar Isi data mewakili satu peserta. Isi nama sekolah dan nomor induk peserta."],
        ["Kabupaten/kota ditulis dengan nama biasa. Nama ini hanya menjadi keterangan dan belum digunakan model."],
        ["Isi informasi yang diketahui; kolom yang belum diketahui boleh dikosongkan. Jangan menebak jawaban atau nilai."],
        ["Kolom kosong akan ditangani oleh model. Baris yang seluruh prediktornya kosong tidak bisa diproses."],
        ["File diproses pada perangkat lokal. Aplikasi tidak menyimpan data mentah peserta."],
        ["Hasil aplikasi adalah estimasi penelitian, bukan nilai resmi AN dan bukan dasar pemeringkatan peserta."],
    ]
    for row in instructions: start.append(row)
    start.column_dimensions["A"].width = 115
    start["A1"].font = Font(bold=True, size=16, color="174F3C")
    for row in range(2, len(instructions)+1):
        start.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
        start.row_dimensions[row].height = 34

    sheet = wb.create_sheet("Isi data")
    headers = ["Nomor Induk Peserta", "Nama sekolah", "Nama kabupaten/kota",
               *[CODE_TO_LABEL[c] for c in UI_FEATURES]]
    sheet.append(headers)
    sheet.freeze_panes = "D2"
    wb.active = 0
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}201"
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", fgColor="176B54")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[1].height = 60
    sheet.column_dimensions["A"].width = 28
    sheet.column_dimensions["B"].width = 32
    sheet.column_dimensions["C"].width = 28
    for col in range(4, len(headers)+1): sheet.column_dimensions[get_column_letter(col)].width = 34

    option_sheet = wb.create_sheet("Pilihan")
    option_sheet.sheet_state = "hidden"
    for item in UI_FIELD_SPECS:
        if not item["choices"]:
            continue
        col_idx = option_sheet.max_column + 1
        letter = get_column_letter(col_idx)
        option_sheet.cell(1, col_idx, item["code"])
        for r_idx, choice in enumerate(item["choices"], start=2):
            option_sheet.cell(r_idx, col_idx, choice["label"])
        wb.defined_names.add(DefinedName(f"pilihan_{item['code']}",
            attr_text=f"'Pilihan'!${letter}$2:${letter}${len(item['choices'])+1}"))
        feature_col = headers.index(CODE_TO_LABEL[item["code"]]) + 1
        validation = DataValidation(type="list", formula1=f"=pilihan_{item['code']}", allow_blank=True)
        validation.errorTitle = "Pilih jawaban yang tersedia"
        validation.error = "Gunakan salah satu pilihan pada daftar."
        validation.showErrorMessage = True
        sheet.add_data_validation(validation)
        validation.add(f"{get_column_letter(feature_col)}2:{get_column_letter(feature_col)}201")
    for item in UI_FIELD_SPECS:
        if item["kind"] != "score":
            continue
        feature_col = headers.index(item["label"]) + 1
        validation = DataValidation(type="decimal", operator="between", formula1="0", formula2="100", allow_blank=True)
        validation.errorTitle = "Periksa rentang nilai"
        validation.error = "Masukkan nilai antara 0 dan 100."
        validation.showErrorMessage = True
        sheet.add_data_validation(validation)
        validation.add(f"{get_column_letter(feature_col)}2:{get_column_letter(feature_col)}201")

    guide = wb.create_sheet("Panduan kolom")
    guide.append(["Nama kolom", "Kelompok", "Cara mengisi", "Pilihan / catatan"])
    for cell in guide[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="176B54")
    guide.append(["Nomor Induk Peserta", "Identitas peserta", "Isi nomor induk peserta.", ""])
    guide.append(["Nama sekolah", "Identitas sekolah", "Isi nama sekolah.", ""])
    guide.append(["Nama kabupaten/kota", "Keterangan wilayah", "Isi nama kabupaten/kota dengan tulisan biasa.", "Hanya keterangan; belum digunakan model"])
    for item in UI_FIELD_SPECS:
        choices = "; ".join(ch["label"] for ch in item["choices"])
        if item["kind"] == "score":
            note = "Nilai angka 0–100; boleh dikosongkan jika belum diketahui."
        elif choices:
            note = "Pilih salah satu jawaban yang tersedia; boleh dikosongkan jika belum diketahui."
        else:
            note = "Isi jika diketahui; boleh dikosongkan."
        guide.append([item["label"], item["group"], note, choices])
    for col, width in zip("ABCD", [48, 32, 72, 100]): guide.column_dimensions[col].width = width
    guide.freeze_panes = "A2"
    guide.auto_filter.ref = guide.dimensions
    for row in guide.iter_rows(min_row=2):
        for cell in row: cell.alignment = Alignment(wrap_text=True, vertical="top")
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name="template_input_ramah_pengguna.xlsx")


def read_upload(file):
    filename = file.filename or "input"
    content = file.read()
    if not content:
        raise ValueError("File yang diunggah kosong.")
    ext = Path(filename).suffix.lower()
    if ext == ".csv":
        frame = pd.read_csv(io.BytesIO(content), dtype=str, low_memory=False)
    elif ext == ".xlsx":
        sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, dtype=str)
        frame = sheets.get("Isi data", next(iter(sheets.values())))
    else:
        raise ValueError("Gunakan file CSV atau Excel (.xlsx).")
    frame.columns = [str(c).strip() for c in frame.columns]
    renames = {
        "Nomor Induk Peserta": "id_peserta",
        "ID peserta (gunakan kode samaran)": "id_peserta",  # kompatibilitas file lama
        "Nama sekolah": ID_COLUMN,
        "Nama sekolah / ruang kerja": ID_COLUMN,
        "Kode sekolah": ID_COLUMN,  # kompatibilitas file lama
        "Nama kabupaten/kota": "nama_kokab",
        "Kabupaten/kota": "nama_kokab",
    }
    for c in frame.columns:
        friendly_code = LABEL_TO_CODE.get(c.casefold())
        if friendly_code:
            renames[c] = friendly_code
    frame = frame.rename(columns=renames)
    row_limit = 1_000 if APP_ENV == "public" else MAX_ROWS
    if len(frame) > row_limit:
        raise ValueError(f"Maksimal {row_limit:,} baris per proses.")
    if frame.empty:
        raise ValueError("Tidak ada baris peserta untuk diprediksi.")
    if ID_COLUMN not in frame:
        raise ValueError("File belum memuat kolom ‘Nama sekolah’.")
    if "id_peserta" not in frame:
        raise ValueError("File belum memuat kolom ‘Nomor Induk Peserta’.")
    provided_features = [c for c in FEATURES if c in frame.columns]
    if not provided_features:
        raise ValueError("Kolom prediktor belum dikenali. Gunakan nama kolom yang tersedia pada lembar ‘Panduan kolom’.")
    for c in FEATURES:
        if c not in frame.columns:
            frame[c] = pd.NA
    for c in FEATURES:
        if pd.api.types.is_object_dtype(frame[c]) or pd.api.types.is_string_dtype(frame[c]):
            frame[c] = frame[c].astype("string").str.strip().replace("", pd.NA)
    for item in FIELD_SPECS:
        c = item["code"]
        if not item["choices"] or c not in frame:
            continue
        lookup = {}
        for choice in item["choices"]:
            lookup[choice["label"].strip().casefold()] = choice["value"]
            lookup[str(choice["value"]).strip().casefold()] = choice["value"]
        for raw_value in frame[c].dropna().unique():
            value = str(raw_value).strip()
            if value.casefold() not in lookup:
                allowed = ", ".join(choice["label"] for choice in item["choices"][:6])
                raise ValueError(f"Jawaban pada kolom ‘{item['label']}’ belum dikenali: ‘{value}’. Pilihan: {allowed}.")
        frame[c] = frame[c].map(lambda v: lookup.get(str(v).strip().casefold(), pd.NA) if pd.notna(v) else pd.NA)
    for c in NUM_FEATURES:
        source = frame[c]
        converted = pd.to_numeric(source, errors="coerce")
        invalid = source.notna() & converted.isna()
        if invalid.any():
            raise ValueError(f"Kolom ‘{CODE_TO_LABEL[c]}’ berisi nilai yang bukan angka. Periksa format angka dan pemisah desimal.")
        outside = converted.notna() & ~converted.between(0, 100)
        if outside.any():
            raise ValueError(f"Nilai pada kolom ‘{CODE_TO_LABEL[c]}’ harus berada pada rentang 0–100.")
        frame[c] = converted
    frame[ID_COLUMN] = frame[ID_COLUMN].astype("string").str.strip()
    if frame[ID_COLUMN].isna().any() or frame[ID_COLUMN].eq("").any():
        raise ValueError("Kolom ‘Nama sekolah’ harus terisi untuk setiap peserta.")
    frame["id_peserta"] = frame["id_peserta"].astype("string").str.strip()
    if frame["id_peserta"].isna().any() or frame["id_peserta"].eq("").any():
        raise ValueError("Kolom ‘Nomor Induk Peserta’ harus terisi untuk setiap peserta.")
    entirely_missing = frame[FEATURES].isna().all(axis=1)
    if entirely_missing.any():
        raise ValueError(f"Ada {int(entirely_missing.sum())} baris tanpa satu pun prediktor terisi. Isi data peserta sebelum memproses.")
    return frame, filename


def prediction_payload(frame, filename):
    missing_cells = int(frame[FEATURES].isna().sum().sum())
    total_cells = int(len(frame) * len(FEATURES))
    row_missing_pct = frame[FEATURES].isna().mean(axis=1) * 100
    high_missing_rows = int(row_missing_pct.gt(30).sum())
    preds = predict(frame).clip(0, 100)
    out = pd.DataFrame({
        "id_peserta": frame["id_peserta"].astype(str) if "id_peserta" in frame else [f"Peserta {i+1}" for i in range(len(frame))],
        "jenis_sek": frame["jenis_sek"].astype("string").fillna("Tidak diisi") if "jenis_sek" in frame else "Tidak diisi",
        "estimasi_LIT": preds["LIT"].round(2), "estimasi_NUM": preds["NUM"].round(2),
    })
    friendly_csv = out.rename(columns={
        "id_peserta": "Nomor Induk Peserta", "jenis_sek": "Jenis sekolah",
        "estimasi_LIT": "Perkiraan skor literasi", "estimasi_NUM": "Perkiraan skor numerasi",
    }).to_csv(index=False)
    return {"filename": filename, "count": len(out), "lit_mean": round(float(preds.LIT.mean()), 2),
            "num_mean": round(float(preds.NUM.mean()), 2), "lit_median": round(float(preds.LIT.median()), 2),
            "num_median": round(float(preds.NUM.median()), 2), "rows": out.to_dict(orient="records"),
            "missing_cells": missing_cells, "missing_pct": round(100 * missing_cells / max(total_cells, 1), 2),
            "high_missing_rows": high_missing_rows, "csv": friendly_csv}


@app.post("/api/predict")
def api_predict():
    try:
        if "file" not in request.files:
            return jsonify({"error": "Pilih file CSV atau Excel terlebih dahulu."}), 400
        frame, filename = read_upload(request.files["file"])
        return jsonify(prediction_payload(frame, filename))
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/api/predict-form")
def api_predict_form():
    try:
        body = request.get_json(force=True)
        rows = body.get("rows", [])
        if not rows:
            return jsonify({"error": "Tambahkan minimal satu peserta ke daftar terlebih dahulu."}), 400
        row_limit = 1_000 if APP_ENV == "public" else MAX_ROWS
        if len(rows) > row_limit:
            return jsonify({"error": f"Maksimal {row_limit:,} peserta per proses."}), 400
        csv_content = pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")

        class FormUpload:
            filename = "form.csv"
            def read(self):
                return csv_content

        frame, _ = read_upload(FormUpload())
        return jsonify(prediction_payload(frame, "Input formulir"))
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400


@app.post("/api/save-run")
def save_run():
    if APP_ENV == "public":
        return jsonify({"error": "Riwayat publik disimpan di browser pengguna."}), 410
    body = request.get_json(force=True)
    required = ["school_name", "period", "n_participants", "lit_mean", "num_mean", "lit_median", "num_median", "filename"]
    if any(k not in body for k in required):
        return jsonify({"error": "Ringkasan proses belum lengkap."}), 400
    with db() as con:
        cur = con.execute("""INSERT INTO prediction_runs
            (created_at, school_name, period, n_participants, lit_mean, num_mean, lit_median, num_median, filename)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (datetime.now().isoformat(timespec="minutes"), body["school_name"], body["period"],
             int(body["n_participants"]), float(body["lit_mean"]), float(body["num_mean"]),
             float(body["lit_median"]), float(body["num_median"]), Path(body["filename"]).name))
        run_id = cur.lastrowid
    return jsonify({"ok": True, "id": run_id})


@app.get("/api/history")
def history():
    if APP_ENV == "public":
        return jsonify({"runs": [], "plans": []})
    with db() as con:
        runs = [dict(r) for r in con.execute("SELECT * FROM prediction_runs ORDER BY id DESC LIMIT 50")]
        plans = [dict(r) for r in con.execute("SELECT * FROM action_plans ORDER BY id DESC LIMIT 50")]
    return jsonify({"runs": runs, "plans": plans})


@app.post("/api/plans")
def add_plan():
    if APP_ENV == "public":
        return jsonify({"error": "Rencana publik disimpan di browser pengguna."}), 410
    body = request.get_json(force=True)
    fields = ["school_name", "focus", "action", "owner", "due_date"]
    if any(not str(body.get(k, "")).strip() for k in fields):
        return jsonify({"error": "Lengkapi semua kolom rencana."}), 400
    with db() as con:
        cur = con.execute("""INSERT INTO action_plans
            (created_at, school_name, focus, action, owner, due_date) VALUES (?, ?, ?, ?, ?, ?)""",
            (datetime.now().isoformat(timespec="minutes"), *(str(body[k]).strip() for k in fields)))
    return jsonify({"ok": True, "id": cur.lastrowid})


@app.post("/api/plans/<int:plan_id>/toggle")
def toggle_plan(plan_id):
    if APP_ENV == "public":
        return jsonify({"error": "Rencana publik disimpan di browser pengguna."}), 410
    with db() as con:
        row = con.execute("SELECT status FROM action_plans WHERE id=?", (plan_id,)).fetchone()
        if not row:
            return jsonify({"error": "Rencana tidak ditemukan."}), 404
        status = "Selesai" if row["status"] != "Selesai" else "Direncanakan"
        con.execute("UPDATE action_plans SET status=? WHERE id=?", (status, plan_id))
    return jsonify({"ok": True, "status": status})


@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": "Ukuran file maksimal 16 MB."}), 413


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5055, debug=False)
