import base64
import io
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template, request
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image
from zoneinfo import ZoneInfo


APP_DIR = Path(__file__).parent.resolve()
ROSTER_PATH = APP_DIR / "students.xlsx"  # <-- your roster
PHOTOS_DIR = APP_DIR / "photos"          # temp save before embedding
PHOTOS_DIR.mkdir(exist_ok=True)

app = Flask(__name__)

def load_roster():
    if not ROSTER_PATH.exists():
        raise FileNotFoundError("students.xlsx not found next to app.py")
    df = pd.read_excel(ROSTER_PATH)
    # Normalize column names
    cols = {c.lower().strip(): c for c in df.columns}
    name_col = cols.get("name") or cols.get("student") or "Name"
    snum_col = cols.get("s-number") or cols.get("s number") or "s-number"
    # Defensive rename
    df = df.rename(columns={name_col: "Name", snum_col: "s-number"})
    df["s-number"] = df["s-number"].astype(str).str.strip()
    df["Name"] = df["Name"].astype(str).str.strip()
    return df[["Name", "s-number"]]

def get_today_xlsx_path():
    today_str = datetime.now().strftime("%Y-%m-%d")
    return APP_DIR / f"attendance_{today_str}.xlsx"

def ensure_workbook(path: Path):
    if path.exists():
        wb = load_workbook(path)
        ws = wb.active
        return wb, ws
    # create new
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    ws.append(["S-Number", "Name", "Timestamp", "Photo"])
    # Some friendly column widths
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 18
    wb.save(path)
    return wb, ws

def already_checked_in(ws, s_number: str):
    # scan col A for existing s-number (skip header)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[0]).strip() == s_number:
            return True
    return False

def first_name(full_name: str) -> str:
    return (full_name or "").split()[0] if full_name else ""

def save_and_embed_photo(ws, row_idx: int, data_url: str, s_number: str):
    # data_url is "data:image/png;base64,...."
    header, b64data = data_url.split(",", 1)
    # Decode to PIL image
    binary = base64.b64decode(b64data)
    pil_img = Image.open(io.BytesIO(binary))

    # (Optional) resize to keep Excel light
    max_w, max_h = 240, 180
    pil_img.thumbnail((max_w, max_h))

    # Save temp PNG
    ts = datetime.now().strftime("%H%M%S")
    temp_name = f"{s_number}_{ts}.png"
    temp_path = PHOTOS_DIR / temp_name
    pil_img.save(temp_path, format="PNG")

    # Embed at column D, current row
    cell_addr = f"D{row_idx}"
    xl_img = XLImage(str(temp_path))
    ws.add_image(xl_img, cell_addr)

    # Optional: adjust row height (roughly 0.75 * pixel)
    ws.row_dimensions[row_idx].height = 140

    return temp_path  # caller may delete after workbook is saved

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/checkin", methods=["POST"])
def checkin():
    payload = request.get_json(force=True)
    s_number = str(payload.get("s_number", "")).strip()
    photo_data_url = payload.get("image_data_url")  # base64 data URL

    if not s_number:
        return jsonify({"ok": False, "error": "Missing s-number"}), 400
    if not photo_data_url or not photo_data_url.startswith("data:image/"):
        return jsonify({"ok": False, "error": "Missing or invalid photo"}), 400

    # Lookup in roster
    roster = load_roster()
    match = roster.loc[roster["s-number"] == s_number]
    if match.empty:
        return jsonify({"ok": False, "error": "S-number not found"}), 404

    full_name = match.iloc[0]["Name"]
    fname = first_name(full_name)

    # Prepare workbook for today
    xlsx_path = get_today_xlsx_path()
    wb, ws = ensure_workbook(xlsx_path)

    # Only first check-in counts
    if already_checked_in(ws, s_number):
        return jsonify({"ok": True, "status": "already", "first_name": fname})

    # Append row
    now = datetime.now(ZoneInfo("America/Chicago"))
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S %Z")  # include CST/CDT abbreviation
    ws.append([s_number, full_name, timestamp, ""])

    # New row index is last row
    row_idx = ws.max_row

    # Embed photo in column D
    temp_path = save_and_embed_photo(ws, row_idx, photo_data_url, s_number)

    # Save workbook (embed stores the image inside .xlsx)
    wb.save(xlsx_path)

    # Clean up temp
    try:
        os.remove(temp_path)
    except Exception:
        pass

    return jsonify({"ok": True, "status": "new", "first_name": fname})

if __name__ == "__main__":
    # For Chromebook local testing; camera works on http://localhost
    app.run(host="0.0.0.0", port=5005, debug=True)
