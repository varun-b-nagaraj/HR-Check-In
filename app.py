import base64
import io
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import (
    Flask, jsonify, render_template, request,
    redirect, url_for, session, send_from_directory
)
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
from zoneinfo import ZoneInfo

APP_DIR = Path(__file__).parent.resolve()
ROSTER_PATH = APP_DIR / "students.xlsx"   # <-- your roster
PHOTOS_ROOT = APP_DIR / "photos"          # we now KEEP a copy for history page
PHOTOS_ROOT.mkdir(exist_ok=True)

# ---- Admin password + session secret
ADMIN_PASSWORD = "abhiMora1!"             # <--- you set this
SECRET_KEY = os.getenv("FLASK_SECRET_KEY", "change-me-please")
# -------------------------------------

app = Flask(__name__)
app.secret_key = SECRET_KEY


def load_roster():
    if not ROSTER_PATH.exists():
        raise FileNotFoundError("students.xlsx not found next to app.py")
    df = pd.read_excel(ROSTER_PATH)
    cols = {c.lower().strip(): c for c in df.columns}
    name_col = cols.get("name") or cols.get("student") or "Name"
    snum_col = cols.get("s-number") or cols.get("s number") or "s-number"
    df = df.rename(columns={name_col: "Name", snum_col: "s-number"})
    df["s-number"] = df["s-number"].astype(str).str.strip()
    df["Name"] = df["Name"].astype(str).str.strip()
    return df[["Name", "s-number"]]


def get_today_str():
    # Explicit US Central (handles CST/CDT correctly)
    return datetime.now(ZoneInfo("America/Chicago")).strftime("%Y-%m-%d")


def get_today_xlsx_path():
    return APP_DIR / f"attendance_{get_today_str()}.xlsx"


def ensure_workbook(path: Path):
    if path.exists():
        wb = load_workbook(path)
        ws = wb.active
        return wb, ws
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    # Added a 5th column PhotoPath for the history page (Photo is still embedded)
    ws.append(["S-Number", "Name", "Timestamp", "Photo", "PhotoPath"])
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 36
    wb.save(path)
    return wb, ws


def already_checked_in(ws, s_number: str):
    for row in ws.iter_rows(min_row=2, values_only=True):
        if str(row[0]).strip() == s_number:
            return True
    return False


def first_name(full_name: str) -> str:
    return (full_name or "").split()[0] if full_name else ""


def save_and_embed_photo(ws, row_idx: int, data_url: str, s_number: str):
    # Decode data URL to PIL image
    header, b64data = data_url.split(",", 1)
    binary = base64.b64decode(b64data)
    pil_img = Image.open(io.BytesIO(binary))

    # Resize to keep Excel light
    max_w, max_h = 240, 180
    pil_img.thumbnail((max_w, max_h))

    # Keep a copy on disk for the History page
    day_dir = PHOTOS_ROOT / get_today_str()
    day_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now(ZoneInfo("America/Chicago")).strftime("%H%M%S")
    filename = f"{s_number}_{ts}.png"
    file_path = day_dir / filename
    pil_img.save(file_path, format="PNG")

    # Embed into Excel (column D)
    xl_img = XLImage(str(file_path))
    cell_addr = f"D{row_idx}"
    ws.add_image(xl_img, cell_addr)
    ws.row_dimensions[row_idx].height = 140

    # Return relative web path for later display
    web_path = f"{get_today_str()}/{filename}"  # served via /photos/<path>
    return web_path


def calculate_analytics(roster, available_dates):
    """Calculate attendance analytics across all sessions"""
    total_sessions = len(available_dates)
    total_students = len(roster)
    
    # Track attendance for each student
    student_records = {}
    
    for _, student in roster.iterrows():
        s_num = str(student["s-number"])
        student_records[s_num] = {
            "name": student["Name"],
            "s_number": s_num,
            "present_count": 0,
            "absent_count": 0
        }
    
    # Count attendance across all dates
    for date in available_dates:
        xlsx_path = APP_DIR / f"attendance_{date}.xlsx"
        if not xlsx_path.exists():
            continue
            
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.active
        present_on_date = set()
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            s_num = str(row[0]).strip()
            present_on_date.add(s_num)
        
        # Update counts
        for s_num in student_records:
            if s_num in present_on_date:
                student_records[s_num]["present_count"] += 1
            else:
                student_records[s_num]["absent_count"] += 1
    
    # Calculate attendance rates
    students_list = []
    total_attendance = 0
    
    for s_num, record in student_records.items():
        if total_sessions > 0:
            attendance_rate = (record["present_count"] / total_sessions) * 100
        else:
            attendance_rate = 0
        
        students_list.append({
            "name": record["name"],
            "s_number": record["s_number"],
            "present_count": record["present_count"],
            "absent_count": record["absent_count"],
            "attendance_rate": attendance_rate
        })
        total_attendance += attendance_rate
    
    # Calculate average attendance
    avg_attendance = total_attendance / total_students if total_students > 0 else 0
    
    # Sort by name
    students_list.sort(key=lambda x: x["name"])
    
    return {
        "total_sessions": total_sessions,
        "total_students": total_students,
        "avg_attendance": avg_attendance,
        "students": students_list
    }


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/checkin", methods=["POST"])
def checkin():
    payload = request.get_json(force=True)
    s_number = str(payload.get("s_number", "")).strip()
    photo_data_url = payload.get("image_data_url")

    if not s_number:
        return jsonify({"ok": False, "error": "Missing s-number"}), 400
    if not photo_data_url or not photo_data_url.startswith("data:image/"):
        return jsonify({"ok": False, "error": "Missing or invalid photo"}), 400

    # Lookup roster
    roster = load_roster()
    match = roster.loc[roster["s-number"] == s_number]
    if match.empty:
        return jsonify({"ok": False, "error": "S-number not found"}), 404

    full_name = match.iloc[0]["Name"]
    fname = first_name(full_name)

    # Prepare workbook
    xlsx_path = get_today_xlsx_path()
    wb, ws = ensure_workbook(xlsx_path)

    # Only first check-in counts today
    if already_checked_in(ws, s_number):
        return jsonify({"ok": True, "status": "already", "first_name": fname})

    # Central time timestamp
    now = datetime.now(ZoneInfo("America/Chicago"))
    timestamp = now.strftime("%Y-%m-%d %H:%M:%S %Z")

    # Append row first (Photo + PhotoPath filled after)
    ws.append([s_number, full_name, timestamp, "", ""])
    row_idx = ws.max_row

    # Embed photo, keep file, and write PhotoPath (col E)
    web_path = save_and_embed_photo(ws, row_idx, photo_data_url, s_number)
    ws.cell(row=row_idx, column=5, value=web_path)

    wb.save(xlsx_path)

    return jsonify({"ok": True, "status": "new", "first_name": fname})


# ------------------ HISTORY (password-gated) ------------------

def is_authed():
    return session.get("authed") is True


@app.route("/history", methods=["GET", "POST"])
def history():
    # Password prompt / validation
    if request.method == "POST":
        pwd = request.form.get("password", "")
        if pwd == ADMIN_PASSWORD:
            session["authed"] = True
            return redirect(url_for("history"))
        return render_template("history_login.html", error="Incorrect password")

    if not is_authed():
        return render_template("history_login.html")

    # Get selected date from query param, default to today
    selected_date = request.args.get("date", get_today_str())
    
    # Get all available dates (all attendance files)
    available_dates = []
    for file in sorted(APP_DIR.glob("attendance_*.xlsx"), reverse=True):
        date_str = file.stem.replace("attendance_", "")
        available_dates.append(date_str)
    
    if not available_dates:
        available_dates = [get_today_str()]

    # Load roster
    roster = load_roster()
    roster["s-number"] = roster["s-number"].astype(str)

    # Get data for selected date
    xlsx_path = APP_DIR / f"attendance_{selected_date}.xlsx"
    present = []
    present_ids = set()

    if xlsx_path.exists():
        wb = load_workbook(xlsx_path, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            s_num = str(row[0]).strip()
            present_ids.add(s_num)
            # Safe read for optional PhotoPath column
            photo_path = ""
            if len(row) > 4 and row[4]:
                photo_path = row[4]
            present.append({
                "s_number": s_num,
                "name": row[1],
                "timestamp": row[2],
                "photo_path": photo_path,
            })

    # Absent = roster - present_ids for selected date
    absent_df = roster[~roster["s-number"].isin(present_ids)].copy()
    absent_df = absent_df.sort_values(by="Name")

    # Calculate analytics across all dates
    analytics = calculate_analytics(roster, available_dates)

    return render_template("history.html",
                           present=present,
                           absent=absent_df.to_dict(orient="records"),
                           selected_date=selected_date,
                           today=get_today_str(),
                           available_dates=available_dates,
                           analytics=analytics)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("history"))


# Serve saved photos (e.g., /photos/2025-10-28/151579_231405.png)
@app.route("/photos/<path:filename>")
def serve_photo(filename):
    return send_from_directory(PHOTOS_ROOT, filename, as_attachment=False)


if __name__ == "__main__":
    # For Chromebook local testing
    app.run(host="0.0.0.0", port=5005, debug=True)