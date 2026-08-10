from flask import Flask, request, jsonify, send_file, render_template
import os
import threading
import time
import datetime
import json
import sqlite3
import shutil
import socket
import smtplib
import importlib.util
import sys
import subprocess
from pathlib import Path
from docx import Document
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from collections import defaultdict

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max

# ==========================================
# CONFIGURATION
# ==========================================
ACCESS_CODE = os.environ.get("ACCESS_CODE", "user2024")
ADMIN_CODE  = os.environ.get("ADMIN_CODE", "admin2024")
GMAIL_USER  = os.environ.get("GMAIL_USER", "your_email@gmail.com")
GMAIL_PASS  = os.environ.get("GMAIL_PASS", "your_app_password")

# ==========================================
# DATA DIRECTORIES
# ==========================================
DATA_DIR    = Path("data")
DATA_DIR.mkdir(exist_ok=True)
EXCELS_DIR  = DATA_DIR / "excels"
EXCELS_DIR.mkdir(exist_ok=True)
LOGS_DB     = DATA_DIR / "logs.db"

# ==========================================
# DATABASE INITIALIZATION
# ==========================================
def init_db():
    conn = sqlite3.connect(str(LOGS_DB))
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS bewerber_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_name TEXT,
            event_type TEXT,
            company_num INTEGER,
            email TEXT,
            firma TEXT,
            status TEXT,
            error_msg TEXT,
            files_sent TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS excel_uploads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_name TEXT,
            filename TEXT,
            storage_path TEXT,
            uploaded_at TEXT DEFAULT (datetime('now'))
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def get_db():
    conn = sqlite3.connect(str(LOGS_DB))
    conn.row_factory = sqlite3.Row
    return conn

# ==========================================
# GLOBAL STATE (thread-safe)
# ==========================================
state = {
    "logged_in": False,
    "is_admin": False,
    "generating": False,
    "sending": False,
    "generated": False,
    "base_dir": None,
    "other": None,
    "extra": None,
    "scheduled_dt": None,
    "send_done": False,
    "interrupted_at": None,
    "gen_progress": 0,
    "gen_log": [],
    "gen_total": 0,
    "send_progress": 0,
    "send_log": [],
    "send_total": 0,
    "bewerbungsname": None,
    "total_companies": 0,
    "anschreiben_pos": 2,
    "delay": 10,
    "start": 1,
    "waiting_scheduled": False,
    "network_error": False,
}

state_lock = threading.Lock()

# ==========================================
# HELPER FUNCTIONS
# ==========================================
def convert(docx_path, pdf_path):
    subprocess.run([
        "libreoffice", "--headless", "--convert-to", "pdf",
        "--outdir", str(Path(pdf_path).parent), str(docx_path)
    ], check=True)

def is_network_error(e):
    err = str(e).lower()
    keywords = ["connection", "network", "timeout", "socket", "refused",
                "unreachable", "errno", "broken pipe", "reset", "ssl", "eof", "timed out"]
    return any(k in err for k in keywords) or isinstance(e, (socket.timeout, socket.gaierror, OSError))

def reset_state():
    with state_lock:
        keys_to_reset = [
            "generating", "sending", "generated",
            "base_dir", "other", "extra", "scheduled_dt",
            "send_done", "interrupted_at",
            "gen_progress", "gen_log", "gen_total",
            "send_progress", "send_log", "send_total",
            "bewerbungsname", "total_companies",
            "anschreiben_pos", "delay", "start",
            "waiting_scheduled", "network_error",
        ]
        for key in keys_to_reset:
            if key in state:
                if isinstance(state[key], list):
                    state[key] = []
                elif isinstance(state[key], dict):
                    state[key] = {}
                elif isinstance(state[key], int) and key in ("gen_progress", "send_progress", "gen_total", "send_total", "total_companies", "start", "anschreiben_pos", "delay"):
                    state[key] = 0 if key not in ("start", "anschreiben_pos", "delay") else (1 if key == "start" else (2 if key == "anschreiben_pos" else 10))
                else:
                    state[key] = None
        state["logged_in"] = True

def log_event(session_name, event_type, company_num=None, email=None,
              firma=None, status=None, error_msg=None, files_sent=None):
    try:
        conn = get_db()
        conn.execute(
            "INSERT INTO bewerber_logs (session_name, event_type, company_num, email, firma, status, error_msg, files_sent) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (session_name, event_type, company_num, email, firma, status,
             error_msg or "", json.dumps(files_sent or []))
        )
        conn.commit()
        conn.close()
    except Exception:
        pass

def save_excel_to_db(file_path, filename, session_name):
    try:
        timestamp = datetime.datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        storage_filename = f"{timestamp}_{filename}"
        storage_path = EXCELS_DIR / storage_filename
        shutil.copy2(str(file_path), str(storage_path))

        conn = get_db()
        conn.execute(
            "INSERT INTO excel_uploads (session_name, filename, storage_path) VALUES (?, ?, ?)",
            (session_name, filename, str(storage_path))
        )
        conn.commit()
        conn.close()
    except Exception:
        pass

def clean_old_excels():
    try:
        conn = get_db()
        yesterday = (datetime.datetime.utcnow() - datetime.timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
        old = conn.execute("SELECT * FROM excel_uploads WHERE uploaded_at < ?", (yesterday,)).fetchall()
        for ex in old:
            try:
                os.remove(ex["storage_path"])
            except Exception:
                pass
            conn.execute("DELETE FROM excel_uploads WHERE id = ?", (ex["id"],))
        conn.commit()
        conn.close()
    except Exception:
        pass

def save_upload(file_storage, target_dir):
    filename = file_storage.filename
    path = target_dir / filename
    file_storage.save(str(path))
    return path

def save_other_file(file_storage, target_dir):
    if file_storage and file_storage.filename:
        path = target_dir / file_storage.filename
        file_storage.save(str(path))
        state["other"] = str(path)

def save_extra_file(file_storage, target_dir):
    if file_storage and file_storage.filename:
        name = file_storage.filename
        if not name or name.strip() == "" or "." not in name:
            name = "Anhang.pdf"
        path = target_dir / name
        file_storage.save(str(path))
        state["extra"] = str(path)
    else:
        state["extra"] = None

def save_old_values(salutation, person_full, email, adresse_3, output_dir):
    content = f'''
old_salutation = "{salutation}"
old_person = "{person_full}"
old_email = "{email}"
old_adresse_3 = "{adresse_3}"'''
    with open(output_dir / "saved_values.py", "w", encoding="utf-8") as f:
        f.write(content)

def generate_letter(template_path, row, output_dir):
    doc = Document(str(template_path))
    person_full = str(row["person"])
    firma = str(row["firma"])
    adresse = str(row["adresse"])
    email = str(row["email"])

    if person_full.startswith("Herr"):
        salutation = "er "
        gender_def = True
    elif person_full.startswith("Frau"):
        salutation = "e "
        gender_def = True
    else:
        salutation = "e "
        gender_def = False

    if "|" in adresse:
        adresse_1, adresse_2 = [x.strip() for x in adresse.split("|", 1)]
    else:
        adresse_1 = adresse
        adresse_2 = adresse
    adresse_3 = adresse_2[6:]

    for p in doc.paragraphs:
        full = "".join(r.text for r in p.runs)
        full = full.replace("{#custom}", firma)
        full = full.replace("{gender}", salutation)
        full = full.replace("{zeit}", datetime.date.today().strftime("%d.%m.%Y"))

        if gender_def:
            full = full.replace("{person}", person_full)
            full = full.replace("{.}", person_full)
            full = full.replace("{/custom}", adresse_1)
            full = full.replace("{/custom2}", adresse_2)
            full = full.replace("{adre}", adresse_3)
        else:
            full = full.replace("{person}", "Damen und Herren")
            full = full.replace("{.}", adresse_1)
            full = full.replace("{/custom}", adresse_2)
            full = full.replace("{/custom2}", "")
            full = full.replace("{adre}", adresse_3)

        for r in p.runs:
            r.text = ""
        if len(p.runs) == 0:
            p.add_run(full)
        else:
            p.runs[0].text = full

    safe_name = firma.replace("/", "_").replace("\\", "_")
    docx_path = output_dir / f"{safe_name}.docx"
    pdf_path = output_dir / f"{safe_name}.pdf"
    doc.save(str(docx_path))
    convert(docx_path, pdf_path)
    os.remove(str(docx_path))
    save_old_values(salutation, person_full, email, adresse_3, output_dir)
    return pdf_path

def merge_pdfs(cv_path, cover_path, position, output_path):
    main_reader = PdfReader(str(cv_path))
    insert_reader = PdfReader(str(cover_path))
    writer = PdfWriter()
    for i in range(len(main_reader.pages)):
        if i == position - 1:
            writer.add_page(insert_reader.pages[0])
        writer.add_page(main_reader.pages[i])
    if position > len(main_reader.pages):
        writer.add_page(insert_reader.pages[0])
    with open(output_path, "wb") as f:
        writer.write(f)

def safe_path(p):
    if p and isinstance(p, str):
        return Path(p)
    return p

def load_module(cmp, saved_path):
    module_name = f"vals_{cmp}"
    spec = importlib.util.spec_from_file_location(module_name, str(saved_path))
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def gmail_send(to_email, subject, body, file1_path=None, file2_path=None, file3_path=None):
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    message = MIMEMultipart()
    message["From"] = GMAIL_USER
    message["To"] = to_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))
    for fpath in [safe_path(file1_path), safe_path(file2_path), safe_path(file3_path)]:
        if fpath:
            with open(fpath, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{fpath.name}"')
            message.attach(part)
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(GMAIL_USER, GMAIL_PASS)
    server.sendmail(GMAIL_USER, to_email, message.as_string())
    server.quit()

# ==========================================
# BACKGROUND THREAD: GENERATE
# ==========================================
def generate_thread(excel_path, cv_path, template_path, other_file_storage,
                    anschreiben_pos, bewerbungsname):
    with state_lock:
        state["generating"] = True
        state["generated"] = False
        state["gen_log"] = []
        state["gen_progress"] = 0

    base_dir = Path(bewerbungsname)
    if base_dir.exists() and base_dir.is_dir():
        shutil.rmtree(str(base_dir))
    base_dir.mkdir(exist_ok=True)

    excel_path_saved = save_upload(excel_path, base_dir)
    cv_path_saved = save_upload(cv_path, base_dir)
    template_path_saved = save_upload(template_path, base_dir)
    save_other_file(other_file_storage, base_dir)

    save_excel_to_db(str(excel_path_saved), excel_path.filename, bewerbungsname)

    df = pd.read_excel(str(excel_path_saved))
    total = len(df)

    with state_lock:
        state["gen_total"] = total
        state["base_dir"] = str(base_dir)
        state["bewerbungsname"] = bewerbungsname
        state["total_companies"] = total
        state["anschreiben_pos"] = anschreiben_pos

    for i, row in df.iterrows():
        out_dir = base_dir / str(i + 1)
        out_dir.mkdir(exist_ok=True)
        firma = str(row.get("firma", ""))

        try:
            cover_pdf = generate_letter(template_path_saved, row, out_dir)
            final_path = out_dir / f"{bewerbungsname}.pdf"
            merge_pdfs(cv_path_saved, cover_pdf, anschreiben_pos, str(final_path))
            log_event(bewerbungsname, "generated", company_num=i+1, firma=firma, status="ok")
            os.remove(str(cover_pdf))
            entry = {"num": i+1, "firma": firma, "status": "ok"}
        except Exception:
            log_event(bewerbungsname, "generated", company_num=i+1, firma=firma, status="error")
            with open(out_dir / "skip.txt", "w") as f:
                f.write("skip")
            entry = {"num": i+1, "firma": firma, "status": "error"}

        with state_lock:
            state["gen_log"].append(entry)
            state["gen_progress"] = (i + 1) / total

    with state_lock:
        state["generating"] = False
        state["generated"] = True

# ==========================================
# BACKGROUND THREAD: SEND
# ==========================================
def send_thread(letter_path, delay, start_num, scheduled_dt, extra_file_storage):
    with state_lock:
        state["sending"] = True
        state["send_done"] = False
        state["interrupted_at"] = None
        state["send_log"] = []
        state["send_progress"] = 0
        state["start"] = start_num
        state["scheduled_dt"] = scheduled_dt
        state["network_error"] = False
        base_dir = Path(state["base_dir"])
        bewerbungsname = state["bewerbungsname"]
        total = state["total_companies"]

    # Save email template
    letter_path_saved = save_upload(letter_path, base_dir)

    # Save extra file
    save_extra_file(extra_file_storage, base_dir)

    # Read email template
    doc = Document(str(letter_path_saved))
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    subject = lines[0]
    message_template = "\n".join(lines[1:])

    # Wait for scheduled time
    if scheduled_dt:
        with state_lock:
            state["waiting_scheduled"] = True

        while True:
            now = datetime.datetime.now()
            remaining = (scheduled_dt - now).total_seconds()
            if remaining <= 0:
                break
            time.sleep(min(remaining, 1))

        with state_lock:
            state["waiting_scheduled"] = False

    # Send emails
    for cmp in range(start_num, total + 1):
        cmp_dir = base_dir / str(cmp)
        saved_path = cmp_dir / "saved_values.py"
        pdf_path = cmp_dir / f"{bewerbungsname}.pdf"

        email = "???"
        try:
            m = load_module(cmp, saved_path)
            email = m.old_email
        except Exception:
            pass

        if (cmp_dir / "skip.txt").exists():
            log_event(bewerbungsname, "skip", company_num=cmp, email=email, status="skip")
            with state_lock:
                state["send_log"].append({"num": cmp, "email": email, "status": "skip"})
                state["send_progress"] = cmp / total
            continue

        try:
            m = load_module(cmp, saved_path)
            email = m.old_email
            gender = m.old_salutation
            person = m.old_person
            adresse_3 = m.old_adresse_3.strip()

            if person == "x":
                person = "Damen und Herren"

            letter = (
                message_template.replace("{person}", person)
                .replace("{gender}", gender)
                .replace("{adre}", adresse_3)
                .replace("{space}", "\n")
                .replace("{2space}", "\n\n")
            )

            gmail_send(
                email, subject, letter,
                str(pdf_path) if pdf_path.exists() else None,
                state.get("extra"),
                state.get("other"),
            )

            files_sent = [str(pdf_path.name)] if pdf_path.exists() else []
            if state.get("extra"):
                files_sent.append(Path(state["extra"]).name)
            if state.get("other"):
                files_sent.append(Path(state["other"]).name)

            log_event(bewerbungsname, "sent", company_num=cmp, email=email,
                      status="sent", files_sent=files_sent)

            with state_lock:
                state["send_log"].append({"num": cmp, "email": email, "status": "sent"})
                state["send_progress"] = cmp / total

            time.sleep(delay)

        except Exception as e:
            if is_network_error(e):
                log_event(bewerbungsname, "network_error", company_num=cmp,
                          email=email, status="network_error", error_msg=str(e))
                with state_lock:
                    state["send_log"].append({"num": cmp, "email": email, "status": "network_error"})
                    state["interrupted_at"] = cmp
                    state["sending"] = False
                    state["network_error"] = True
                return
            else:
                log_event(bewerbungsname, "error", company_num=cmp,
                          email=email, status="error", error_msg=str(e))
                with state_lock:
                    state["send_log"].append({"num": cmp, "email": email, "status": "error"})
                    state["send_progress"] = cmp / total

    with state_lock:
        state["sending"] = False
        state["scheduled_dt"] = None
        state["send_done"] = True

# ==========================================
# ROUTES
# ==========================================
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json
    code = data.get('code', '')
    if code == ADMIN_CODE:
        with state_lock:
            state["logged_in"] = True
            state["is_admin"] = True
        return jsonify({"success": True, "is_admin": True})
    elif code == ACCESS_CODE:
        with state_lock:
            state["logged_in"] = True
            state["is_admin"] = False
        return jsonify({"success": True, "is_admin": False})
    else:
        return jsonify({"success": False, "error": "Falscher Code"}), 401

@app.route('/api/state', methods=['GET'])
def get_state():
    with state_lock:
        return jsonify({
            "logged_in": state["logged_in"],
            "is_admin": state["is_admin"],
            "generating": state["generating"],
            "generated": state["generated"],
            "sending": state["sending"],
            "send_done": state["send_done"],
            "interrupted_at": state["interrupted_at"],
            "bewerbungsname": state["bewerbungsname"],
            "total_companies": state["total_companies"],
            "waiting_scheduled": state.get("waiting_scheduled", False),
            "scheduled_dt": state["scheduled_dt"].isoformat() if state.get("scheduled_dt") else None,
        })

@app.route('/api/generate', methods=['POST'])
def generate():
    if not state.get("logged_in"):
        return jsonify({"error": "Not logged in"}), 403

    excel_file = request.files.get('excel')
    cv_file = request.files.get('cv')
    template_file = request.files.get('template')
    other_file = request.files.get('other')
    anschreiben_pos = int(request.form.get('position', 2))
    bewerbungsname = request.form.get('bewerbungsname', 'Bewerbung')

    if not excel_file or not cv_file or not template_file or not other_file:
        return jsonify({"error": "Alle Dateien müssen hochgeladen werden"}), 400

    if state.get("generating"):
        return jsonify({"error": "Generierung läuft bereits"}), 400

    thread = threading.Thread(target=generate_thread, args=(
        excel_file, cv_file, template_file, other_file,
        anschreiben_pos, bewerbungsname
    ))
    thread.daemon = True
    thread.start()

    return jsonify({"success": True})

@app.route('/api/generate/status', methods=['GET'])
def generate_status():
    with state_lock:
        return jsonify({
            "generating": state["generating"],
            "generated": state["generated"],
            "progress": state["gen_progress"],
            "log": state["gen_log"],
            "total": state["gen_total"],
        })

@app.route('/api/send', methods=['POST'])
def send():
    if not state.get("logged_in"):
        return jsonify({"error": "Not logged in"}), 403

    if not state.get("generated"):
        return jsonify({"error": "Bitte zuerst Anschreiben generieren"}), 400

    if state.get("sending"):
        return jsonify({"error": "Senden läuft bereits"}), 400

    letter_file = request.files.get('letter')
    extra_file = request.files.get('extra')
    delay = int(request.form.get('delay', 10))
    start_num = int(request.form.get('start', 1))
    schedule_str = request.form.get('scheduled_dt', '')

    if not letter_file:
        return jsonify({"error": "Email Template erforderlich"}), 400

    scheduled_dt = None
    if schedule_str:
        try:
            scheduled_dt = datetime.datetime.fromisoformat(schedule_str)
            if scheduled_dt <= datetime.datetime.now():
                return jsonify({"error": "Die gewählte Zeit liegt in der Vergangenheit"}), 400
        except Exception:
            return jsonify({"error": "Ungültiges Datum/Zeit Format"}), 400

    thread = threading.Thread(target=send_thread, args=(
        letter_file, delay, start_num, scheduled_dt, extra_file
    ))
    thread.daemon = True
    thread.start()

    return jsonify({"success": True})

@app.route('/api/send/status', methods=['GET'])
def send_status():
    with state_lock:
        return jsonify({
            "sending": state["sending"],
            "send_done": state["send_done"],
            "interrupted_at": state["interrupted_at"],
            "progress": state["send_progress"],
            "log": state["send_log"],
            "total": state["total_companies"],
            "waiting_scheduled": state.get("waiting_scheduled", False),
            "scheduled_dt": state["scheduled_dt"].isoformat() if state.get("scheduled_dt") else None,
            "bewerbungsname": state["bewerbungsname"],
        })

@app.route('/api/send/resume', methods=['POST'])
def resume_send():
    if not state.get("logged_in"):
        return jsonify({"error": "Not logged in"}), 403

    data = request.json
    resume_from = int(data.get('resume_from', 1))

    with state_lock:
        state["interrupted_at"] = None
        state["network_error"] = False

    # Need to re-read the letter template - it was already saved
    base_dir = Path(state["base_dir"])
    # Find the letter template in base_dir
    letter_files = list(base_dir.glob("*.docx"))
    if not letter_files:
        return jsonify({"error": "Email Template nicht gefunden"}), 400

    letter_path = letter_files[0]

    thread = threading.Thread(target=send_thread, args=(
        type('F', (), {'filename': letter_path.name, 'save': lambda self, p: None})(),
        state["delay"], resume_from, None, None
    ))
    thread.daemon = True
    thread.start()

    return jsonify({"success": True})

@app.route('/api/reset', methods=['POST'])
def reset():
    reset_state()
    return jsonify({"success": True})

@app.route('/api/dashboard', methods=['GET'])
def dashboard():
    if not state.get("logged_in") or not state.get("is_admin"):
        return jsonify({"error": "Access denied"}), 403

    session_filter = request.args.get('session', 'Alle')
    hours_filter = request.args.get('hours', type=int)

    conn = get_db()

    query = "SELECT * FROM bewerber_logs"
    conditions = []
    params = []

    if hours_filter:
        cutoff = (datetime.datetime.utcnow() - datetime.timedelta(hours=hours_filter)).strftime("%Y-%m-%d %H:%M:%S")
        conditions.append("created_at >= ?")
        params.append(cutoff)

    if conditions:
        query += " WHERE " + " AND ".join(conditions)

    query += " ORDER BY created_at DESC LIMIT 1000"
    rows = conn.execute(query, params).fetchall()
    conn.close()

    # Stats
    sent_n = sum(1 for r in rows if r["event_type"] == "sent")
    skip_n = sum(1 for r in rows if r["event_type"] == "skip")
    error_n = sum(1 for r in rows if r["event_type"] in ("error", "network_error"))
    gen_ok = sum(1 for r in rows if r["event_type"] == "generated" and r["status"] == "ok")

    # Sessions list
    sessions = list(set(r["session_name"] for r in rows if r["session_name"]))

    # Group by company
    companies = defaultdict(lambda: {
        "session": "", "generiert": "—", "gesendet": "Nein",
        "email_firma": "", "fehler": "—", "zeit": ""
    })

    for r in sorted(rows, key=lambda x: x["created_at"]):
        key = (r["session_name"], r["company_num"])
        e = companies[key]
        e["session"] = r["session_name"] or ""

        if r["event_type"] == "generated":
            e["generiert"] = "Ja" if r["status"] == "ok" else "Nein"
            e["email_firma"] = r["firma"] or ""
            e["zeit"] = r["created_at"][:16] if r["created_at"] else ""

        if r["event_type"] == "sent":
            e["gesendet"] = "Ja"
            e["email_firma"] = r["email"] or e["email_firma"]
            e["zeit"] = r["created_at"][:16] if r["created_at"] else ""

        if r["event_type"] == "skip":
            e["gesendet"] = "Nein"

        if r["event_type"] in ("error", "network_error"):
            e["fehler"] = "Ja"

    # Filter by session
    if session_filter != "Alle":
        companies = {k: v for k, v in companies.items() if k[0] == session_filter}

    # Sort by company number
    company_list = []
    for (session, cmp_num), e in sorted(companies.items(), key=lambda x: (x[0][1] or 0)):
        company_list.append({
            "session": e["session"],
            "cmp_num": cmp_num,
            "generiert": e["generiert"],
            "gesendet": e["gesendet"],
            "email_firma": e["email_firma"],
            "fehler": e["fehler"],
            "zeit": e["zeit"],
        })

    return jsonify({
        "stats": {
            "generated": gen_ok,
            "sent": sent_n,
            "skipped": skip_n,
            "errors": error_n,
        },
        "sessions": sessions,
        "companies": company_list,
    })

@app.route('/api/dashboard/excels', methods=['GET'])
def dashboard_excels():
    if not state.get("logged_in") or not state.get("is_admin"):
        return jsonify({"error": "Access denied"}), 403

    clean_old_excels()

    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM excel_uploads ORDER BY uploaded_at DESC LIMIT 200"
    ).fetchall()
    conn.close()

    session_filter = request.args.get('session', 'Alle')

    files = []
    for r in rows:
        if session_filter != "Alle" and r["session_name"] != session_filter:
            continue
        files.append({
            "id": r["id"],
            "filename": r["filename"],
            "session_name": r["session_name"],
            "uploaded_at": r["uploaded_at"][:16] if r["uploaded_at"] else "",
        })

    all_sessions = list(set(r["session_name"] for r in rows if r["session_name"]))

    return jsonify({"files": files, "sessions": all_sessions})

@app.route('/api/dashboard/excel/download/<int:file_id>')
def download_excel(file_id):
    if not state.get("logged_in") or not state.get("is_admin"):
        return jsonify({"error": "Access denied"}), 403

    conn = get_db()
    row = conn.execute("SELECT * FROM excel_uploads WHERE id = ?", (file_id,)).fetchone()
    conn.close()

    if not row:
        return jsonify({"error": "File not found"}), 404

    return send_file(row["storage_path"], as_attachment=True, download_name=row["filename"])

# ==========================================
# MAIN
# ==========================================
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)