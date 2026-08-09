import os
import io
import time
import socket
import shutil
import base64
import datetime
import threading
import importlib.util
import sys
import subprocess
from pathlib import Path
from collections import defaultdict

from flask import Flask, request, jsonify, render_template, send_file, session
from werkzeug.utils import secure_filename

from docx import Document
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from supabase import create_client
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "super_secret_key_123")

# ==========================================
# CONFIG & SECRETS (Replace with your actual keys or .env)
# ==========================================
ACCESS_CODE = os.environ.get("ACCESS_CODE", "user123")
ADMIN_CODE = os.environ.get("ADMIN_CODE", "admin123")
SUPABASE_URL = os.environ.get("SUPABASE_URL", "your_supabase_url")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "your_supabase_key")
GMAIL_USER = os.environ.get("GMAIL_USER", "your_email@gmail.com")
GMAIL_PASS = os.environ.get("GMAIL_PASS", "your_app_password")

# ==========================================
# GLOBAL APP STATE (Replaces Streamlit Session State)
# ==========================================
app_state = {
    "logged_in": False,
    "is_admin": False,
    "page": "anschreiben",
    "generating": False,
    "sending": False,
    "generated": False,
    "base_dir": None,
    "other": None,
    "extra": None,
    "scheduled_dt": None,
    "send_done": False,
    "interrupted_at": None,
    "progress": 0,
    "logs": [],
    "total_companies": 0,
    "bewerbungsname": "",
    "anschreiben_pos": 2,
    "delay": 10,
    "start": 1,
    "countdown_msg": ""
}

# ==========================================
# HELPERS (Unchanged logic)
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

def reset_all():
    keys_to_reset = ["generating", "sending", "generated", "base_dir", "other", "extra", 
                     "scheduled_dt", "total_companies", "send_done", "interrupted_at", "progress", "logs"]
    for key in keys_to_reset:
        app_state[key] = None if key not in ["progress", "logs", "total_companies"] else (0 if key == "progress" else ([] if key == "logs" else 0))
    
    app_state["logged_in"] = True
    app_state["page"] = "anschreiben"
    app_state["generated"] = False
    app_state["send_done"] = False
    app_state["interrupted_at"] = None

def get_supabase():
    return create_client(SUPABASE_URL, SUPABASE_KEY)

def save_excel_to_db(file_bytes, filename, session_name):
    try:
        sb = get_supabase()
        timestamp = datetime.datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        storage_path = f"excels/{timestamp}_{filename}"
        sb.storage.from_("excel-files").upload(
            storage_path, file_bytes,
            {"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
        )
        sb.table("excel_uploads").insert({
            "session_name": session_name, "filename": filename, "storage_path": storage_path,
        }).execute()
    except Exception:
        pass

def log_event(session_name, event_type, company_num=None, email=None, firma=None, status=None, error_msg=None, files_sent=None):
    try:
        get_supabase().table("bewerber_logs").insert({
            "session_name": session_name, "event_type": event_type, "company_num": company_num,
            "email": email, "firma": firma, "status": status, "error_msg": error_msg or "",
            "files_sent": files_sent or [],
        }).execute()
    except Exception:
        pass

def save_file(file_storage, target_dir):
    filename = secure_filename(file_storage.filename)
    path = target_dir / filename
    file_storage.save(path)
    return path

def save_old_values(salutation, person_full, email, adresse_3, output_dir):
    content = f'''
old_salutation = "{salutation}"
old_person = "{person_full}"
old_email = "{email}"
old_adresse_3 = "{adresse_3}"'''
    with open(output_dir / "saved_values.py", "w", encoding="utf-8") as f:
        f.write(content)

def generate_letter(template_path, row, output_dir):
    doc = Document(template_path)
    person_full = row["person"]
    firma = row["firma"]
    adresse = row["adresse"]
    email = row["email"]

    if person_full.startswith("Herr"): salutation, gender_def = "er ", True
    elif person_full.startswith("Frau"): salutation, gender_def = "e ", True
    else: salutation, gender_def = "e ", False

    if "|" in adresse: adresse_1, adresse_2 = [x.strip() for x in adresse.split("|", 1)]
    else: adresse_1, adresse_2 = adresse, adresse
    adresse_3 = adresse_2[6:]

    for p in doc.paragraphs:
        full = "".join(r.text for r in p.runs)
        full = full.replace("{#custom}", firma).replace("{gender}", salutation).replace("{zeit}", datetime.date.today().strftime("%d.%m.%Y"))
        if gender_def:
            full = full.replace("{person}", person_full).replace("{.}", person_full).replace("{/custom}", adresse_1).replace("{/custom2}", adresse_2).replace("{adre}", adresse_3)
        else:
            full = full.replace("{person}", "Damen und Herren").replace("{.}", adresse_1).replace("{/custom}", adresse_2).replace("{/custom2}", "").replace("{adre}", adresse_3)
        for r in p.runs: r.text = ""
        if len(p.runs) == 0: p.add_run(full)
        else: p.runs[0].text = full

    safe_name = firma.replace("/", "_").replace("\\", "_")
    docx_path = output_dir / f"{safe_name}.docx"
    pdf_path = output_dir / f"{safe_name}.pdf"
    doc.save(docx_path)
    convert(docx_path, pdf_path)
    os.remove(docx_path)
    save_old_values(salutation, person_full, email, adresse_3, output_dir)
    return pdf_path

def merge_pdfs(cv_path, cover_path, position, output_path):
    main_reader = PdfReader(cv_path)
    insert_reader = PdfReader(cover_path)
    writer = PdfWriter()
    for i in range(len(main_reader.pages)):
        if i == position - 1: writer.add_page(insert_reader.pages[0])
        writer.add_page(main_reader.pages[i])
    if position > len(main_reader.pages): writer.add_page(insert_reader.pages[0])
    with open(output_path, "wb") as f: writer.write(f)

def safe_path(p):
    if p and isinstance(p, str): return Path(p)
    return p

def load_module(cmp, saved_path):
    module_name = f"vals_{cmp}"
    spec = importlib.util.spec_from_file_location(module_name, saved_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module

def gmail_send(creds_unused, to_email, subject, body, file1_path=None, file2_path=None, file3_path=None):
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
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(GMAIL_USER, GMAIL_PASS)
    server.sendmail(GMAIL_USER, to_email, message.as_string())
    server.quit()


# ==========================================
# BACKGROUND TASKS (Threads)
# ==========================================
def generation_thread(files, pos, bewerbungsname):
    app_state['generating'] = True
    app_state['logs'] = []
    app_state['progress'] = 0
    
    base_dir = Path(bewerbungsname)
    if base_dir.exists() and base_dir.is_dir(): shutil.rmtree(base_dir)
    base_dir.mkdir(exist_ok=True)

    excel_path = save_file(files['excel'], base_dir)
    cv_path = save_file(files['cv'], base_dir)
    template_path = save_file(files['template'], base_dir)
    
    other_path = None
    if files.get('other'):
        other_path = save_file(files['other'], base_dir)
        app_state['other'] = str(other_path)

    df = pd.read_excel(excel_path)
    save_excel_to_db(files['excel'].read(), files['excel'].filename, bewerbungsname)

    for i, row in df.iterrows():
        out_dir = base_dir / str(i + 1)
        out_dir.mkdir(exist_ok=True)
        try:
            cover_pdf = generate_letter(template_path, row, out_dir)
            final_path = out_dir / f"{bewerbungsname}.pdf"
            merge_pdfs(cv_path, cover_pdf, pos, final_path)
            log_event(bewerbungsname, "generated", company_num=i+1, firma=row["firma"], status="ok")
            os.remove(cover_pdf)
            app_state['logs'].append({"id": i+1, "name": row['firma'], "status": "ok"})
        except Exception:
            log_event(bewerbungsname, "generated", company_num=i+1, firma=row.get("firma", ""), status="error")
            with open(out_dir / "skip.txt", "w") as f: f.write("skip")
            app_state['logs'].append({"id": i+1, "name": row['firma'], "status": "error"})
        
        app_state['progress'] = (i + 1) / len(df)

    app_state['total_companies'] = len(df)
    app_state['base_dir'] = str(base_dir)
    app_state['generated'] = True
    app_state['generating'] = False

def sending_thread(docx_path, delay, start, scheduled_dt, extra_path=None):
    app_state['sending'] = True
    app_state['logs'] = []
    app_state['progress'] = 0
    app_state['send_done'] = False
    app_state['interrupted_at'] = None
    
    if extra_path:
        app_state['extra'] = str(extra_path)

    doc = Document(docx_path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    subject = lines[0]
    message_template = "\n".join(lines[1:])

    if scheduled_dt:
        while True:
            remaining = (scheduled_dt - datetime.datetime.now()).total_seconds()
            if remaining <= 0:
                app_state['countdown_msg'] = "Time reached! Starting to send..."
                break
            h = int(remaining // 3600); m = int((remaining % 3600) // 60); s = int(remaining % 60)
            app_state['countdown_msg'] = f"Sending in: {h:02d}:{m:02d}:{s:02d}"
            time.sleep(1)

    base_dir = Path(app_state['base_dir'])
    network_error = False

    for cmp in range(start, app_state['total_companies'] + 1):
        cmp_dir = base_dir / str(cmp)
        saved_path = cmp_dir / "saved_values.py"
        pdf_path = cmp_dir / f"{app_state['bewerbungsname']}.pdf"
        email = "???"

        try:
            m = load_module(cmp, saved_path)
            email = m.old_email
        except Exception: pass

        if (cmp_dir / "skip.txt").exists():
            log_event(app_state['bewerbungsname'], "skip", company_num=cmp, email=email, status="skip")
            app_state['logs'].append({"id": cmp, "name": email, "status": "skip"})
            app_state['progress'] = cmp / app_state['total_companies']
            continue

        try:
            m = load_module(cmp, saved_path)
            email = m.old_email; gender = m.old_salutation; person = m.old_person; adresse_3 = m.old_adresse_3.strip()
            if person == "x": person = "Damen und Herren"

            letter = (message_template.replace("{person}", person).replace("{gender}", gender)
                      .replace("{adre}", adresse_3).replace("{space}", "\n").replace("{2space}", "\n\n"))

            gmail_send("", email, subject, letter, pdf_path, app_state['extra'], app_state['other'])
            files_sent = [str(pdf_path.name)]
            if app_state['extra']: files_sent.append(Path(app_state['extra']).name)
            if app_state['other']: files_sent.append(Path(app_state['other']).name)
            log_event(app_state['bewerbungsname'], "sent", company_num=cmp, email=email, status="sent", files_sent=files_sent)
            
            app_state['logs'].append({"id": cmp, "name": email, "status": "sent"})
            time.sleep(delay)
        except Exception as e:
            if is_network_error(e):
                log_event(app_state['bewerbungsname'], "network_error", company_num=cmp, email=email, status="network_error", error_msg=str(e))
                app_state['logs'].append({"id": cmp, "name": f"{email} - Verbindung unterbrochen", "status": "error"})
                app_state['interrupted_at'] = cmp
                network_error = True
                break
            else:
                log_event(app_state['bewerbungsname'], "error", company_num=cmp, email=email, status="error", error_msg=str(e))
                app_state['logs'].append({"id": cmp, "name": email, "status": "error"})

        app_state['progress'] = cmp / app_state['total_companies']

    if not network_error:
        app_state['send_done'] = True
    app_state['sending'] = False


# ==========================================
# FLASK ROUTES (APIs)
# ==========================================
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json
    code = data.get('code')
    if code == ADMIN_CODE:
        app_state['logged_in'] = True; app_state['is_admin'] = True
        return jsonify({"success": True, "admin": True})
    elif code == ACCESS_CODE:
        app_state['logged_in'] = True; app_state['is_admin'] = False
        return jsonify({"success": True, "admin": False})
    return jsonify({"success": False, "message": "Falscher Code"}), 401

@app.route('/api/state', methods=['GET'])
def get_state():
    return jsonify({k: v for k, v in app_state.items() if k != 'logs'}) 

@app.route('/api/reset', methods=['POST'])
def reset():
    reset_all()
    return jsonify({"success": True})

@app.route('/api/generate', methods=['POST'])
def generate():
    if not app_state['logged_in']: return jsonify({"error": "Unauthorized"}), 401
    
    files = {
        'excel': request.files.get('excel'),
        'cv': request.files.get('cv'),
        'template': request.files.get('template'),
        'other': request.files.get('other')
    }
    pos = int(request.form.get('pos', 2))
    bewerbungsname = request.form.get('bewerbungsname', 'Bewerbung')
    
    app_state['bewerbungsname'] = bewerbungsname
    app_state['anschreiben_pos'] = pos

    thread = threading.Thread(target=generation_thread, args=(files, pos, bewerbungsname))
    thread.start()
    return jsonify({"success": True, "message": "Generation started"})

@app.route('/api/send_emails', methods=['POST'])
def send_emails():
    if not app_state['logged_in']: return jsonify({"error": "Unauthorized"}), 401
    
    docx_file = request.files.get('docx_letter')
    extra_file = request.files.get('extra_file')
    delay = int(request.form.get('delay', 10))
    start = int(request.form.get('start', 1))
    schedule_str = request.form.get('scheduled_dt', None)
    
    base_dir = Path(app_state['base_dir'])
    docx_path = save_file(docx_file, base_dir)
    extra_path = save_file(extra_file, base_dir) if extra_file else None
    
    scheduled_dt = datetime.datetime.fromisoformat(schedule_str) if schedule_str else None
    app_state['scheduled_dt'] = str(scheduled_dt) if scheduled_dt else None

    thread = threading.Thread(target=sending_thread, args=(docx_path, delay, start, scheduled_dt, extra_path))
    thread.start()
    return jsonify({"success": True})

@app.route('/api/dashboard', methods=['GET'])
def dashboard():
    if not app_state.get('is_admin'): return jsonify({"error": "Access denied"}), 403
    try:
        sb = get_supabase()
        resp = sb.table("bewerber_logs").select("*").order("created_at", desc=True).limit(1000).execute()
        rows = resp.data or []
        ex_resp = sb.table("excel_uploads").select("*").order("uploaded_at", desc=True).limit(200).execute()
        excels = ex_resp.data or []
        return jsonify({"logs": rows, "excels": excels})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/download_excel/<file_id>', methods=['GET'])
def download_excel(file_id):
    try:
        sb = get_supabase()
        ex = sb.table("excel_uploads").select("*").eq("id", file_id).single().execute().data
        if not ex: return jsonify({"error": "Not found"}), 404
        
        file_bytes = sb.storage.from_("excel-files").download(ex["storage_path"])
        return send_file(
            io.BytesIO(file_bytes),
            download_name=ex["filename"],
            as_attachment=True
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, threaded=True)