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
from werkzeug.utils import secure_filename


# ==========================================
# APP
# ==========================================

app = Flask(__name__)

# DO NOT CHANGE
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100MB


# ==========================================
# CONFIGURATION
# ==========================================

ACCESS_CODE = os.environ.get("ACCESS_CODE", "user2024")
ADMIN_CODE = os.environ.get("ADMIN_CODE", "admin2024")

GMAIL_USER = os.environ.get(
    "GMAIL_USER",
    "wapoyassin08@gmail.com"
)

GMAIL_PASS = os.environ.get(
    "GMAIL_PASS",
    "ioat otqj kyte vduq"
)


# ==========================================
# DATA DIRECTORIES
# ==========================================

DATA_DIR = Path("data")
DATA_DIR.mkdir(parents=True, exist_ok=True)

EXCELS_DIR = DATA_DIR / "excels"
EXCELS_DIR.mkdir(parents=True, exist_ok=True)

APPLICATIONS_DIR = DATA_DIR / "applications"
APPLICATIONS_DIR.mkdir(parents=True, exist_ok=True)

LOGS_DB = DATA_DIR / "logs.db"


# ==========================================
# DATABASE
# ==========================================

def init_db():

    conn = sqlite3.connect(str(LOGS_DB))

    c = conn.cursor()

    c.execute("""
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
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS excel_uploads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_name TEXT,
            filename TEXT,
            storage_path TEXT,
            uploaded_at TEXT DEFAULT (datetime('now'))
        )
    """)

    conn.commit()
    conn.close()


init_db()


def get_db():

    conn = sqlite3.connect(
        str(LOGS_DB),
        check_same_thread=False
    )

    conn.row_factory = sqlite3.Row

    return conn


# ==========================================
# GLOBAL STATE
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
# HELPERS
# ==========================================

def convert(docx_path, pdf_path):

    docx_path = Path(docx_path)
    pdf_path = Path(pdf_path)

    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(pdf_path.parent),
            str(docx_path)
        ],
        check=True
    )


def is_network_error(e):

    err = str(e).lower()

    keywords = [
        "connection",
        "network",
        "timeout",
        "socket",
        "refused",
        "unreachable",
        "errno",
        "broken pipe",
        "reset",
        "ssl",
        "eof",
        "timed out"
    ]

    return (
        any(k in err for k in keywords)
        or isinstance(
            e,
            (
                socket.timeout,
                socket.gaierror,
                OSError
            )
        )
    )


# ==========================================
# RESET
# ==========================================

def reset_state():

    with state_lock:

        keys_to_reset = [
            "generating",
            "sending",
            "generated",

            "base_dir",

            "other",
            "extra",

            "scheduled_dt",

            "send_done",
            "interrupted_at",

            "gen_progress",
            "gen_log",
            "gen_total",

            "send_progress",
            "send_log",
            "send_total",

            "bewerbungsname",
            "total_companies",

            "anschreiben_pos",
            "delay",
            "start",

            "waiting_scheduled",
            "network_error",
        ]

        for key in keys_to_reset:

            if key not in state:
                continue

            if isinstance(state[key], list):

                state[key] = []

            elif isinstance(state[key], int):

                if key == "start":
                    state[key] = 1

                elif key == "anschreiben_pos":
                    state[key] = 2

                elif key == "delay":
                    state[key] = 10

                else:
                    state[key] = 0

            else:

                state[key] = None

        state["logged_in"] = True


# ==========================================
# LOGGING
# ==========================================

def log_event(
    session_name,
    event_type,
    company_num=None,
    email=None,
    firma=None,
    status=None,
    error_msg=None,
    files_sent=None
):

    conn = None

    try:

        conn = get_db()

        conn.execute(
            """
            INSERT INTO bewerber_logs
            (
                session_name,
                event_type,
                company_num,
                email,
                firma,
                status,
                error_msg,
                files_sent
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                session_name,
                event_type,
                company_num,
                email,
                firma,
                status,
                error_msg or "",
                json.dumps(files_sent or [])
            )
        )

        conn.commit()

    except Exception as e:

        print(
            "LOG ERROR:",
            repr(e)
        )

    finally:

        if conn:
            conn.close()


# ==========================================
# EXCEL LOCAL STORAGE
# ==========================================

def save_excel_to_db(
    file_path,
    filename,
    session_name
):

    try:

        file_path = Path(file_path)

        if not file_path.exists():

            raise FileNotFoundError(
                f"Excel file not found: {file_path}"
            )

        timestamp = datetime.datetime.utcnow().strftime(
            "%Y%m%d_%H%M%S_%f"
        )

        safe_filename = secure_filename(filename)

        if not safe_filename:

            safe_filename = "upload.xlsx"

        storage_filename = (
            f"{timestamp}_{safe_filename}"
        )

        storage_path = (
            EXCELS_DIR / storage_filename
        )

        shutil.copy2(
            str(file_path),
            str(storage_path)
        )

        conn = get_db()

        conn.execute(
            """
            INSERT INTO excel_uploads
            (
                session_name,
                filename,
                storage_path
            )
            VALUES (?, ?, ?)
            """,
            (
                session_name,
                filename,
                str(storage_path)
            )
        )

        conn.commit()
        conn.close()

        return True

    except Exception as e:

        print(
            "EXCEL STORAGE ERROR:",
            repr(e)
        )

        return False


# ==========================================
# CLEAN OLD EXCEL FILES
# ==========================================

def clean_old_excels():

    conn = None

    try:

        conn = get_db()

        yesterday = (
            datetime.datetime.utcnow()
            - datetime.timedelta(days=1)
        ).strftime("%Y-%m-%d %H:%M:%S")

        old = conn.execute(
            """
            SELECT *
            FROM excel_uploads
            WHERE uploaded_at < ?
            """,
            (yesterday,)
        ).fetchall()

        for ex in old:

            try:

                path = Path(
                    ex["storage_path"]
                )

                if path.exists():

                    path.unlink()

            except Exception as e:

                print(
                    "OLD EXCEL DELETE ERROR:",
                    repr(e)
                )

            conn.execute(
                """
                DELETE FROM excel_uploads
                WHERE id = ?
                """,
                (ex["id"],)
            )

        conn.commit()

    except Exception as e:

        print(
            "CLEAN EXCEL ERROR:",
            repr(e)
        )

    finally:

        if conn:
            conn.close()


# ==========================================
# FILE HELPERS
# ==========================================

def get_safe_filename(filename, fallback):

    if not filename:

        return fallback

    filename = secure_filename(
        Path(filename).name
    )

    if not filename:

        return fallback

    return filename


def save_upload(
    file_storage,
    target_dir,
    fallback_name="upload.bin"
):

    if file_storage is None:

        raise ValueError(
            "No file supplied"
        )

    filename = get_safe_filename(
        file_storage.filename,
        fallback_name
    )

    target_dir = Path(target_dir)

    target_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    path = target_dir / filename

    file_storage.save(str(path))

    if not path.exists():

        raise IOError(
            f"File was not saved: {path}"
        )

    if path.stat().st_size <= 0:

        raise IOError(
            f"Saved file is empty: {path}"
        )

    return path


def save_file_from_request(
    file_storage,
    target_dir,
    fallback_name
):

    return save_upload(
        file_storage,
        target_dir,
        fallback_name
    )


def save_other_file(
    file_storage,
    target_dir
):

    if not file_storage:

        return None

    if not file_storage.filename:

        return None

    path = save_upload(
        file_storage,
        target_dir,
        "Zeugnisse.pdf"
    )

    with state_lock:

        state["other"] = str(path)

    return path


def save_extra_file(
    file_storage,
    target_dir
):

    if not file_storage:

        with state_lock:
            state["extra"] = None

        return None

    if not file_storage.filename:

        with state_lock:
            state["extra"] = None

        return None

    filename = get_safe_filename(
        file_storage.filename,
        "Anhang.pdf"
    )

    if "." not in filename:

        filename = "Anhang.pdf"

    path = save_upload(
        file_storage,
        target_dir,
        filename
    )

    with state_lock:

        state["extra"] = str(path)

    return path


# ==========================================
# SAVE ALL GENERATE FILES
# ==========================================

def save_generate_files(
    excel_file,
    cv_file,
    template_file,
    other_file,
    base_dir
):

    base_dir = Path(base_dir)

    base_dir.mkdir(
        parents=True,
        exist_ok=True
    )

    # --------------------------------------
    # Excel
    # --------------------------------------

    excel_path = save_file_from_request(
        excel_file,
        base_dir,
        "companies.xlsx"
    )

    # --------------------------------------
    # CV
    # --------------------------------------

    cv_path = save_file_from_request(
        cv_file,
        base_dir,
        "Lebenslauf.pdf"
    )

    # --------------------------------------
    # Template
    # --------------------------------------

    template_path = save_file_from_request(
        template_file,
        base_dir,
        "Anschreiben.docx"
    )

    # --------------------------------------
    # Other
    # --------------------------------------

    other_path = save_file_from_request(
        other_file,
        base_dir,
        "Zeugnisse.pdf"
    )

    # --------------------------------------
    # Verify everything
    # --------------------------------------

    files = [
        excel_path,
        cv_path,
        template_path,
        other_path
    ]

    for path in files:

        if not path.exists():

            raise IOError(
                f"File missing after upload: {path}"
            )

        if path.stat().st_size <= 0:

            raise IOError(
                f"File empty after upload: {path}"
            )

    return (
        excel_path,
        cv_path,
        template_path,
        other_path
    )


# ==========================================
# OLD VALUES
# ==========================================

def save_old_values(
    salutation,
    person_full,
    email,
    adresse_3,
    output_dir
):

    content = f'''
old_salutation = {salutation!r}
old_person = {person_full!r}
old_email = {email!r}
old_adresse_3 = {adresse_3!r}
'''

    with open(
        Path(output_dir) / "saved_values.py",
        "w",
        encoding="utf-8"
    ) as f:

        f.write(content)


# ==========================================
# GENERATE LETTER
# ==========================================

def generate_letter(
    template_path,
    row,
    output_dir
):

    template_path = Path(template_path)
    output_dir = Path(output_dir)

    doc = Document(
        str(template_path)
    )

    person_full = str(
        row["person"]
    )

    firma = str(
        row["firma"]
    )

    adresse = str(
        row["adresse"]
    )

    email = str(
        row["email"]
    )

    # --------------------------------------
    # Gender
    # --------------------------------------

    if person_full.startswith("Herr"):

        salutation = "er "
        gender_def = True

    elif person_full.startswith("Frau"):

        salutation = "e "
        gender_def = True

    else:

        salutation = "e "
        gender_def = False

    # --------------------------------------
    # Address
    # --------------------------------------

    if "|" in adresse:

        adresse_1, adresse_2 = [
            x.strip()
            for x in adresse.split("|", 1)
        ]

    else:

        adresse_1 = adresse
        adresse_2 = adresse

    adresse_3 = adresse_2[6:]

    # --------------------------------------
    # Replace placeholders
    # --------------------------------------

    for p in doc.paragraphs:

        full = "".join(
            r.text
            for r in p.runs
        )

        full = full.replace(
            "{#custom}",
            firma
        )

        full = full.replace(
            "{gender}",
            salutation
        )

        full = full.replace(
            "{zeit}",
            datetime.date.today().strftime(
                "%d.%m.%Y"
            )
        )

        if gender_def:

            full = full.replace(
                "{person}",
                person_full
            )

            full = full.replace(
                "{.}",
                person_full
            )

            full = full.replace(
                "{/custom}",
                adresse_1
            )

            full = full.replace(
                "{/custom2}",
                adresse_2
            )

            full = full.replace(
                "{adre}",
                adresse_3
            )

        else:

            full = full.replace(
                "{person}",
                "Damen und Herren"
            )

            full = full.replace(
                "{.}",
                adresse_1
            )

            full = full.replace(
                "{/custom}",
                adresse_2
            )

            full = full.replace(
                "{/custom2}",
                ""
            )

            full = full.replace(
                "{adre}",
                adresse_3
            )

        for r in p.runs:

            r.text = ""

        if len(p.runs) == 0:

            p.add_run(full)

        else:

            p.runs[0].text = full

    # --------------------------------------
    # Save DOCX
    # --------------------------------------

    safe_name = (
        firma
        .replace("/", "_")
        .replace("\\", "_")
    )

    safe_name = (
        get_safe_filename(
            safe_name,
            "Firma"
        )
    )

    docx_path = (
        output_dir /
        f"{safe_name}.docx"
    )

    pdf_path = (
        output_dir /
        f"{safe_name}.pdf"
    )

    doc.save(
        str(docx_path)
    )

    # --------------------------------------
    # Convert
    # --------------------------------------

    convert(
        docx_path,
        pdf_path
    )

    # --------------------------------------
    # Delete DOCX
    # --------------------------------------

    if docx_path.exists():

        docx_path.unlink()

    # --------------------------------------
    # Save values
    # --------------------------------------

    save_old_values(
        salutation,
        person_full,
        email,
        adresse_3,
        output_dir
    )

    return pdf_path


# ==========================================
# MERGE PDF
# ==========================================

def merge_pdfs(
    cv_path,
    cover_path,
    position,
    output_path
):

    cv_path = Path(cv_path)
    cover_path = Path(cover_path)
    output_path = Path(output_path)

    main_reader = PdfReader(
        str(cv_path)
    )

    insert_reader = PdfReader(
        str(cover_path)
    )

    writer = PdfWriter()

    for i in range(
        len(main_reader.pages)
    ):

        if i == position - 1:

            writer.add_page(
                insert_reader.pages[0]
            )

        writer.add_page(
            main_reader.pages[i]
        )

    if position > len(
        main_reader.pages
    ):

        writer.add_page(
            insert_reader.pages[0]
        )

    with open(
        output_path,
        "wb"
    ) as f:

        writer.write(f)


# ==========================================
# PATH
# ==========================================

def safe_path(p):

    if p and isinstance(p, str):

        return Path(p)

    return p


# ==========================================
# LOAD SAVED VALUES
# ==========================================

def load_module(
    cmp,
    saved_path
):

    module_name = (
        f"vals_{cmp}"
    )

    spec = (
        importlib.util
        .spec_from_file_location(
            module_name,
            str(saved_path)
        )
    )

    if spec is None or spec.loader is None:

        raise ImportError(
            f"Could not load {saved_path}"
        )

    module = (
        importlib.util
        .module_from_spec(spec)
    )

    sys.modules[
        module_name
    ] = module

    spec.loader.exec_module(
        module
    )

    return module


# ==========================================
# GMAIL
# ==========================================

def gmail_send(
    to_email,
    subject,
    body,
    file1_path=None,
    file2_path=None,
    file3_path=None
):

    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587

    message = MIMEMultipart()

    message["From"] = GMAIL_USER
    message["To"] = to_email
    message["Subject"] = subject

    message.attach(
        MIMEText(
            body,
            "plain"
        )
    )

    for fpath in [
        safe_path(file1_path),
        safe_path(file2_path),
        safe_path(file3_path)
    ]:

        if not fpath:
            continue

        if not fpath.exists():
            continue

        with open(
            fpath,
            "rb"
        ) as f:

            part = MIMEBase(
                "application",
                "octet-stream"
            )

            part.set_payload(
                f.read()
            )

        encoders.encode_base64(
            part
        )

        part.add_header(
            "Content-Disposition",
            f'attachment; filename="{fpath.name}"'
        )

        message.attach(part)

    server = smtplib.SMTP(
        SMTP_SERVER,
        SMTP_PORT
    )

    try:

        server.starttls()

        server.login(
            GMAIL_USER,
            GMAIL_PASS
        )

        server.sendmail(
            GMAIL_USER,
            to_email,
            message.as_string()
        )

    finally:

        server.quit()


# ==========================================
# BACKGROUND THREAD
# GENERATE
# ==========================================

def generate_thread(
    excel_path,
    cv_path,
    template_path,
    other_path,
    anschreiben_pos,
    bewerbungsname,
    base_dir
):

    try:

        excel_path = Path(
            excel_path
        )

        cv_path = Path(
            cv_path
        )

        template_path = Path(
            template_path
        )

        other_path = Path(
            other_path
        ) if other_path else None

        base_dir = Path(
            base_dir
        )

        with state_lock:

            state["generating"] = True
            state["generated"] = False

            state["gen_progress"] = 0
            state["gen_log"] = []

            state["base_dir"] = str(
                base_dir
            )

            state["bewerbungsname"] = (
                bewerbungsname
            )

            state["anschreiben_pos"] = (
                anschreiben_pos
            )

            state["other"] = (
                str(other_path)
                if other_path
                else None
            )

        # --------------------------------------
        # Save Excel local copy
        # --------------------------------------

        save_excel_to_db(
            excel_path,
            excel_path.name,
            bewerbungsname
        )

        # --------------------------------------
        # Read Excel
        # --------------------------------------

        df = pd.read_excel(
            str(excel_path)
        )

        total = len(df)

        with state_lock:

            state["gen_total"] = total
            state["total_companies"] = total

        # --------------------------------------
        # No companies
        # --------------------------------------

        if total == 0:

            with state_lock:

                state["generating"] = False
                state["generated"] = True
                state["gen_progress"] = 1

            return

        # --------------------------------------
        # Generate
        # --------------------------------------

        for i, row in df.iterrows():

            cmp_num = i + 1

            out_dir = (
                base_dir /
                str(cmp_num)
            )

            out_dir.mkdir(
                parents=True,
                exist_ok=True
            )

            firma = str(
                row.get(
                    "firma",
                    ""
                )
            )

            try:

                cover_pdf = generate_letter(
                    template_path,
                    row,
                    out_dir
                )

                final_path = (
                    out_dir /
                    f"{bewerbungsname}.pdf"
                )

                merge_pdfs(
                    cv_path,
                    cover_pdf,
                    anschreiben_pos,
                    final_path
                )

                log_event(
                    bewerbungsname,
                    "generated",
                    company_num=cmp_num,
                    firma=firma,
                    status="ok"
                )

                if cover_pdf.exists():

                    cover_pdf.unlink()

                entry = {
                    "num": cmp_num,
                    "firma": firma,
                    "status": "ok"
                }

            except Exception as e:

                print(
                    f"GENERATION ERROR #{cmp_num}:",
                    repr(e)
                )

                log_event(
                    bewerbungsname,
                    "generated",
                    company_num=cmp_num,
                    firma=firma,
                    status="error",
                    error_msg=str(e)
                )

                with open(
                    out_dir / "skip.txt",
                    "w",
                    encoding="utf-8"
                ) as f:

                    f.write("skip")

                entry = {
                    "num": cmp_num,
                    "firma": firma,
                    "status": "error",
                    "error": str(e)
                }

            with state_lock:

                state["gen_log"].append(
                    entry
                )

                state["gen_progress"] = (
                    cmp_num / total
                )

        # --------------------------------------
        # Complete
        # --------------------------------------

        with state_lock:

            state["generating"] = False
            state["generated"] = True
            state["gen_progress"] = 1

    except Exception as e:

        print(
            "FATAL GENERATION ERROR:",
            repr(e)
        )

        log_event(
            bewerbungsname,
            "generation_fatal",
            status="error",
            error_msg=str(e)
        )

        with state_lock:

            state["generating"] = False
            state["generated"] = False

            state["gen_log"].append({
                "num": 0,
                "firma": "",
                "status": "fatal",
                "error": str(e)
            })


# ==========================================
# BACKGROUND THREAD
# SEND
# ==========================================

def send_thread(
    letter_path,
    delay,
    start_num,
    scheduled_dt,
    extra_path
):

    try:

        letter_path = Path(
            letter_path
        )

        extra_path = (
            Path(extra_path)
            if extra_path
            else None
        )

        with state_lock:

            state["sending"] = True
            state["send_done"] = False
            state["interrupted_at"] = None

            state["send_log"] = []
            state["send_progress"] = 0

            state["start"] = start_num
            state["scheduled_dt"] = scheduled_dt

            state["network_error"] = False

            base_dir = Path(
                state["base_dir"]
            )

            bewerbungsname = (
                state["bewerbungsname"]
            )

            total = (
                state["total_companies"]
            )

            state["extra"] = (
                str(extra_path)
                if extra_path
                else None
            )

        # --------------------------------------
        # Check template
        # --------------------------------------

        if not letter_path.exists():

            raise FileNotFoundError(
                f"Email template not found: {letter_path}"
            )

        # --------------------------------------
        # Read template
        # --------------------------------------

        doc = Document(
            str(letter_path)
        )

        lines = [
            p.text.strip()
            for p in doc.paragraphs
            if p.text.strip()
        ]

        if not lines:

            raise ValueError(
                "Email template is empty"
            )

        subject = lines[0]

        message_template = "\n".join(
            lines[1:]
        )

        # --------------------------------------
        # Scheduled
        # --------------------------------------

        if scheduled_dt:

            with state_lock:

                state[
                    "waiting_scheduled"
                ] = True

            while True:

                now = datetime.datetime.now()

                remaining = (
                    scheduled_dt - now
                ).total_seconds()

                if remaining <= 0:
                    break

                time.sleep(
                    min(
                        remaining,
                        1
                    )
                )

            with state_lock:

                state[
                    "waiting_scheduled"
                ] = False

        # --------------------------------------
        # Send
        # --------------------------------------

        for cmp in range(
            start_num,
            total + 1
        ):

            cmp_dir = (
                base_dir /
                str(cmp)
            )

            saved_path = (
                cmp_dir /
                "saved_values.py"
            )

            pdf_path = (
                cmp_dir /
                f"{bewerbungsname}.pdf"
            )

            email = "???"

            try:

                m = load_module(
                    cmp,
                    saved_path
                )

                email = m.old_email

            except Exception:

                pass

            # ----------------------------------
            # Skip
            # ----------------------------------

            if (
                cmp_dir /
                "skip.txt"
            ).exists():

                log_event(
                    bewerbungsname,
                    "skip",
                    company_num=cmp,
                    email=email,
                    status="skip"
                )

                with state_lock:

                    state[
                        "send_log"
                    ].append({
                        "num": cmp,
                        "email": email,
                        "status": "skip"
                    })

                    state[
                        "send_progress"
                    ] = cmp / total

                continue

            # ----------------------------------
            # Send company
            # ----------------------------------

            try:

                m = load_module(
                    cmp,
                    saved_path
                )

                email = m.old_email
                gender = m.old_salutation
                person = m.old_person
                adresse_3 = (
                    m.old_adresse_3.strip()
                )

                if person == "x":

                    person = (
                        "Damen und Herren"
                    )

                letter = (
                    message_template
                    .replace(
                        "{person}",
                        person
                    )
                    .replace(
                        "{gender}",
                        gender
                    )
                    .replace(
                        "{adre}",
                        adresse_3
                    )
                    .replace(
                        "{space}",
                        "\n"
                    )
                    .replace(
                        "{2space}",
                        "\n\n"
                    )
                )

                with state_lock:

                    other_path = (
                        state.get("other")
                    )

                    current_extra = (
                        state.get("extra")
                    )

                gmail_send(
                    email,
                    subject,
                    letter,
                    str(pdf_path)
                    if pdf_path.exists()
                    else None,
                    current_extra,
                    other_path
                )

                files_sent = []

                if pdf_path.exists():

                    files_sent.append(
                        pdf_path.name
                    )

                if current_extra:

                    files_sent.append(
                        Path(
                            current_extra
                        ).name
                    )

                if other_path:

                    files_sent.append(
                        Path(
                            other_path
                        ).name
                    )

                log_event(
                    bewerbungsname,
                    "sent",
                    company_num=cmp,
                    email=email,
                    status="sent",
                    files_sent=files_sent
                )

                with state_lock:

                    state[
                        "send_log"
                    ].append({
                        "num": cmp,
                        "email": email,
                        "status": "sent"
                    })

                    state[
                        "send_progress"
                    ] = cmp / total

                time.sleep(
                    delay
                )

            except Exception as e:

                print(
                    f"SEND ERROR #{cmp}:",
                    repr(e)
                )

                if is_network_error(e):

                    log_event(
                        bewerbungsname,
                        "network_error",
                        company_num=cmp,
                        email=email,
                        status="network_error",
                        error_msg=str(e)
                    )

                    with state_lock:

                        state[
                            "send_log"
                        ].append({
                            "num": cmp,
                            "email": email,
                            "status": "network_error"
                        })

                        state[
                            "interrupted_at"
                        ] = cmp

                        state[
                            "sending"
                        ] = False

                        state[
                            "network_error"
                        ] = True

                    return

                else:

                    log_event(
                        bewerbungsname,
                        "error",
                        company_num=cmp,
                        email=email,
                        status="error",
                        error_msg=str(e)
                    )

                    with state_lock:

                        state[
                            "send_log"
                        ].append({
                            "num": cmp,
                            "email": email,
                            "status": "error"
                        })

                        state[
                            "send_progress"
                        ] = cmp / total

        # --------------------------------------
        # Complete
        # --------------------------------------

        with state_lock:

            state["sending"] = False
            state["scheduled_dt"] = None
            state["send_done"] = True

    except Exception as e:

        print(
            "FATAL SEND ERROR:",
            repr(e)
        )

        with state_lock:

            state["sending"] = False
            state["waiting_scheduled"] = False
            state["send_done"] = False

            state[
                "send_log"
            ].append({
                "num": 0,
                "email": "",
                "status": "fatal",
                "error": str(e)
            })


# ==========================================
# ROUTES
# ==========================================

@app.route("/")
def index():

    return render_template(
        "index.html"
    )


# ==========================================
# LOGIN
# ==========================================

@app.route(
    "/api/login",
    methods=["POST"]
)
def login():

    data = request.json or {}

    code = data.get(
        "code",
        ""
    )

    if code == ADMIN_CODE:

        with state_lock:

            state["logged_in"] = True
            state["is_admin"] = True

        return jsonify({
            "success": True,
            "is_admin": True
        })

    elif code == ACCESS_CODE:

        with state_lock:

            state["logged_in"] = True
            state["is_admin"] = False

        return jsonify({
            "success": True,
            "is_admin": False
        })

    return jsonify({
        "success": False,
        "error": "Falscher Code"
    }), 401


# ==========================================
# STATE
# ==========================================

@app.route(
    "/api/state",
    methods=["GET"]
)
def get_state():

    with state_lock:

        return jsonify({

            "logged_in":
                state["logged_in"],

            "is_admin":
                state["is_admin"],

            "generating":
                state["generating"],

            "generated":
                state["generated"],

            "sending":
                state["sending"],

            "send_done":
                state["send_done"],

            "interrupted_at":
                state["interrupted_at"],

            "bewerbungsname":
                state["bewerbungsname"],

            "total_companies":
                state["total_companies"],

            "waiting_scheduled":
                state.get(
                    "waiting_scheduled",
                    False
                ),

            "scheduled_dt":
                (
                    state["scheduled_dt"].isoformat()
                    if state.get("scheduled_dt")
                    else None
                ),
        })


# ==========================================
# GENERATE
# ==========================================

@app.route(
    "/api/generate",
    methods=["POST"]
)
def generate():

    with state_lock:

        if not state["logged_in"]:

            return jsonify({
                "error": "Not logged in"
            }), 403

        if state["generating"]:

            return jsonify({
                "error": "Generierung läuft bereits"
            }), 400

    # --------------------------------------
    # Receive files
    # --------------------------------------

    excel_file = request.files.get(
        "excel"
    )

    cv_file = request.files.get(
        "cv"
    )

    template_file = request.files.get(
        "template"
    )

    other_file = request.files.get(
        "other"
    )

    # --------------------------------------
    # Validate
    # --------------------------------------

    if not excel_file:
        return jsonify({
            "error": "Excel fehlt"
        }), 400

    if not cv_file:
        return jsonify({
            "error": "Lebenslauf fehlt"
        }), 400

    if not template_file:
        return jsonify({
            "error": "Anschreiben fehlt"
        }), 400

    if not other_file:
        return jsonify({
            "error": "Zeugnisse und Zertifikate fehlen"
        }), 400

    # --------------------------------------
    # Position
    # --------------------------------------

    try:

        anschreiben_pos = int(
            request.form.get(
                "position",
                2
            )
        )

        if anschreiben_pos < 1:

            anschreiben_pos = 1

    except Exception:

        anschreiben_pos = 2

    # --------------------------------------
    # Application name
    # --------------------------------------

    bewerbungsname = (
        request.form.get(
            "bewerbungsname",
            "Bewerbung"
        ).strip()
    )

    if not bewerbungsname:

        bewerbungsname = "Bewerbung"

    # --------------------------------------
    # Safe directory name
    # --------------------------------------

    safe_bewerbungsname = secure_filename(
        bewerbungsname
    )

    if not safe_bewerbungsname:

        safe_bewerbungsname = "Bewerbung"

    # --------------------------------------
    # Create directory
    # --------------------------------------

    base_dir = (
        APPLICATIONS_DIR /
        safe_bewerbungsname
    )

    try:

        if base_dir.exists():

            shutil.rmtree(
                str(base_dir)
            )

        base_dir.mkdir(
            parents=True,
            exist_ok=True
        )

        # ----------------------------------
        # IMPORTANT:
        # SAVE FILES INSIDE REQUEST
        # BEFORE THREAD
        # ----------------------------------

        (
            excel_path,
            cv_path,
            template_path,
            other_path
        ) = save_generate_files(
            excel_file,
            cv_file,
            template_file,
            other_file,
            base_dir
        )

    except Exception as e:

        print(
            "UPLOAD ERROR:",
            repr(e)
        )

        return jsonify({
            "success": False,
            "error": (
                "Fehler beim Speichern der Dateien: "
                + str(e)
            )
        }), 500

    # --------------------------------------
    # Update state
    # --------------------------------------

    with state_lock:

        state["generating"] = True
        state["generated"] = False

        state["gen_progress"] = 0
        state["gen_log"] = []

        state["gen_total"] = 0

        state["base_dir"] = str(
            base_dir
        )

        state["bewerbungsname"] = (
            safe_bewerbungsname
        )

        state["anschreiben_pos"] = (
            anschreiben_pos
        )

        state["other"] = str(
            other_path
        )

        state["extra"] = None

    # --------------------------------------
    # Save Excel copy
    # --------------------------------------

    save_excel_to_db(
        excel_path,
        excel_file.filename,
        safe_bewerbungsname
    )

    # --------------------------------------
    # START THREAD
    #
    # ONLY PATHS ARE SENT
    # NO request.files
    # --------------------------------------

    thread = threading.Thread(
        target=generate_thread,
        args=(
            str(excel_path),
            str(cv_path),
            str(template_path),
            str(other_path),
            anschreiben_pos,
            safe_bewerbungsname,
            str(base_dir)
        )
    )

    thread.daemon = True

    thread.start()

    return jsonify({
        "success": True
    })


# ==========================================
# GENERATE STATUS
# ==========================================

@app.route(
    "/api/generate/status",
    methods=["GET"]
)
def generate_status():

    with state_lock:

        return jsonify({

            "generating":
                state["generating"],

            "generated":
                state["generated"],

            "progress":
                state["gen_progress"],

            "log":
                state["gen_log"],

            "total":
                state["gen_total"],

        })


# ==========================================
# SEND
# ==========================================

@app.route(
    "/api/send",
    methods=["POST"]
)
def send():

    with state_lock:

        if not state["logged_in"]:

            return jsonify({
                "error": "Not logged in"
            }), 403

        if not state["generated"]:

            return jsonify({
                "error":
                "Bitte zuerst Anschreiben generieren"
            }), 400

        if state["sending"]:

            return jsonify({
                "error":
                "Senden läuft bereits"
            }), 400

    letter_file = request.files.get(
        "letter"
    )

    extra_file = request.files.get(
        "extra"
    )

    if not letter_file:

        return jsonify({
            "error":
            "Email Template erforderlich"
        }), 400

    try:

        delay = int(
            request.form.get(
                "delay",
                10
            )
        )

    except Exception:

        delay = 10

    try:

        start_num = int(
            request.form.get(
                "start",
                1
            )
        )

    except Exception:

        start_num = 1

    schedule_str = request.form.get(
        "scheduled_dt",
        ""
    )

    scheduled_dt = None

    if schedule_str:

        try:

            scheduled_dt = (
                datetime.datetime
                .fromisoformat(
                    schedule_str
                )
            )

            if (
                scheduled_dt
                <= datetime.datetime.now()
            ):

                return jsonify({
                    "error":
                    "Die gewählte Zeit liegt in der Vergangenheit"
                }), 400

        except Exception:

            return jsonify({
                "error":
                "Ungültiges Datum/Zeit Format"
            }), 400

    # --------------------------------------
    # SAVE EMAIL TEMPLATE BEFORE THREAD
    # --------------------------------------

    with state_lock:

        base_dir = Path(
            state["base_dir"]
        )

    try:

        letter_path = save_upload(
            letter_file,
            base_dir,
            "Email_Template.docx"
        )

        if extra_file and extra_file.filename:

            extra_path = save_extra_file(
                extra_file,
                base_dir
            )

        else:

            extra_path = None

            with state_lock:

                state["extra"] = None

    except Exception as e:

        print(
            "SEND UPLOAD ERROR:",
            repr(e)
        )

        return jsonify({
            "success": False,
            "error":
            "Fehler beim Speichern der Email-Dateien: "
            + str(e)
        }), 500

    # --------------------------------------
    # State
    # --------------------------------------

    with state_lock:

        state["sending"] = True
        state["send_done"] = False
        state["interrupted_at"] = None

        state["send_progress"] = 0
        state["send_log"] = []

        state["delay"] = delay
        state["start"] = start_num

        state["scheduled_dt"] = scheduled_dt

        state["network_error"] = False

    # --------------------------------------
    # Start thread
    # ONLY PATHS
    # --------------------------------------

    thread = threading.Thread(
        target=send_thread,
        args=(
            str(letter_path),
            delay,
            start_num,
            scheduled_dt,
            str(extra_path)
            if extra_path
            else None
        )
    )

    thread.daemon = True

    thread.start()

    return jsonify({
        "success": True
    })


# ==========================================
# SEND STATUS
# ==========================================

@app.route(
    "/api/send/status",
    methods=["GET"]
)
def send_status():

    with state_lock:

        return jsonify({

            "sending":
                state["sending"],

            "send_done":
                state["send_done"],

            "interrupted_at":
                state["interrupted_at"],

            "progress":
                state["send_progress"],

            "log":
                state["send_log"],

            "total":
                state["total_companies"],

            "waiting_scheduled":
                state.get(
                    "waiting_scheduled",
                    False
                ),

            "scheduled_dt":
                (
                    state["scheduled_dt"].isoformat()
                    if state.get("scheduled_dt")
                    else None
                ),

            "bewerbungsname":
                state["bewerbungsname"],

        })


# ==========================================
# RESUME SEND
# ==========================================

@app.route(
    "/api/send/resume",
    methods=["POST"]
)
def resume_send():

    with state_lock:

        if not state["logged_in"]:

            return jsonify({
                "error": "Not logged in"
            }), 403

        base_dir = Path(
            state["base_dir"]
        )

        total = (
            state["total_companies"]
        )

        delay = (
            state["delay"]
        )

    data = request.json or {}

    try:

        resume_from = int(
            data.get(
                "resume_from",
                1
            )
        )

    except Exception:

        resume_from = 1

    if resume_from < 1:

        resume_from = 1

    if total and resume_from > total:

        return jsonify({
            "error":
            "Ungültige Startnummer"
        }), 400

    # --------------------------------------
    # Find saved email template
    # --------------------------------------

    letter_files = list(
        base_dir.glob(
            "*.docx"
        )
    )

    if not letter_files:

        return jsonify({
            "error":
            "Email Template nicht gefunden"
        }), 400

    letter_path = letter_files[0]

    # --------------------------------------
    # Extra attachment
    # --------------------------------------

    with state_lock:

        extra_path = state.get(
            "extra"
        )

    with state_lock:

        state["interrupted_at"] = None
        state["network_error"] = False

        state["sending"] = True
        state["send_done"] = False

        state["start"] = resume_from

        state["send_progress"] = 0
        state["send_log"] = []

    # --------------------------------------
    # Start with real PATH
    # --------------------------------------

    thread = threading.Thread(
        target=send_thread,
        args=(
            str(letter_path),
            delay,
            resume_from,
            None,
            extra_path
        )
    )

    thread.daemon = True

    thread.start()

    return jsonify({
        "success": True
    })


# ==========================================
# RESET
# ==========================================

@app.route(
    "/api/reset",
    methods=["POST"]
)
def reset():

    reset_state()

    return jsonify({
        "success": True
    })


# ==========================================
# DASHBOARD
# ==========================================

@app.route(
    "/api/dashboard",
    methods=["GET"]
)
def dashboard():

    with state_lock:

        if (
            not state["logged_in"]
            or not state["is_admin"]
        ):

            return jsonify({
                "error":
                "Access denied"
            }), 403

    session_filter = request.args.get(
        "session",
        "Alle"
    )

    hours_filter = request.args.get(
        "hours",
        type=int
    )

    conn = get_db()

    query = """
        SELECT *
        FROM bewerber_logs
    """

    conditions = []
    params = []

    if hours_filter:

        cutoff = (
            datetime.datetime.utcnow()
            - datetime.timedelta(
                hours=hours_filter
            )
        ).strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        conditions.append(
            "created_at >= ?"
        )

        params.append(
            cutoff
        )

    if conditions:

        query += (
            " WHERE "
            + " AND ".join(
                conditions
            )
        )

    query += """
        ORDER BY created_at DESC
        LIMIT 1000
    """

    rows = conn.execute(
        query,
        params
    ).fetchall()

    conn.close()

    # --------------------------------------
    # Stats
    # --------------------------------------

    sent_n = sum(
        1
        for r in rows
        if r["event_type"] == "sent"
    )

    skip_n = sum(
        1
        for r in rows
        if r["event_type"] == "skip"
    )

    error_n = sum(
        1
        for r in rows
        if r["event_type"]
        in (
            "error",
            "network_error"
        )
    )

    gen_ok = sum(
        1
        for r in rows
        if (
            r["event_type"]
            == "generated"
            and r["status"]
            == "ok"
        )
    )

    # --------------------------------------
    # Sessions
    # --------------------------------------

    sessions = list(
        set(
            r["session_name"]
            for r in rows
            if r["session_name"]
        )
    )

    # --------------------------------------
    # Companies
    # --------------------------------------

    companies = defaultdict(
        lambda: {
            "session": "",
            "generiert": "—",
            "gesendet": "Nein",
            "email_firma": "",
            "fehler": "—",
            "zeit": ""
        }
    )

    for r in sorted(
        rows,
        key=lambda x: x["created_at"]
    ):

        key = (
            r["session_name"],
            r["company_num"]
        )

        e = companies[key]

        e["session"] = (
            r["session_name"]
            or ""
        )

        if r["event_type"] == "generated":

            e["generiert"] = (
                "Ja"
                if r["status"] == "ok"
                else "Nein"
            )

            e["email_firma"] = (
                r["firma"]
                or ""
            )

            e["zeit"] = (
                r["created_at"][:16]
                if r["created_at"]
                else ""
            )

        if r["event_type"] == "sent":

            e["gesendet"] = "Ja"

            e["email_firma"] = (
                r["email"]
                or e["email_firma"]
            )

            e["zeit"] = (
                r["created_at"][:16]
                if r["created_at"]
                else ""
            )

        if r["event_type"] == "skip":

            e["gesendet"] = "Nein"

        if r["event_type"] in (
            "error",
            "network_error"
        ):

            e["fehler"] = "Ja"

    # --------------------------------------
    # Session filter
    # --------------------------------------

    if session_filter != "Alle":

        companies = {
            k: v
            for k, v
            in companies.items()
            if k[0] == session_filter
        }

    # --------------------------------------
    # Company list
    # --------------------------------------

    company_list = []

    for (
        session,
        cmp_num
    ), e in sorted(
        companies.items(),
        key=lambda x: (
            x[0][1]
            or 0
        )
    ):

        company_list.append({

            "session":
                e["session"],

            "cmp_num":
                cmp_num,

            "generiert":
                e["generiert"],

            "gesendet":
                e["gesendet"],

            "email_firma":
                e["email_firma"],

            "fehler":
                e["fehler"],

            "zeit":
                e["zeit"],
        })

    return jsonify({

        "stats": {

            "generated":
                gen_ok,

            "sent":
                sent_n,

            "skipped":
                skip_n,

            "errors":
                error_n,

        },

        "sessions":
            sessions,

        "companies":
            company_list,

    })


# ==========================================
# DASHBOARD EXCELS
# ==========================================

@app.route(
    "/api/dashboard/excels",
    methods=["GET"]
)
def dashboard_excels():

    with state_lock:

        if (
            not state["logged_in"]
            or not state["is_admin"]
        ):

            return jsonify({
                "error":
                "Access denied"
            }), 403

    clean_old_excels()

    conn = get_db()

    rows = conn.execute(
        """
        SELECT *
        FROM excel_uploads
        ORDER BY uploaded_at DESC
        LIMIT 200
        """
    ).fetchall()

    conn.close()

    session_filter = request.args.get(
        "session",
        "Alle"
    )

    files = []

    for r in rows:

        if (
            session_filter != "Alle"
            and r["session_name"]
            != session_filter
        ):

            continue

        files.append({

            "id":
                r["id"],

            "filename":
                r["filename"],

            "session_name":
                r["session_name"],

            "uploaded_at":
                (
                    r["uploaded_at"][:16]
                    if r["uploaded_at"]
                    else ""
                ),

        })

    all_sessions = list(
        set(
            r["session_name"]
            for r in rows
            if r["session_name"]
        )
    )

    return jsonify({

        "files":
            files,

        "sessions":
            all_sessions

    })


# ==========================================
# DOWNLOAD EXCEL
# ==========================================

@app.route(
    "/api/dashboard/excel/download/<int:file_id>"
)
def download_excel(file_id):

    with state_lock:

        if (
            not state["logged_in"]
            or not state["is_admin"]
        ):

            return jsonify({
                "error":
                "Access denied"
            }), 403

    conn = get_db()

    row = conn.execute(
        """
        SELECT *
        FROM excel_uploads
        WHERE id = ?
        """,
        (file_id,)
    ).fetchone()

    conn.close()

    if not row:

        return jsonify({
            "error":
            "File not found"
        }), 404

    path = Path(
        row["storage_path"]
    )

    if not path.exists():

        return jsonify({
            "error":
            "File no longer exists"
        }), 404

    return send_file(
        str(path),
        as_attachment=True,
        download_name=row["filename"]
    )


# ==========================================
# ERROR HANDLER
# ==========================================

@app.errorhandler(413)
def request_entity_too_large(error):

    return jsonify({
        "success": False,
        "error":
        "Die hochgeladenen Dateien überschreiten das erlaubte Limit von 100 MB."
    }), 413


# ==========================================
# MAIN
# ==========================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )