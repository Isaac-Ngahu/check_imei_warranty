
import csv
import time
from datetime import datetime
import os
from pathlib import Path
from zoneinfo import ZoneInfo
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


from openpyxl import Workbook
from flask import Flask, jsonify, request, send_file,render_template
import traceback
# ---------------- CONFIGURATION ----------------

SEARCH_URL = (
    "https://www.carlcare.com/sa_en/warranty-check/"
)

RESULT_URL = (
    "https://www.carlcare.com/sa_en/warranty-check/warranty/show/"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMEI_FILE = BASE_DIR / "imeis.txt"


ACTIVE_CSV = BASE_DIR / "active_imeis.csv"
RESULTS_CSV = BASE_DIR / "search_results.csv"

KENYA_TZ = ZoneInfo("Africa/Nairobi")

WAIT_TIMEOUT = 15
MIN_RESULT_WAIT = 3
MAX_RETRIES = 3

# Set to True if you want Chrome to run in the background.
HEADLESS = True

import math
import threading
import uuid

from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, jsonify, request,json

# ---------------- API / WORKER CONFIGURATION ----------------

app = Flask(__name__)

MAX_RETRIES = 3
MAX_WORKERS = 11

# Only allow one submitted batch to run at a time.
job_semaphore = threading.Semaphore(1)
MAX_BATCH_SIZE = 10000

executor = ThreadPoolExecutor(max_workers=MAX_WORKERS)

jobs = {}
jobs_lock = threading.Lock()
csv_lock = threading.Lock()
job_stop_events = {}
BATCHES_DIR = BASE_DIR / "batches"
BATCHES_DIR.mkdir(parents=True, exist_ok=True)

latest_job_id = None


#------------------------------db configs---------------------------#
import sqlite3
DB_FILE = os.path.join(BASE_DIR, "jobs.db")
def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn
def init_db():
    conn = get_db()
    conn.execute("""
    CREATE TABLE IF NOT EXISTS jobs (
    id TEXT PRIMARY KEY,
    status TEXT NOT NULL,
    total INTEGER NOT NULL,
    completed INTEGER DEFAULT 0,
    under_warranty INTEGER DEFAULT 0,
    errors INTEGER DEFAULT 0,
    worker_failures INTEGER DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    completed_at DATETIME
    )
    """)
    conn.commit()
    conn.close()

def create_job_record(job_id, total):
    conn = get_db()

    conn.execute("""
        INSERT INTO jobs (
            id,
            status,
            total
        )
        VALUES (?, ?, ?)
    """, (
        job_id,
        "queued",
        total
    ))

    conn.commit()
    conn.close()

def update_job_status(job_id, status):
    conn = get_db()

    if status in (
        "completed",
        "completed_with_errors",
        "stopped",
    ):
        conn.execute("""
            UPDATE jobs
            SET status = ?,
                completed_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (status, job_id))
    else:
        conn.execute("""
            UPDATE jobs
            SET status = ?
            WHERE id = ?
        """, (status, job_id))

    conn.commit()
    conn.close()




def update_job_progress(
    job_id,
    completed,
    under_warranty,
    errors
):
    conn = get_db()

    conn.execute("""
        UPDATE jobs
        SET completed = ?,
            under_warranty = ?,
            errors = ?
        WHERE id = ?
    """, (
        completed,
        under_warranty,
        errors,
        job_id,
    ))

    conn.commit()
    conn.close()

def find_active_job():
    conn = get_db()

    row = conn.execute("""
        SELECT id
        FROM jobs
        WHERE status IN ('queued', 'running')
        LIMIT 1
    """).fetchone()

    conn.close()

    return row["id"] if row else None
#-----------------------------end db configs--------------------------#

@app.route("/", methods=["GET"])
def warranty_dashboard():
    return render_template("index.html")

def get_batch_paths(job_id):
    """Return CSV paths belonging to one batch."""
    batch_dir = BATCHES_DIR / job_id
    batch_dir.mkdir(parents=True, exist_ok=True)

    return (
        batch_dir / "active_imeis.csv",
        batch_dir / "search_results.csv",
    )


class RateLimitedError(Exception):
    """The website appears to be throttling or blocking requests."""
    pass


def get_document_status(driver):
    """Return the latest HTTP status for the warranty page's main document."""
    latest_status = None

    for entry in driver.get_log("performance"):
        try:
            message = json.loads(entry["message"])["message"]

            if message.get("method") != "Network.responseReceived":
                continue

            params = message.get("params", {})
            response = params.get("response", {})

            if (
                params.get("type") == "Document"
                and "carlcare.com" in response.get("url", "")
            ):
                latest_status = int(response["status"])

        except (ValueError, KeyError, TypeError):
            continue

    return latest_status


def detect_rate_limit(driver):
    """Raise RateLimitedError when a throttling response is detected."""
    status = get_document_status(driver)

    if status == 429:
        raise RateLimitedError("HTTP 429: Too Many Requests")

    if status == 503:
        raise RateLimitedError(
            "HTTP 503: Service temporarily unavailable"
        )

    # Some sites display a block page with HTTP 200 or 403.
    try:
        page_text = driver.find_element("tag name", "body").text.lower()
    except WebDriverException:
        return

    blocked_phrases = (
        "too many requests",
        "rate limit exceeded",
        "temporarily blocked",
    )

    if any(phrase in page_text for phrase in blocked_phrases):
        raise RateLimitedError(
            f"Possible website rate limit; HTTP status={status}"
        )

    if status == 403:
        raise RateLimitedError(
            "HTTP 403: Access forbidden; check whether the site is blocking access"
        )

def create_batch_csvs(job_id):
    """Initialize separate CSV files for a batch."""
    active_csv, results_csv = get_batch_paths(job_id)

    if not active_csv.exists():
        with open(active_csv, "w", newline="", encoding="utf-8-sig") as file:
            csv.writer(file).writerow(["IMEI", "Processed At"])

    if not results_csv.exists():
        with open(results_csv, "w", newline="", encoding="utf-8-sig") as file:
            csv.writer(file).writerow(
                ["IMEI", "Warranty Status", "Processed At"]
            )


def save_result_safely(imei, status, job_id):
    processed_at = datetime.now(KENYA_TZ).strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    with csv_lock:
        record_result(
            imei,
            status,
            processed_at,
            job_id=job_id,
        )


# def process_chunk(job_id, imeis):
#     """Process one chunk using a single reusable Chrome browser."""

#     driver = None
#     completed = 0
#     active = 0
#     errors = 0
#     status = "Error"

#     try:
#         driver = create_browser()

#         for imei in imeis:
#             status = "Error"
#             try:
#                 status = check_imei(driver, imei)
#                 save_result_safely(imei, status,job_id)

#                 if status == "Under warranty":
#                     active += 1

#                 if status in ("Unknown", "Error"):
#                     errors += 1

#             except Exception as exc:
#                 status = "Error"
#                 print(f"[WORKER ERROR] IMEI {imei}: {exc}")
#                 traceback.print_exc()
#                 errors += 1

#                 try:
#                     save_result_safely(imei, "Error", job_id)
#                 except Exception as save_exc:
#                     print(f"[CSV ERROR] IMEI {imei}: {save_exc}")

#             finally:
#                 completed += 1

#                 with jobs_lock:
#                     job = jobs[job_id]
#                     job["completed"] += 1
#                     job["under_warranty"] += int(status == "Under warranty")
#                     job["errors"] += int(status in ("Unknown", "Error"))

#     except Exception as exc:
#         print(f"[WORKER FAILED] Job {job_id}: {exc}")

#         with jobs_lock:
#             jobs[job_id]["worker_failures"] += 1
#             jobs[job_id]["messages"].append(str(exc))

#     finally:
#         if driver is not None:
#             try:
#                 driver.quit()
#             except WebDriverException:
#                 pass


def process_chunk(job_id, imeis):
    """Process a chunk and recreate Chrome if its session fails."""

    driver = None
    errors = 0
    active = 0

    try:
        stop_event = job_stop_events[job_id]
        for imei in imeis:
            if stop_event.is_set():
                print(
                    f"[JOB {job_id}] Stop requested. "
                    "Worker will not process more IMEIs."
                )
                break
            status = "Error"

            print(
                f"[JOB {job_id}] Processing IMEI {imei}"
            )

            # Initial attempt plus one attempt with a replacement browser.
            for browser_attempt in range(1, 3):
                try:
                    if driver is None:
                        print(
                            f"[BROWSER START] Worker creating browser "
                            f"for IMEI {imei}"
                        )
                        driver = create_browser()

                    status = check_imei(driver, imei)
                    break

                except WebDriverException as exc:
                    print(
                        f"[BROWSER RESTART] IMEI {imei}, "
                        f"attempt {browser_attempt}/2: {exc}"
                    )

                    if driver is not None:
                        try:
                            driver.quit()
                        except WebDriverException:
                            pass

                        driver = None

                    if browser_attempt < 2:
                        time.sleep(3)

                except Exception as exc:
                    print(f"[WORKER ERROR] IMEI {imei}: {exc}")
                    traceback.print_exc()
                    status = "Error"
                    break

            if status in ("Unknown", "Error"):
                errors += 1

            if status == "Under warranty":
                active += 1

            try:
                save_result_safely(imei, status, job_id)
            except Exception as exc:
                print(f"[CSV ERROR] IMEI {imei}: {exc}")

            with jobs_lock:
                job = jobs[job_id]
                job["completed"] += 1
                job["under_warranty"] += int(status == "Under warranty")
                job["errors"] += int(status in ("Unknown", "Error"))
                update_job_progress(
                    job_id,
                    job["completed"],
                    job["under_warranty"],
                    job["errors"]
                )

            # Be gentle with the website between IMEIs.
            if not stop_event.is_set():
                time.sleep(2)

    except Exception as exc:
        print(f"[WORKER FAILED] Job {job_id}: {exc}")
        traceback.print_exc()

        with jobs_lock:
            jobs[job_id]["worker_failures"] += 1
            jobs[job_id]["messages"].append(str(exc))

    finally:
        if driver is not None:
            try:
                driver.quit()
            except WebDriverException:
                pass



# def run_job(job_id, imeis):
#     """Split a batch across a bounded number of browser workers."""
#     print(
#         f"[JOB {job_id}] Starting batch with "
#         f"{len(imeis)} IMEIs"
#     )

#     chunks = [
#         imeis[i::MAX_WORKERS]
#         for i in range(MAX_WORKERS)
#     ]
#     chunks = [chunk for chunk in chunks if chunk]

#     with jobs_lock:
#         jobs[job_id] = {}
#         update_job_status(job_id, "running")
#         jobs[job_id]["workers"] = len(chunks)

#     futures = [
#         executor.submit(process_chunk, job_id, chunk)
#         for chunk in chunks
#     ]

#     for future in as_completed(futures):
#         # process_chunk handles its own exceptions.
#         future.result()

#     with jobs_lock:
#         job = jobs[job_id]
#         final_status = (
#             "completed"
#             if job["completed"] == job["total"]
#             else "completed_with_errors"
#         )

#         job["status"] = final_status

#         update_job_status(
#             job_id,
#             final_status
#         )


def run_job(job_id, imeis):
    """Run a batch with graceful stop support."""

    stop_event = job_stop_events[job_id]

    print(
        f"[JOB {job_id}] Starting batch with "
        f"{len(imeis)} IMEIs"
    )

    # The job may have been stopped while waiting for the semaphore.
    if stop_event.is_set():
        with jobs_lock:
            jobs[job_id]["status"] = "stopped"
            update_job_status(job_id, "stopped")
        return

    chunks = [
        imeis[i::MAX_WORKERS]
        for i in range(MAX_WORKERS)
    ]
    chunks = [chunk for chunk in chunks if chunk]

    with jobs_lock:
        # Do not replace jobs[job_id]; preserve its counters.
        jobs[job_id]["status"] = "running"
        jobs[job_id]["workers"] = len(chunks)
        update_job_status(job_id, "running")

    futures = [
        executor.submit(process_chunk, job_id, chunk)
        for chunk in chunks
    ]

    for future in as_completed(futures):
        future.result()

    with jobs_lock:
        job = jobs[job_id]

        if stop_event.is_set():
            final_status = "stopped"
        elif job["completed"] == job["total"]:
            final_status = "completed"
        else:
            final_status = "completed_with_errors"

        job["status"] = final_status
        update_job_status(job_id, final_status)

        print(
            f"[JOB {job_id}] Finished with status "
            f"{final_status}. Processed "
            f"{job['completed']}/{job['total']} IMEIs."
        )


def run_job_serialized(job_id, imeis):
    """Allow only one batch to use the browser workers at a time."""
    with job_semaphore:
        run_job(job_id, imeis)


def recover_interrupted_jobs():
    conn = get_db()

    conn.execute("""
        UPDATE jobs
        SET status = 'interrupted',
            completed_at = CURRENT_TIMESTAMP
        WHERE status IN ('queued', 'running')
    """)

    conn.commit()
    conn.close()

# def find_active_job():
#     """Return an existing queued or running job, if any."""
#     with jobs_lock:
#         for job_id, job in jobs.items():
#             if job["status"] in ("queued", "running"):
#                 return job_id

#     return None
# ---------------- API ENDPOINTS ----------------
@app.route("/api/jobs", methods=["GET"])
def list_jobs():

    conn = get_db()

    rows = conn.execute("""
        SELECT *
        FROM jobs
        ORDER BY created_at DESC
    """).fetchall()

    conn.close()

    return jsonify([
        dict(row)
        for row in rows
    ])

@app.route("/api/warranty-check", methods=["POST"])
def start_warranty_check():
    """Submit a batch of IMEIs for background processing."""
    global latest_job_id

    data = request.get_json(silent=True)

    print(type(data))

    if not isinstance(data, dict):
        return jsonify({
            "error": "Send a JSON object containing an imeis array."
        }), 400

    imeis = data.get("imeis")

    if not isinstance(imeis, list) or not imeis:
        return jsonify({
            "error": "imeis must be a non-empty array."
        }), 400

    if len(imeis) > MAX_BATCH_SIZE:
        return jsonify({
            "error": f"Maximum batch size is {MAX_BATCH_SIZE} IMEIs."
        }), 400

    # Normalize and deduplicate while preserving order.
    imeis = list(dict.fromkeys(
        str(value).strip()
        for value in imeis
        if value is not None and str(value).strip()
    ))

    if not imeis:
        return jsonify({"error": "No valid IMEIs supplied."}), 400

    active_job_id = find_active_job()

    if active_job_id:
        return jsonify({
            "error": "Another warranty check is already in progress.",
            "active_job_id": active_job_id,
            "status_url": f"/api/warranty-check/{active_job_id}",
        }), 409
    
    job_id = str(uuid.uuid4())
    create_batch_csvs(job_id)

    with jobs_lock:
        latest_job_id = job_id

        jobs[job_id] = {
            "status": "queued",
            "total": len(imeis),
            "completed": 0,
            "under_warranty": 0,
            "errors": 0,
            "worker_failures": 0,
            "workers": 0,
            "messages": [],
        }

        job_stop_events[job_id] = threading.Event()

        create_job_record(job_id, len(imeis))

    # Run the batch outside the HTTP request.
    # executor.submit(run_job, job_id, imeis)

    threading.Thread(
    target=run_job_serialized,
    args=(job_id, imeis),
    daemon=True,
    ).start()

    return jsonify({
        "message": "Warranty check job accepted.",
        "job_id": job_id,
        "total_imeis": len(imeis),
        "status_url": f"/api/warranty-check/{job_id}",
    }), 202

@app.route("/api/warranty-check/<job_id>/stop", methods=["POST"])
def stop_warranty_job(job_id):
    with jobs_lock:
        conn = get_db()
        row = conn.execute(
            "SELECT status FROM jobs WHERE id = ?",
            (job_id,)
        ).fetchone()
        conn.close()

        if row is None:
            return jsonify({"error": "Job not found."}), 404

        if row["status"] not in ("queued", "running"):
            return jsonify({
                "message": "Job is no longer running.",
                "status": row["status"]
            }), 409

        # SQLite may contain jobs from a previous Flask process.
        job = jobs.get(job_id)
        stop_event = job_stop_events.get(job_id)

        if job is None or stop_event is None:
            return jsonify({
                "error": (
                    "This job is in the database but is not registered "
                    "in the current Flask process. It may have been "
                    "interrupted by an application restart."
                ),
                "job_id": job_id,
                "database_status": row["status"]
            }), 409

        stop_event.set()

        return jsonify({
            "message": (
                "Stop requested. Workers will finish their current "
                "IMEI checks before stopping."
            ),
            "job_id": job_id,
            "status": "stopping"
        }), 202

# @app.route("/api/warranty-check/<job_id>/stop", methods=["POST"])
# def stop_warranty_job(job_id):
#     with jobs_lock:
#         # Check persistent job status.
#         conn = get_db()
#         row = conn.execute(
#             "SELECT status FROM jobs WHERE id = ?",
#             (job_id,)
#         ).fetchone()
#         conn.close()

#         if row is None:
#             return jsonify({
#                 "error": "Job not found."
#             }), 404

#         if row["status"] not in ("queued", "running"):
#             return jsonify({
#                 "message": "Job is no longer running.",
#                 "status": row["status"]
#             }), 409

#         stop_event = job_stop_events.get(job_id)

#         if stop_event is None:
#             return jsonify({
#                 "error": "Stop signal is unavailable for this job."
#             }), 409

#         stop_event.set()

#         return jsonify({
#             "message": "Stop requested. The current IMEI will be allowed to finish.",
#             "job_id": job_id,
#             "status": "stopping"
#         }), 202

# @app.route("/api/warranty-check/<job_id>", methods=["GET"])
# def get_warranty_job(job_id):
#     """Return progress for a submitted batch."""

#     with jobs_lock:
#         job = jobs.get(job_id)

#         if job is None:
#             return jsonify({"error": "Job not found."}), 404

#         result = dict(job)

#     result["progress_percent"] = round(
#         result["completed"] / result["total"] * 100, 2
#     ) if result["total"] else 0

#     return jsonify({
#         "job_id": job_id,
#         **result,
#         "active_csv": str(ACTIVE_CSV),
#         "results_csv": str(RESULTS_CSV),
#     })

# @app.route("/api/warranty-check/<job_id>", methods=["GET"])
# def get_warranty_job(job_id):
#     """Return progress for a submitted batch."""

#     with jobs_lock:
#         job = jobs.get(job_id)

#         if job is None:
#             return jsonify({
#                 "error": "Job not found."
#             }), 404

#         result = dict(job)

#     result["progress_percent"] = round(
#         result["completed"] / result["total"] * 100, 2
#     ) if result["total"] else 0

#     active_csv, results_csv = get_batch_paths(job_id)

#     return jsonify({
#         "job_id": job_id,
#         **result,
#         "active_csv": str(active_csv),
#         "results_csv": str(results_csv),
#     })

@app.route("/api/warranty-check/<job_id>", methods=["GET"])
def get_warranty_job(job_id):
    conn = get_db()

    print("DB_FILE:", DB_FILE)
    print("DB exists:", os.path.exists(DB_FILE))

    tables = conn.execute("""
        SELECT name
        FROM sqlite_master
        WHERE type = 'table'
    """).fetchall()

    print("Tables:", [row["name"] for row in tables])

    job = conn.execute("""
    SELECT *
    FROM jobs
    WHERE id = ?
    """, (job_id,)).fetchone()
    
    conn.close()
    if job is None:
        return jsonify({
        "error": "Job not found"
        }), 404

    result = dict(job)

    result["progress_percent"] = round(
    result["completed"] / result["total"] * 100,
    2
    ) if result["total"] else 0

    active_csv, results_csv = get_batch_paths(job_id)

    return jsonify({
    **result,
    "active_csv": str(active_csv),
    "results_csv": str(results_csv)
    })


# ---------------- EXCEL HELPERS ----------------

# def create_workbook_if_missing():
#     """Create the workbook and its worksheets if necessary."""

#     if EXCEL_FILE.exists():
#         wb = load_workbook(EXCEL_FILE)
#     else:
#         wb = Workbook()

#         active_ws = wb.active
#         active_ws.title = "Active IMEIs"
#         active_ws.append(["IMEI"])

#         results_ws = wb.create_sheet("Search Results")
#         results_ws.append(["IMEI", "Warranty Status"])

#         wb.save(EXCEL_FILE)

#     if "Active IMEIs" not in wb.sheetnames:
#         ws = wb.create_sheet("Active IMEIs")
#         ws.append(["IMEI"])

#     if "Search Results" not in wb.sheetnames:
#         ws = wb.create_sheet("Search Results")
#         ws.append(["IMEI", "Warranty Status"])

#     # Make sure IMEIs are stored as text.
#     for ws in wb.worksheets:
#         for cell in ws["A"]:
#             cell.number_format = "@"

#     wb.save(EXCEL_FILE)
#     return wb


# def refresh_excel_table(ws, table_name, headers):
#     """Create or update an Excel table to cover all populated rows."""

#     if table_name in ws.tables:
#         del ws.tables[table_name]

#     last_row = max(ws.max_row, 1)
#     last_col = len(headers)

#     # Excel tables require at least a header and a data row.
#     if last_row < 2:
#         ws.append([""] * last_col)
#         last_row = 2

#     table = Table(
#         displayName=table_name,
#         ref=f"A1:{chr(64 + last_col)}{last_row}"
#     )

#     table.tableStyleInfo = TableStyleInfo(
#         name="TableStyleMedium2",
#         showFirstColumn=False,
#         showLastColumn=False,
#         showRowStripes=True,
#         showColumnStripes=False,
#     )

#     ws.add_table(table)


# def record_result(imei, status):
#     """Save every result and add matching IMEIs to the active table."""

#     wb = load_workbook(EXCEL_FILE)

#     active_ws = wb["Active IMEIs"]
#     results_ws = wb["Search Results"]

#     # Record the result only once per IMEI.
#     existing_results = {
#         str(row[0].value).strip()
#         for row in results_ws.iter_rows(min_row=2, max_col=1)
#         if row[0].value is not None
#     }

#     if imei not in existing_results:
#         results_ws.append([imei, status])

#     # Add only confirmed under-warranty IMEIs.
#     existing_active = {
#         str(row[0].value).strip()
#         for row in active_ws.iter_rows(min_row=2, max_col=1)
#         if row[0].value is not None
#     }

#     if status == "Under warranty" and imei not in existing_active:
#         active_ws.append([imei])

#     refresh_excel_table(
#         active_ws,
#         "active_imeis",
#         ["IMEI"]
#     )

#     refresh_excel_table(
#         results_ws,
#         "search_results",
#         ["IMEI", "Warranty Status"]
#     )

#     wb.save(EXCEL_FILE)

# ---------------- CSV HELPERS ----------------

def create_csv_if_missing():
    """Create CSV files with headers if they do not exist."""

    if not ACTIVE_CSV.exists():
        with open(ACTIVE_CSV, "w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(["IMEI", "Processed At"])

    if not RESULTS_CSV.exists():
        with open(RESULTS_CSV, "w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file)
            writer.writerow(["IMEI", "Warranty Status", "Processed At"])


def read_csv_column(filepath):
    """Read existing IMEIs from the first CSV column."""

    existing = set()

    if not filepath.exists():
        return existing

    with open(filepath, "r", newline="", encoding="utf-8-sig") as file:
        reader = csv.reader(file)
        next(reader, None)  # Skip header

        for row in reader:
            if row and row[0].strip():
                existing.add(row[0].strip())

    return existing


# def record_result(imei, status, processed_at):
#     """Record the warranty result and track active IMEIs."""

#     existing_results = read_csv_column(RESULTS_CSV)

#     # Save the first result for each IMEI only.
#     if imei not in existing_results:
#         with open(
#             RESULTS_CSV, "a", newline="", encoding="utf-8-sig"
#         ) as file:
#             writer = csv.writer(file)
#             writer.writerow([imei, status, processed_at])

#     existing_active = read_csv_column(ACTIVE_CSV)

#     # Add only confirmed under-warranty IMEIs.
#     if status == "Under warranty" and imei not in existing_active:
#         with open(
#             ACTIVE_CSV, "a", newline="", encoding="utf-8-sig"
#         ) as file:
#             writer = csv.writer(file)
#             writer.writerow([imei, processed_at])

def record_result(imei, status, processed_at, job_id=None):
    """Save results to a batch-specific CSV, or global CSV if no batch."""

    if job_id:
        active_csv, results_csv = get_batch_paths(job_id)
    else:
        active_csv, results_csv = ACTIVE_CSV, RESULTS_CSV

    existing_results = read_csv_column(results_csv)

    if imei not in existing_results:
        with open(
            results_csv, "a", newline="", encoding="utf-8-sig"
        ) as file:
            csv.writer(file).writerow([imei, status, processed_at])

    existing_active = read_csv_column(active_csv)

    if status == "Under warranty" and imei not in existing_active:
        with open(
            active_csv, "a", newline="", encoding="utf-8-sig"
        ) as file:
            csv.writer(file).writerow([imei, processed_at])



# ---------------- BROWSER SETUP ----------------

def create_browser():
    options = webdriver.ChromeOptions()


    options.add_argument("--headless=new")

    options.add_argument("--start-maximized")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-notifications")
    
    # Capture network responses, including HTTP status codes
    options.set_capability(
        "goog:loggingPrefs",
        {"performance": "ALL"}
    )

    driver = webdriver.Chrome(options=options)

    return driver


# ---------------- WARRANTY CHECK ----------------

# def check_imei(driver, imei):
#     """Search one IMEI and return its warranty status."""

#     wait = WebDriverWait(driver, WAIT_TIMEOUT)

#     try:
#         # Start every search from the page containing the IMEI field.
#         driver.get(SEARCH_URL)

#         imei_input = wait.until(
#             EC.element_to_be_clickable(
#                 (By.CSS_SELECTOR, 'input[placeholder="IMEI"]')
#             )
#         )

#         imei_input.clear()
#         imei_input.send_keys(imei)

#         search_button = wait.until(
#             EC.element_to_be_clickable(
#                 (
#                     By.CSS_SELECTOR,
#                     "div.exp_check_input_right"
#                 )
#             )
#         )

#         search_button.click()

#         # Always wait at least one second after clicking.
#     #     time.sleep(MIN_RESULT_WAIT)

#     #     # Wait for the warranty result field to appear.
#     #     wait.until(
#     #         EC.presence_of_element_located(
#     #             (
#     #                 By.XPATH,
#     #                 "//h4[normalize-space()='Warranty status']"
#     #             )
#     #         )
#     #     )

#     #     # Locate the paragraph belonging to Warranty status.
#     #     status_element = wait.until(
#     #         EC.presence_of_element_located(
#     #             (
#     #                 By.XPATH,
#     #                 "//h4[normalize-space()='Warranty status']"
#     #                 "/following-sibling::p[1]"
#     #             )
#     #         )
#     #     )

#     #     status = status_element.text.strip()

#     #     if status:
#     #         return status

#     #     return "Unknown"

#     # except TimeoutException:
#     #     print(
#     #         f"[TIMEOUT] Could not retrieve the warranty result "
#     #         f"for IMEI {imei}"
#     #     )
#     #     return "Unknown"

#     # except WebDriverException as exc:
#     #     print(f"[BROWSER ERROR] IMEI {imei}: {exc}")
#     #     return "Error"
    
#         # Wait for the result and identify which selector appears.
#         time.sleep(MIN_RESULT_WAIT)

#         try:
#             wait.until(
#                 EC.presence_of_element_located(
#                     (
#                         By.XPATH,
#                         "//h4[normalize-space()='Warranty status']"
#                     )
#                 )
#             )
#         except TimeoutException:
#             print("[DEBUG] Warranty heading was not found.")
#             print(f"[DEBUG] Current URL: {driver.current_url}")
#             print(f"[DEBUG] Page title: {driver.title}")

#             # Preserve the browser state for inspection.
#             driver.save_screenshot(f"timeout_{imei}.png")

#             Path(f"timeout_{imei}.html").write_text(
#                 driver.page_source,
#                 encoding="utf-8"
#             )

#             print(f"[DEBUG] Saved timeout_{imei}.png")
#             print(f"[DEBUG] Saved timeout_{imei}.html")
#             print("[DEBUG] Visible page text:")
#             print(driver.find_element(By.TAG_NAME, "body").text[:3000])

#             return "Unknown"

#         status_element = wait.until(
#             EC.presence_of_element_located(
#                 (
#                     By.XPATH,
#                     "//h4[normalize-space()='Warranty status']"
#                     "/following-sibling::p[1]"
#                 )
#             )
#         )

#         status = status_element.text.strip()
#         return status if status else "Unknown"


# ---------------- WARRANTY CHECK ----------------

# 


def check_imei(driver, imei):
    """Check an IMEI with retries and more robust result detection."""

    def extract_status(browser):
        try:
            body_text = browser.find_element(
                By.TAG_NAME, "body"
            ).text

            if "Device activation is in progress" in body_text:
                return "Device activation is in progress"

            headings = browser.find_elements(
                By.XPATH,
                "//*[normalize-space()='Warranty status']"
            )

            for heading in headings:
                siblings = heading.find_elements(
                    By.XPATH, "./following-sibling::*[1]"
                )

                if siblings:
                    status = siblings[0].text.strip()
                    if status:
                        return status

                # Some page layouts place the status inside a parent.
                parent_text = heading.find_element(
                    By.XPATH, ".."
                ).text.strip()

                lines = [
                    line.strip()
                    for line in parent_text.splitlines()
                    if line.strip()
                ]

                for index, line in enumerate(lines):
                    if (
                        line.lower() == "warranty status"
                        and index + 1 < len(lines)
                    ):
                        return lines[index + 1]

            # Handle known messages that may appear without the heading.
            known_statuses = [
                "Under warranty",
                "Out of warranty",
                "Warranty expired",
                "Warranty is expired",
            ]

            for known_status in known_statuses:
                if known_status.lower() in body_text.lower():
                    return known_status

            return None

        except WebDriverException:
            return None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            print(
                f"[ATTEMPT {attempt}/{MAX_RETRIES}] "
                f"Checking IMEI {imei}"
            )

            try:
                driver.get_log("performance")
            except WebDriverException:
                pass

            driver.get(SEARCH_URL)

            # Check the page response before submitting the IMEI.
            detect_rate_limit(driver)
            wait = WebDriverWait(driver, WAIT_TIMEOUT)

            imei_input = wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, 'input[placeholder="IMEI"]')
                )
            )

            imei_input.clear()
            imei_input.send_keys(imei)

            search_button = wait.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "div.exp_check_input_right")
                )
            )

            search_button.click()

            # Wait for a recognizable result instead of sleeping
            # for a fixed number of seconds.
            status = WebDriverWait(
                driver, WAIT_TIMEOUT
            ).until(extract_status)

            print(f"[RESULT] IMEI {imei}: {status}")
            return status

        except TimeoutException:
    # A timeout can mean a slow page or a dead browser.
            try:
                driver.execute_script("return 1")
            except WebDriverException:
                raise

            print(
                f"[TIMEOUT] IMEI {imei}, "
                f"attempt {attempt}/{MAX_RETRIES}: "
                "No recognizable result appeared."
            )

        except WebDriverException as exc:
            print(
                f"[BROWSER SESSION FAILED] IMEI {imei}: {exc}"
            )
            raise

        except RateLimitedError as exc:
            print(
                f"[RATE LIMIT ERROR] IMEI {imei}: {exc}"
            )
            raise

        except Exception as exc:
            print(
                f"[UNEXPECTED ERROR] IMEI {imei}, "
                f"attempt {attempt}/{MAX_RETRIES}: {exc}"
            )
            traceback.print_exc()

        if attempt < MAX_RETRIES:
            time.sleep(2)

    return "Unknown"


# ---------------- WARRANTY CHECK ----------------

# def check_imei(driver, imei):
#     """Check an IMEI, recognize different result states, and retry failures."""

#     result_url_part = "/warranty/show/"

#     def get_result_status():
#         """Return the detected result, or None while the page is loading."""

#         body_text = driver.find_element(
#             By.TAG_NAME, "body"
#         ).text

#         # The result page has loaded, but activation is not complete.
#         if "Device activation is in progress" in body_text:
#             return "Device activation is in progress"

#         # Look for the normal warranty status heading.
#         headings = driver.find_elements(
#             By.XPATH,
#             "//*[normalize-space()='Warranty status']"
#         )

#         for heading in headings:
#             # Try the next sibling element first.
#             siblings = heading.find_elements(
#                 By.XPATH, "./following-sibling::*[1]"
#             )

#             if siblings:
#                 status = siblings[0].text.strip()
#                 if status:
#                     return status

#             # Also support a status nested inside a parent container.
#             try:
#                 parent = heading.find_element(By.XPATH, "..")
#                 text = parent.text.strip()
#                 lines = [
#                     line.strip()
#                     for line in text.splitlines()
#                     if line.strip()
#                 ]

#                 for index, line in enumerate(lines):
#                     if line == "Warranty status" and index + 1 < len(lines):
#                         return lines[index + 1]
#             except WebDriverException:
#                 pass

#         # The result page may have loaded with a different status layout.
#         if "Your warranty info" in body_text:
#             known_messages = [
#                 "Under warranty",
#                 "Out of warranty",
#                 "Warranty expired",
#                 "Warranty is expired",
#                 "Warranty status unavailable",
#             ]

#             for message in known_messages:
#                 if message.lower() in body_text.lower():
#                     return message

#             # Do not misclassify a loaded result as a timeout.
#             return "Result loaded - status not identified"

#         return None

#     for attempt in range(1, MAX_RETRIES + 1):
#         try:
#             print(
#                 f"[ATTEMPT {attempt}/{MAX_RETRIES}] "
#                 f"Checking IMEI {imei}"
#             )

#             driver.get(SEARCH_URL)
#             wait = WebDriverWait(driver, WAIT_TIMEOUT)

#             imei_input = wait.until(
#                 EC.element_to_be_clickable(
#                     (By.CSS_SELECTOR, 'input[placeholder="IMEI"]')
#                 )
#             )

#             imei_input.clear()
#             imei_input.send_keys(imei)

#             search_button = wait.until(
#                 EC.element_to_be_clickable(
#                     (By.CSS_SELECTOR, "div.exp_check_input_right")
#                 )
#             )

#             search_button.click()

#             time.sleep(MIN_RESULT_WAIT)

#             # Wait until a recognizable result appears.
#             def result_ready(browser):
#                 try:
#                     if result_url_part not in browser.current_url:
#                         return False

#                     return get_result_status() is not None
#                 except WebDriverException:
#                     return False

#             wait.until(result_ready)

#             status = get_result_status()

#             if status:
#                 return status

#             print("[RETRY] Result page loaded, status not yet available.")

#         except TimeoutException:
#             print(
#                 f"[TIMEOUT] Attempt {attempt}/{MAX_RETRIES} "
#                 f"for IMEI {imei}"
#             )

#         except WebDriverException as exc:
#             print(
#                 f"[BROWSER ERROR] Attempt {attempt}/{MAX_RETRIES}: {exc}"
#             )

#         # Capture evidence after the final failed attempt.
#         if attempt == MAX_RETRIES:
#             try:
#                 driver.save_screenshot(f"timeout_{imei}.png")

#                 Path(f"timeout_{imei}.html").write_text(
#                     driver.page_source,
#                     encoding="utf-8"
#                 )

#                 print(f"[DEBUG] Saved timeout_{imei}.png and HTML")
#                 print(
#                     "[DEBUG] Current URL:",
#                     driver.current_url
#                 )
#                 print(
#                     "[DEBUG] Page text:",
#                     driver.find_element(
#                         By.TAG_NAME, "body"
#                     ).text[:3000]
#                 )
#             except WebDriverException as exc:
#                 print(f"[DEBUG] Could not save diagnostics: {exc}")

#             return "Unknown"

#         time.sleep(2)

#     return "Unknown"


# @app.route("/api/warranty-check/active-imeis", methods=["GET"])
# def get_active_imeis():
#     """Return currently recorded active IMEIs while jobs are running."""

#     create_csv_if_missing()

#     with csv_lock:
#         active_imeis = []

#         with open(
#             ACTIVE_CSV,
#             "r",
#             newline="",
#             encoding="utf-8-sig"
#         ) as file:
#             reader = csv.DictReader(file)

#             for row in reader:
#                 imei = (row.get("IMEI") or "").strip()

#                 if imei:
#                     active_imeis.append({
#                         "imei": imei,
#                         "processed_at": row.get("Processed At", "")
#                     })

#     return jsonify({
#         "count": len(active_imeis),
#         "imeis": active_imeis
#     }), 200


# @app.route("/api/warranty-check/active-imeis", methods=["GET"])
# def get_active_imeis():
#     """Generate and return an Excel file containing active IMEIs."""

#     create_csv_if_missing()

#     excel_file = BASE_DIR / "active_imeis.xlsx"

#     with csv_lock:
#         active_imeis = read_csv_column(ACTIVE_CSV)

#         # Create a new workbook with only the IMEI column.
#         workbook = Workbook()
#         worksheet = workbook.active
#         worksheet.title = "Active IMEIs"

#         worksheet.append(["IMEI"])

#         for imei in sorted(active_imeis):
#             worksheet.append([str(imei)])
#             worksheet.cell(
#                 row=worksheet.max_row,
#                 column=1
#             ).number_format = "@"

#         workbook.save(excel_file)
#         workbook.close()

#     return send_file(
#         excel_file,
#         as_attachment=True,
#         download_name="active_imeis.xlsx",
#         mimetype=(
#             "application/vnd.openxmlformats-officedocument."
#             "spreadsheetml.sheet"
#         )
#     )

@app.route("/api/warranty-check/active-imeis", methods=["GET"])
def get_active_imeis():
    """Download active IMEIs from the latest batch, or a specified batch."""

    requested_job_id = request.args.get("job_id")

    with jobs_lock:
        job_id = requested_job_id or latest_job_id

        if not job_id or job_id not in jobs:
            return jsonify({
                "error": "No valid warranty batch found."
            }), 404

    active_csv, _ = get_batch_paths(job_id)

    if not active_csv.exists():
        return jsonify({
            "error": "Active IMEI file not found for this batch."
        }), 404

    excel_file = BATCHES_DIR / job_id / "active_imeis.xlsx"

    with csv_lock:
        active_imeis = read_csv_column(active_csv)

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Active IMEIs"
        worksheet.append(["IMEI"])

        for imei in sorted(active_imeis):
            worksheet.append([str(imei)])
            worksheet.cell(
                row=worksheet.max_row,
                column=1
            ).number_format = "@"

        workbook.save(excel_file)
        workbook.close()

    return send_file(
        excel_file,
        as_attachment=True,
        download_name=f"active_imeis_{job_id}.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    ) #in excel format

# @app.route("/api/warranty-check/active-imeis", methods=["GET"])
# def get_active_imeis():
#     """Download active IMEIs from the latest batch, or a specified batch."""

#     requested_job_id = request.args.get("job_id")

#     with jobs_lock:
#         job_id = requested_job_id or latest_job_id

#         if not job_id or job_id not in jobs:
#             return jsonify({
#                 "error": "No valid warranty batch found."
#             }), 404

#     active_csv, _ = get_batch_paths(job_id)

#     if not active_csv.exists():
#         return jsonify({
#             "error": "Active IMEI file not found for this batch."
#         }), 404

#     csv_file = BATCHES_DIR / job_id / "active_imeis_download.csv"

#     with csv_lock:
#         active_imeis = read_csv_column(active_csv)

#         with open(csv_file, "w", newline="", encoding="utf-8-sig") as file:
#             writer = csv.writer(file)
#             writer.writerow(["IMEI"])

#             for imei in sorted(active_imeis):
#                 writer.writerow([str(imei)])

#     return send_file(
#         csv_file,
#         as_attachment=True,
#         download_name=f"active_imeis_{job_id}.csv",
#         mimetype="text/csv",
#     ) #in csv format



# ---------------- MAIN ----------------

def main():
    if not IMEI_FILE.exists():
        print(f"IMEI input file not found: {IMEI_FILE}")
        return

    # Read IMEIs as strings to preserve leading zeroes.
    imeis = [
        line.strip()
        for line in IMEI_FILE.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]

    # Remove duplicates while preserving input order.
    imeis = list(dict.fromkeys(imeis))

    if not imeis:
        print("No IMEIs found in imeis.txt.")
        return

    create_csv_if_missing()

    driver = None
    active_count = 0

    try:
        driver = create_browser()

        for index, imei in enumerate(imeis, start=1):
            print(f"\n[{index}/{len(imeis)}] Checking IMEI: {imei}")

            status = check_imei(driver, imei)

            print(f"Warranty status: {status}")

            process_time = datetime.now()

            # Save each result immediately.
            record_result(imei, status,process_time)

            if status == "Under warranty":
                active_count += 1
                print(f"[ACTIVE] {imei} added to active_imeis.")

            # The next iteration navigates back to SEARCH_URL.
            time.sleep(0.5)

    except KeyboardInterrupt:
        print("\nStopped by user. Saved results are retained.")

    except WebDriverException as exc:
        print(f"Browser initialization or execution failed: {exc}")

    finally:
        if driver is not None:
            driver.quit()

        print("\nWarranty checks finished.")
        print(f"IMEIs processed: {len(imeis)}")
        print(f"Under warranty during this run: {active_count}")
        print(f"Excel file: {RESULTS_CSV}")


# if __name__ == "__main__":
#     main()


if __name__ == "__main__":
    create_csv_if_missing()
    init_db()
    recover_interrupted_jobs()
    app.run(
        host="0.0.0.0",
        port=5006,
        threaded=True,
        use_reloader=False,
        debug=True
    )

