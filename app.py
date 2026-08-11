#!/usr/bin/env python3
"""
Transcript Conversion System — Flask Backend
Run:  python app.py
Open: http://localhost:5000
"""
import os, re, sys, json, uuid, traceback, threading, requests, time
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, request, jsonify, send_file, render_template, send_from_directory

sys.path.insert(0, os.path.dirname(__file__))
from pdf_to_excel_engine import (
    parse_pdf, validate_all, build_excel,
    StudentRecord, CourseRecord, DEFAULT_SCHOOL_ID, log as engine_log
)
from matchers import (
    build_candidates, build_name_queries, classify_match,
    name_similarity, matricule_similarity, parse_api_payload,
    score_candidate,
)
from form_b import (
    load_form_b, find_catalogs, majority_specialty,
    infer_department, detect_department, cross_check, build_form_b_excel,
    parse_catalog_filename, derive_form_b_from_students,
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

BASE      = Path(__file__).parent
UPLOAD_DIR = BASE / "uploads"
OUTPUT_DIR = BASE / "outputs"
FORM_B_DIR = BASE / "form_b"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
FORM_B_DIR.mkdir(exist_ok=True)

jobs = {}
jobs_lock = threading.Lock()


# ── helpers ──────────────────────────────────────────────────────────────────
def allowed(fn): return fn.lower().endswith('.pdf')

def student_to_dict(s: StudentRecord):
    errs = s.validation_errors()
    courses = []
    for c in s.courses:
        ce = c.validation_errors()
        courses.append({
            "code": c.code, "title": c.title, "status": c.status,
            "credit_value": c.credit_value, "credit_earned": c.credit_earned,
            "ca": c.ca, "exam": c.exam, "total": c.total,
            "grade_point": c.grade_point, "grade": c.grade,
            "semester": c.semester, "errors": ce, "valid": len(ce) == 0
        })
    sems = sorted(set(c["semester"] for c in courses))
    sem_gpas = {}
    for sem in sems:
        sc = [c for c in courses if c["semester"] == sem]
        tw = sum(c["credit_value"] * c["grade_point"] for c in sc)
        tv = sum(c["credit_value"] for c in sc)
        sem_gpas[str(sem)] = round(tw / tv, 2) if tv else 0.0
    return {
        "matricule": s.matricule, "name": s.name,
        "faculty": s.faculty, "specialty": s.specialty,
        "department": s.department, "level": s.level,
        "academic_year": s.academic_year, "source_page": s.source_page,
        "courses": courses, "course_count": len(courses),
        "semesters": sems,
        "semester_counts": {str(sem): sum(1 for c in courses if c["semester"] == sem) for sem in sems},
        "errors": errs, "warnings": s.parse_warnings,
        "valid": len(errs) == 0, "gpa_sem": sem_gpas,
    }


# ── background jobs ───────────────────────────────────────────────────────────
def run_parse_job(job_id, pdf_paths, school_id):
    def push(msg, level="info"):
        with jobs_lock:
            jobs[job_id]["log"].append({"t": datetime.now().strftime("%H:%M:%S"), "msg": msg, "level": level})
    def set_prog(p):
        with jobs_lock:
            jobs[job_id]["progress"] = p

    try:
        push("Engine starting…")
        all_students = []
        for fi, path in enumerate(pdf_paths):
            push(f"Parsing {fi+1}/{len(pdf_paths)}: {Path(path).name}")
            set_prog(int(fi / len(pdf_paths) * 40))

            # redirect engine log → job log
            orig = engine_log.ok, engine_log.warn, engine_log.err, engine_log.info
            engine_log.ok   = lambda m: push(f"✓ {m}", "ok")
            engine_log.warn = lambda m: push(f"⚠ {m}", "warn")
            engine_log.err  = lambda m: push(f"✗ {m}", "error")
            engine_log.info = lambda m: push(m, "info")
            try:
                students = parse_pdf(path, school_id)
            finally:
                engine_log.ok, engine_log.warn, engine_log.err, engine_log.info = orig

            push(f"Extracted {len(students)} student(s) from {Path(path).name}", "ok")
            all_students.extend(students)

        set_prog(55)
        push("Validating records…")
        student_dicts, err_map = [], {}
        for s in all_students:
            sd = student_to_dict(s)
            student_dicts.append(sd)
            if not sd["valid"]:
                err_map[s.matricule] = sd["errors"]

        lv = "ok" if not err_map else "warn"
        push(f"Validation: {len(all_students)-len(err_map)} passed, {len(err_map)} with issues", lv)
        set_prog(75)

        with jobs_lock:
            jobs[job_id].update({
                "students": student_dicts, "errors": err_map, "_raw": all_students,
                "status": "awaiting_confirmation", "progress": 80,
                "meta": {
                    "total_students": len(all_students),
                    "total_courses": sum(len(s.courses) for s in all_students),
                    "files": [Path(p).name for p in pdf_paths],
                    "school_id": school_id,
                    "parsed_at": datetime.now().isoformat(),
                    "specialty": majority_specialty(all_students),
                    "level": all_students[0].level if all_students else "",
                    "year": all_students[0].academic_year if all_students else "",
                }
            })
        push("Ready for review — confirm to generate Excel", "ok")

    except Exception as e:
        push(f"Fatal: {e}", "error")
        push(traceback.format_exc(), "error")
        with jobs_lock:
            jobs[job_id]["status"] = "failed"


def run_excel_job(job_id, school_id, overrides):
    def push(msg, level="info"):
        with jobs_lock:
            jobs[job_id]["log"].append({"t": datetime.now().strftime("%H:%M:%S"), "msg": msg, "level": level})
    try:
        with jobs_lock:
            raw = jobs[job_id]["_raw"]

        # Apply overrides
        for mat, ovs in overrides.items():
            for s in raw:
                if s.matricule == mat:
                    for ov in ovs:
                        for c in s.courses:
                            if c.code == ov["code"] and c.semester == ov["semester"]:
                                if "ca"   in ov: c.ca   = float(ov["ca"])
                                if "exam" in ov: c.exam = float(ov["exam"])
                                c.total = round(c.ca + c.exam, 2)
                                push(f"Override: {mat}/{c.code} Sem{c.semester} CA={c.ca} Exam={c.exam}", "warn")

        push("Building Excel…")
        fname    = _results_fname(job_id, raw)
        out_path = str(OUTPUT_DIR / fname)
        with jobs_lock:
            form_b_rows = jobs[job_id]["form_b"].get("rows") or []
        build_excel(raw, out_path, school_id, form_b_rows=form_b_rows)

        with jobs_lock:
            jobs[job_id].update({"output_file": fname, "status": "done", "progress": 100})
        push(f"Excel ready: {fname}", "ok")

    except Exception as e:
        push(f"Excel build failed: {e}", "error")
        with jobs_lock:
            jobs[job_id]["status"] = "failed"


# ── routes ────────────────────────────────────────────────────────────────────
@app.route("/")
def index(): return render_template("index.html")

@app.route("/api/upload", methods=["POST"])
def api_upload():
    files     = request.files.getlist("files")
    school_id = int(request.form.get("school_id", DEFAULT_SCHOOL_ID))
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No files provided"}), 400

    saved = []
    for f in files:
        if not allowed(f.filename):
            return jsonify({"error": f"Not a PDF: {f.filename}"}), 400
        fn   = f"{uuid.uuid4().hex}_{f.filename}"
        path = str(UPLOAD_DIR / fn)
        f.save(path); saved.append(path)

    job_id = uuid.uuid4().hex[:12]
    with jobs_lock:
        jobs[job_id] = {"status":"parsing","progress":0,"log":[],"students":[],
                        "errors":{},"_raw":[],"output_file":None,"meta":{},"dismissed_errors":False,
                        "mat_check":{"status":"idle","progress":0,"results":[],"checked":0,"total":0},
                        "form_b":{"status":"idle","catalog":None,"department":"","level":"",
                                  "rows":[],"stats":None,"output_file":None,"error":None}}

    threading.Thread(target=run_parse_job, args=(job_id,saved,school_id), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/api/job/<job_id>")
def api_job(job_id):
    with jobs_lock:
        if job_id not in jobs: return jsonify({"error":"Unknown job"}),404
        j = jobs[job_id]
        return jsonify({
            "status": j["status"], "progress": j["progress"],
            "log": j["log"][-60:], "meta": j["meta"],
            "error_count": len(j["errors"]), "student_count": len(j["students"]),
            "dismissed_errors": j.get("dismissed_errors", False),
        })


@app.route("/api/job/<job_id>/students")
def api_students(job_id):
    with jobs_lock:
        if job_id not in jobs: return jsonify({"error":"Unknown job"}),404
        return jsonify({"students":jobs[job_id]["students"],
                        "errors":jobs[job_id]["errors"],
                        "meta":jobs[job_id]["meta"]})


@app.route("/api/job/<job_id>/confirm", methods=["POST"])
def api_confirm(job_id):
    data      = request.get_json() or {}
    overrides = data.get("overrides", {})
    school_id = int(data.get("school_id", DEFAULT_SCHOOL_ID))
    with jobs_lock:
        if job_id not in jobs: return jsonify({"error":"Unknown job"}),404
        if jobs[job_id]["status"] not in ("awaiting_confirmation","done"):
            return jsonify({"error":"Not ready"}),400
        jobs[job_id]["status"] = "building"; jobs[job_id]["progress"] = 85
    threading.Thread(target=run_excel_job, args=(job_id,school_id,overrides), daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/job/<job_id>/dismiss-errors", methods=["POST"])
def api_dismiss_errors(job_id):
    with jobs_lock:
        if job_id not in jobs: return jsonify({"error":"Unknown job"}),404
        jobs[job_id]["dismissed_errors"] = True
        jobs[job_id]["errors"] = {}
    return jsonify({"ok": True})


@app.route("/api/job/<job_id>/download")
def api_download(job_id):
    with jobs_lock:
        if job_id not in jobs: return jsonify({"error":"Unknown job"}),404
        fname = jobs[job_id].get("output_file")
    if not fname: return jsonify({"error":"File not ready"}),400
    path = OUTPUT_DIR / fname
    if not path.exists(): return jsonify({"error":"File missing"}),404
    return send_file(str(path), as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/job/<job_id>/student/<matricule>", methods=["PATCH"])
def api_patch(job_id, matricule):
    data = request.get_json() or {}
    with jobs_lock:
        if job_id not in jobs: return jsonify({"error":"Unknown job"}),404
        raw = jobs[job_id]["_raw"]
    for s in raw:
        if s.matricule == matricule:
            for key, vals in data.items():
                # key format: "ACC202_sem1"
                parts = key.split("_sem")
                if len(parts) != 2: continue
                code, sem_str = parts[0], parts[1]
                sem = int(sem_str)
                for c in s.courses:
                    if c.code == code and c.semester == sem:
                        if "ca"   in vals: c.ca   = float(vals["ca"])
                        if "exam" in vals: c.exam = float(vals["exam"])
                        c.total = round(c.ca + c.exam, 2)
            sd = student_to_dict(s)
            with jobs_lock:
                for i,ex in enumerate(jobs[job_id]["students"]):
                    if ex["matricule"] == matricule:
                        jobs[job_id]["students"][i] = sd; break
                if sd["valid"]: jobs[job_id]["errors"].pop(matricule, None)
                else:           jobs[job_id]["errors"][matricule] = sd["errors"]
            return jsonify({"student": sd})
    return jsonify({"error":"Student not found"}),404


# ── Matricule Check ────────────────────────────────────────────────────────
API_BASE   = "https://study.landmark.cm/matriculeCheck.php"
API_YEARS  = ["2025/2026", "2024/2025", "2023/2024"]
API_SKEY   = "123456Q"
MC_WORKERS      = 8      # concurrent portal requests
MC_TIMEOUT      = 8      # seconds per request
MC_RETRIES      = 2      # retries on transient failure
MC_QUERY_PHASE1 = 6      # queries per student in the first pass
MC_QUERY_PHASE2 = 16     # expanded budget for unresolved students
MC_RATE_INTERVAL = 0.12  # min seconds between portal requests


def academic_years(back=15):
    """Descending list of academic years from the current one, e.g.
    ['2025/2026', '2024/2025', ...]. A new academic year begins in September."""
    now = datetime.now()
    start = now.year if now.month >= 9 else now.year - 1
    return [f"{y}/{y + 1}" for y in range(start, start - back - 1, -1)]


class _RateLimiter:
    """Paces outbound requests so the portal never gets hammered."""

    def __init__(self, min_interval=MC_RATE_INTERVAL):
        self._lock = threading.Lock()
        self._min = min_interval
        self._next = 0.0

    def wait(self):
        with self._lock:
            now = time.monotonic()
            delay = self._next - now
            self._next = max(now, self._next) + self._min
        if delay > 0:
            time.sleep(delay)


def _ordered_years(student, years):
    """Student's own academic year first, then the fallback year list."""
    own = student.get("academic_year", "")
    ordered = []
    for y in ([own] if own else []) + list(years):
        if y not in ordered:
            ordered.append(y)
    return ordered


def _query_portal(q, year, session, cache, limiter):
    key = (q.upper(), year)
    if key in cache:
        return cache[key]
    for attempt in range(MC_RETRIES + 1):
        try:
            limiter.wait()
            resp = session.get(API_BASE, params={
                "year": year, "skey": API_SKEY, "name": q,
            }, timeout=MC_TIMEOUT)
            payload = None
            if resp.status_code == 200:
                payload = parse_api_payload(resp.json())
            cache[key] = (resp.status_code, payload)
            return cache[key]
        except Exception:
            if attempt < MC_RETRIES:
                time.sleep(0.5 * (attempt + 1))
                continue
            cache[key] = (0, None)
            return cache[key]


def _run_matricule_check(student_dicts, years=None, progress_cb=None):
    """Check every student against the portal with bounded, cached,
    concurrent queries across the given academic years. Returns the
    per-student result list."""
    years = years or API_YEARS
    total = len(student_dicts)
    checked = 0
    cache = {}
    limiter = _RateLimiter()
    # idx -> {api_matricule: {"name","years":set,"query_used"}}
    raw = [{} for _ in student_dicts]
    unresolved = set(range(total))

    def progress(delta):
        nonlocal checked
        checked += delta
        if progress_cb:
            progress_cb(min(checked, total), total)

    with requests.Session() as session:
        for phase, cutoff in ((1, MC_QUERY_PHASE1), (2, MC_QUERY_PHASE2)):
            tasks = []
            for idx in sorted(unresolved):
                s = student_dicts[idx]
                queries = build_name_queries(s.get("name", ""), limit=cutoff)
                qs = queries[:MC_QUERY_PHASE1] if phase == 1 else queries[MC_QUERY_PHASE1:]
                for q in qs:
                    for year in _ordered_years(s, years):
                        tasks.append((idx, q, year))
            if not tasks:
                break

            with ThreadPoolExecutor(max_workers=MC_WORKERS) as ex:
                futures = {
                    ex.submit(_query_portal, q, year, session, cache, limiter): (idx, q, year)
                    for (idx, q, year) in tasks
                }
                for fut in as_completed(futures):
                    idx, q, year = futures[fut]
                    _, payload = fut.result()
                    if payload:
                        rec = raw[idx]
                        mat = payload["matricule"]
                        entry = rec.get(mat)
                        if entry is None:
                            rec[mat] = {
                                "name": payload["name"],
                                "years": {year},
                                "query_used": q,
                            }
                        else:
                            if payload["name"] and not entry["name"]:
                                entry["name"] = payload["name"]
                            entry["years"].add(year)
                    progress(1)

            if phase == 1:
                still = set()
                for idx in unresolved:
                    s = student_dicts[idx]
                    cands = build_candidates(s.get("name", ""), s.get("matricule", ""), raw[idx])
                    if cands and classify_match(s.get("name", ""), s.get("matricule", ""), cands) == "verified":
                        continue
                    still.add(idx)
                unresolved = still

    results = []
    for idx, s in enumerate(student_dicts):
        pdf_name = s.get("name", "")
        pdf_mat = s.get("matricule", "")
        res = {
            "matricule": pdf_mat, "name": pdf_name,
            "status": "skipped", "api_matched": False,
            "api_matricule": None, "api_name": None,
            "mismatch": False, "similarity": 0.0, "confidence": 0.0,
            "name_similarity": 0.0, "matricule_matched": False,
            "candidates": [], "years": [], "error": None, "reason": "",
        }
        if not pdf_name:
            res["error"] = "No name to search"
            results.append(res)
            continue

        cands = build_candidates(pdf_name, pdf_mat, raw[idx])[:8]
        status = classify_match(pdf_name, pdf_mat, cands)
        res["status"] = status
        res["candidates"] = cands
        if cands:
            best = cands[0]
            res.update({
                "api_matched": True,
                "api_matricule": best["matricule"],
                "api_name": best["name"],
                "similarity": best["confidence"],
                "confidence": best["confidence"],
                "name_similarity": best["name_similarity"],
                "matricule_matched": best["matricule_similarity"] >= 0.9,
                "years": best["years"],
                "mismatch": status == "mismatch",
            })
            if status == "mismatch":
                res["reason"] = "Portal matricule differs from PDF matricule"
            elif status == "review":
                res["reason"] = "Low confidence — review candidates manually"
        else:
            res["error"] = "Not found in API"
        results.append(res)
    return results


def run_matricule_check_job(job_id):
    def progress(done, total):
        with jobs_lock:
            mc = jobs[job_id].get("mat_check")
            if not mc:
                return
            mc["checked"] = done
            mc["progress"] = round(done / total * 100) if total else 0

    try:
        with jobs_lock:
            student_dicts = jobs[job_id].get("students", [])
            years = jobs[job_id].get("mat_check", {}).get("years") or API_YEARS
        results = _run_matricule_check(student_dicts, years, progress_cb=progress)
        with jobs_lock:
            jobs[job_id]["mat_check"].update({
                "status": "done", "progress": 100, "results": results,
                "checked": len(student_dicts), "total": len(student_dicts),
                "years": years,
            })
    except Exception as e:
        traceback.print_exc()
        with jobs_lock:
            jobs[job_id]["mat_check"].update({"status": "failed", "error": str(e)})


@app.route("/api/academic-years")
def api_academic_years():
    return jsonify({"years": academic_years()})


@app.route("/api/job/<job_id>/matricule-check", methods=["GET", "POST"])
def api_matricule_check(job_id):
    with jobs_lock:
        if job_id not in jobs:
            return jsonify({"error": "Unknown job"}), 404

    if request.method == "POST":
        data = request.get_json(silent=True) or {}
        valid = set(academic_years())
        years = [str(y).strip() for y in (data.get("years") or [])]
        years = [y for y in years if y in valid] or API_YEARS
        with jobs_lock:
            mc = jobs[job_id].get("mat_check")
            if mc and mc.get("status") == "running":
                return jsonify({"ok": True, "running": True})
            jobs[job_id]["mat_check"] = {
                "status": "running", "progress": 0, "results": [],
                "checked": 0, "total": 0, "years": years,
            }
        threading.Thread(target=run_matricule_check_job, args=(job_id,), daemon=True).start()
        return jsonify({"ok": True, "running": True})

    with jobs_lock:
        mc = jobs[job_id].get("mat_check") or {
            "status": "idle", "progress": 0, "results": [],
            "checked": 0, "total": 0, "years": API_YEARS,
        }
    return jsonify(mc)


def _apply_override_locked(job_id, old_mat, new_mat, old_name=None):
    """Apply one matricule override and rebuild the affected student dict.
    When ``old_mat`` is empty the student is located by ``old_name`` instead
    (for transcripts whose matricule was never detected).
    Caller must hold ``jobs_lock``. Returns ``(ok, payload)``."""
    raw = jobs[job_id]["_raw"]
    stu_list = jobs[job_id]["students"]

    target = None
    if old_mat:
        target = next((s for s in raw if s.matricule == old_mat), None)
    if target is None and old_name:
        target = next((s for s in raw if s.name == old_name), None)
    if target is None:
        return False, {"error": f"Student {old_mat or old_name} not found in raw data"}

    target.matricule = new_mat
    old_key = old_mat if old_mat else (old_name or "")

    for s in raw:
        if s.matricule == new_mat:
            sd = student_to_dict(s)
            for i, ex in enumerate(stu_list):
                if (old_mat and ex["matricule"] == old_mat) or \
                   (not old_mat and ex["name"] == old_name):
                    stu_list[i] = sd
                    errors = jobs[job_id]["errors"]
                    if old_key in errors:
                        errors[new_mat] = errors.pop(old_key)
                    if sd["valid"]:
                        errors.pop(new_mat, None)
                    else:
                        errors[new_mat] = sd["errors"]
                    return True, {"student": sd, "old_matricule": old_key,
                                  "new_matricule": new_mat}
            break
    return False, {"error": "Could not apply override"}


@app.route("/api/job/<job_id>/matricule-override", methods=["POST"])
def api_matricule_override(job_id):
    data = request.get_json() or {}
    old_mat = (data.get("old_matricule") or "").strip()
    new_mat = (data.get("new_matricule") or "").strip()
    if not old_mat or not new_mat:
        return jsonify({"error": "old_matricule and new_matricule required"}), 400

    with jobs_lock:
        if job_id not in jobs:
            return jsonify({"error": "Unknown job"}), 404
        ok, payload = _apply_override_locked(job_id, old_mat, new_mat)
    status = 404 if "not found" in payload.get("error", "") else 400
    return jsonify(payload), (200 if ok else status)


@app.route("/api/job/<job_id>/matricule-bulk", methods=["POST"])
def api_matricule_bulk(job_id):
    data = request.get_json() or {}
    mappings = data.get("mappings") or []

    with jobs_lock:
        if job_id not in jobs:
            return jsonify({"error": "Unknown job"}), 404
        applied, failures = [], []
        for m in mappings:
            old_m = (m.get("old_matricule") or m.get("from") or "").strip()
            old_n = (m.get("old_name") or "").strip()
            new_m = (m.get("new_matricule") or m.get("to") or "").strip()
            if (not old_m and not old_n) or not new_m:
                failures.append({"old": old_m or old_n, "new": new_m,
                                 "error": "missing old/new matricule"})
                continue
            ok, payload = _apply_override_locked(job_id, old_m, new_m, old_name=old_n)
            if ok:
                applied.append(payload)
            else:
                failures.append({"old": old_m or old_n, "new": new_m,
                                 "error": payload.get("error")})
    return jsonify({"ok": True, "applied": len(applied), "failures": failures})


# ── FORM B ──────────────────────────────────────────────────────────────────
def _safe_tag(text):
    tag = re.sub(r"[^A-Z0-9]+", "-", str(text or "").upper()).strip("-")
    return tag or ""


def _results_fname(job_id, raw):
    """Descriptive output name: Results_<DEPARTMENT>_<LEVEL>_<YEAR>_<ts>.xlsx"""
    with jobs_lock:
        meta = jobs[job_id].get("meta", {})
    dept  = infer_department(raw)
    level = meta.get("level") or (raw[0].level if raw else "")
    year  = meta.get("year") or (raw[0].academic_year if raw else "")
    parts = [p for p in ["Results", _safe_tag(dept), _safe_tag(level),
                         _safe_tag(year)] if p]
    core  = "_".join(parts)
    return f"{core}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"


def _job_form_b_context(job_id):
    """Detect department/level for a job and list available catalogues."""
    with jobs_lock:
        if job_id not in jobs:
            return None
        students = jobs[job_id].get("students", [])
        level = jobs[job_id].get("meta", {}).get("level", "")
        state = jobs[job_id]["form_b"]

    matricules = [s.get("matricule", "") for s in students]
    department = state.get("department") or infer_department(students)
    if not level:
        level = state.get("level", "")

    catalogs = find_catalogs(str(FORM_B_DIR), department, int(level) if level else 0)
    for c in catalogs:
        c.pop("path", None)
    return {"department": department, "level": level, "catalogs": catalogs}


@app.route("/api/job/<job_id>/form-b")
def api_form_b(job_id):
    ctx = _job_form_b_context(job_id)
    if ctx is None:
        return jsonify({"error": "Unknown job"}), 404
    with jobs_lock:
        state = {k: jobs[job_id]["form_b"].get(k) for k in
                 ("status", "catalog", "rows", "output_file", "error")}
    return jsonify({**ctx, "state": state})


@app.route("/api/job/<job_id>/form-b/generate", methods=["POST"])
def api_form_b_generate(job_id):
    data = request.get_json() or {}
    catalog_name = (data.get("catalog") or "").strip()
    ctx = _job_form_b_context(job_id)
    if ctx is None:
        return jsonify({"error": "Unknown job"}), 404

    with jobs_lock:
        if job_id not in jobs:
            return jsonify({"error": "Unknown job"}), 404
        raw = jobs[job_id]["_raw"]
        meta = jobs[job_id].get("meta", {})

    school_id = meta.get("school_id") or DEFAULT_SCHOOL_ID
    year = meta.get("year") or ""
    department = ctx["department"]
    level = ctx["level"]

    # FORM B is derived from the uploaded transcripts (deduplicated across
    # students). The official catalogue in form_b/ is only an optional
    # reference used to produce a comparison report.
    try:
        rows, stats = derive_form_b_from_students(
            raw, department=department, level=level, year=year,
            school_id=school_id)
    except Exception as e:
        return jsonify({"error": f"Cross-check failed: {e}"}), 500

    catalog_compare = None
    if catalog_name:
        catalog_path = str(FORM_B_DIR / catalog_name)
        catalog_compare = {"name": catalog_name}
        if not os.path.isfile(catalog_path):
            catalog_compare["error"] = "catalogue file not found in form_b/"
        else:
            try:
                catalog = load_form_b(catalog_path)
                cc = cross_check(raw, catalog)
                catalog_compare["coverage"] = {cs["code"]: cs["count"]
                                               for cs in cc["course_stats"]}
                catalog_compare["per_student"] = [{
                    "matricule": ps["matricule"], "name": ps["name"],
                    "taken": ps["taken"], "matched": ps["matched"],
                    "missing": [r["code"] for r in ps["missing"]],
                    "unexpected": [u["code"] for u in ps["unexpected"]],
                    "issues": ps["issues"],
                } for ps in cc["students"]]
                catalog_compare["catalog_count"] = cc["catalog_count"]
            except Exception as e:
                catalog_compare["error"] = str(e)

    dept_tag = _safe_tag(department) or "DEPARTMENT"
    lvl_tag  = _safe_tag(level) or "X"
    fname = f"FORM_B_{dept_tag}_{lvl_tag}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    out_path = str(OUTPUT_DIR / fname)
    try:
        build_form_b_excel(rows, stats, out_path)
    except Exception as e:
        return jsonify({"error": f"FORM B build failed: {e}"}), 500

    with jobs_lock:
        jobs[job_id]["form_b"].update({
            "status": "done", "catalog": catalog_name,
            "department": department, "level": level,
            "rows": rows, "stats": stats, "output_file": fname, "error": None,
        })
    return jsonify({
        "ok": True, "output_file": fname,
        "department": department, "level": level,
        "catalog": catalog_name, "source": "uploaded",
        "stats": {
            "catalog_count": stats["catalog_count"],
            "student_count": stats["student_count"],
            "courses": [{
                "code": cs["code"], "description": cs["description"],
                "semester": cs["semester"], "credit": cs["credit"],
                "count": cs["count"],
            } for cs in stats["course_stats"]],
            "coverage": {cs["code"]: cs["count"] for cs in stats["course_stats"]},
            "per_student": [{
                "matricule": ps["matricule"], "name": ps["name"],
                "taken": ps["taken"], "matched": ps["matched"],
                "missing": ps["missing"], "unexpected": ps["unexpected"],
                "issues": ps["issues"],
            } for ps in stats["students"]],
        },
        "catalog_compare": catalog_compare,
    })


@app.route("/api/form-b/<filename>/download")
def api_form_b_dl(filename):
    path = OUTPUT_DIR / filename
    if not path.exists(): return jsonify({"error":"Not found"}),404
    return send_file(str(path), as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/history")
def api_history():
    files = sorted(OUTPUT_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return jsonify({"files": [{"name":p.name,
                                "size_kb": round(p.stat().st_size/1024,1),
                                "created": datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M")}
                               for p in files[:30]]})


@app.route("/api/history/<filename>/download")
def api_hist_dl(filename):
    path = OUTPUT_DIR / filename
    if not path.exists(): return jsonify({"error":"Not found"}),404
    return send_file(str(path), as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    print("\n" + "=" * 50)
    print("  TRANSCRIPT CONVERSION SYSTEM  v2.0")
    print("  -> http://localhost:5000")
    print("=" * 50 + "\n")
    app.run(debug=False, host="0.0.0.0", port=5000, threaded=True)
