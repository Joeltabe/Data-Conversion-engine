#!/usr/bin/env python3
"""
Transcript Conversion System — Flask Backend
Run:  python app.py
Open: http://localhost:5000
"""
import os, sys, json, uuid, traceback, threading, requests
from pathlib import Path
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template, send_from_directory

sys.path.insert(0, os.path.dirname(__file__))
from pdf_to_excel_engine import (
    parse_pdf, validate_all, build_excel,
    StudentRecord, CourseRecord, DEFAULT_SCHOOL_ID, log as engine_log
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

BASE      = Path(__file__).parent
UPLOAD_DIR = BASE / "uploads"
OUTPUT_DIR = BASE / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

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
                    "specialty": all_students[0].specialty if all_students else "",
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
        fname    = f"Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        out_path = str(OUTPUT_DIR / fname)
        build_excel(raw, out_path, school_id)

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
                        "errors":{},"_raw":[],"output_file":None,"meta":{},"dismissed_errors":False}

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
API_BASE  = "https://study.landmark.cm/matriculeCheck.php"
API_YEAR  = "2024/2025"
API_SKEY  = "123456Q"

import difflib

def _name_variations(name):
    parts = name.strip().upper().split()
    if not parts:
        return []
    seen = set()
    out = []
    def add(v):
        v = v.strip()
        if v and v not in seen:
            seen.add(v); out.append(v)
    add(name.strip())
    if len(parts) == 1:
        return out
    # Every single word
    for p in parts:
        add(p)
    # Every rotation
    for i in range(len(parts)):
        r = parts[i:] + parts[:i]
        add(" ".join(r))
        if len(r) >= 2:
            add(f"{r[0]} {r[-1][0]}")    # "FIRST L."
            add(f"{r[-1]} {r[0][0]}")    # "LAST F."
    # Common orderings: first+last, last+first
    add(f"{parts[0]} {parts[-1]}")
    add(f"{parts[-1]} {parts[0]}")
    # With initials
    if len(parts[-1]) >= 1:
        add(f"{parts[0]} {parts[-1][0]}")
    if len(parts[0]) >= 1:
        add(f"{parts[-1]} {parts[0][0]}")
    # All word pairs
    for i in range(len(parts)):
        for j in range(len(parts)):
            if i != j:
                add(f"{parts[i]} {parts[j]}")
    return out

def _name_similarity(a, b):
    return difflib.SequenceMatcher(None, a.strip().upper(), b.strip().upper()).ratio()

@app.route("/api/job/<job_id>/matricule-check", methods=["POST"])
def api_matricule_check(job_id):
    with jobs_lock:
        if job_id not in jobs: return jsonify({"error":"Unknown job"}),404
        students_data = jobs[job_id].get("students", [])

    # Query cache across all students to avoid redundant API calls
    _qcache = {}

    def _query_portal(q):
        if q in _qcache:
            return _qcache[q]
        try:
            resp = requests.get(API_BASE, params={
                "year": API_YEAR, "skey": API_SKEY, "name": q,
            }, timeout=5)
            _qcache[q] = (resp.status_code, resp.json() if resp.status_code == 200 else None)
        except Exception:
            _qcache[q] = (0, None)
        return _qcache[q]

    results = []
    for s in students_data:
        mat  = s.get("matricule","")
        name = s.get("name","")
        res = {
            "matricule": mat, "name": name,
            "api_matched": False, "api_matricule": None,
            "api_name": None, "mismatch": False,
            "similarity": 0.0, "candidates": [],
            "error": None,
        }
        if not name:
            res["error"] = "No name to search"; results.append(res); continue

        queries = _name_variations(name)
        seen_candidates = {}
        for q in queries:
            status_code, api = _query_portal(q)
            if status_code == 200 and api and api.get("matricule"):
                api_mat = api["matricule"].strip()
                api_nam = api.get("fname","").strip()
                if api_mat not in seen_candidates:
                    sim = _name_similarity(name, api_nam)
                    seen_candidates[api_mat] = {
                        "matricule": api_mat,
                        "name": api_nam,
                        "similarity": round(sim, 4),
                        "query_used": q,
                    }

        candidates = sorted(seen_candidates.values(), key=lambda c: -c["similarity"])
        if candidates:
            best = candidates[0]
            res["api_matched"] = True
            res["api_matricule"] = best["matricule"]
            res["api_name"] = best["name"]
            res["similarity"] = best["similarity"]
            res["candidates"] = candidates
            if best["matricule"].upper() != mat.strip().upper():
                res["mismatch"] = True
        else:
            res["candidates"] = []
            if not res["error"]:
                res["error"] = "Not found in API"

        results.append(res)

    return jsonify({"results": results})


@app.route("/api/job/<job_id>/matricule-override", methods=["POST"])
def api_matricule_override(job_id):
    data = request.get_json() or {}
    old_mat = data.get("old_matricule","")
    new_mat = data.get("new_matricule","")
    if not old_mat or not new_mat:
        return jsonify({"error":"old_matricule and new_matricule required"}),400

    with jobs_lock:
        if job_id not in jobs: return jsonify({"error":"Unknown job"}),404
        raw = jobs[job_id]["_raw"]
        stu_list = jobs[job_id]["students"]

    # Update in raw StudentRecord objects
    for s in raw:
        if s.matricule == old_mat:
            s.matricule = new_mat
            break
    else:
        return jsonify({"error":"Student not found in raw data"}),404

    # Rebuild dict from updated raw record
    for s in raw:
        if s.matricule == new_mat:
            sd = student_to_dict(s)
            for i, ex in enumerate(stu_list):
                if ex["matricule"] == old_mat:
                    stu_list[i] = sd
                    # Move errors key if present
                    with jobs_lock:
                        if old_mat in jobs[job_id]["errors"]:
                            jobs[job_id]["errors"][new_mat] = jobs[job_id]["errors"].pop(old_mat)
                        if sd["valid"]:
                            jobs[job_id]["errors"].pop(new_mat, None)
                        else:
                            jobs[job_id]["errors"][new_mat] = sd["errors"]
                    return jsonify({"student": sd, "old_matricule": old_mat, "new_matricule": new_mat})
            break

    return jsonify({"error":"Could not apply override"}),500


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
