# -*- coding: utf-8 -*-
"""HTTP-level integration test through Flask's test client.
Run: python test_api_flow.py
"""
import http.server
import io
import os
import sys
import threading
import time
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import app  # noqa: E402

PORTAL_DB = [
    ("LMU-24NUR001", "ASEH IYA MAKANE", ["MAKANE", "ASEH"]),
    ("LMU24NUR014", "NGWA LEM JANET", ["NGWA", "JANET"]),
    ("LMU24NUR016", "EYAMBE MBOTAKE DAUGHTER", ["EYAMBE", "DAUGHTER"]),
]


class Handler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        from urllib.parse import parse_qs, urlparse
        qs = parse_qs(urlparse(self.path).query)
        name = (qs.get("name") or [""])[0].upper()
        for mat, fname, keys in PORTAL_DB:
            if all(k in name for k in keys):
                body = f'{{"matricule": "{mat}", "fname": "{fname}"}}'.encode()
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
        body = b'{"ans3": "Did not find a Name like that !"}'
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *a):
        pass


def wait_for(client, url, key, value, tries=120, delay=0.5):
    for _ in range(tries):
        d = client.get(url).json
        if d.get(key) == value:
            return d
        time.sleep(delay)
    raise RuntimeError(f"timeout waiting {url} {key}={value}, last={d}")


def main():
    pdf = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "uploads", "07c4a0ac1e374ae185ff84245f049076_Accounting L200.pdf")
    server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    port = server.server_address[1]
    threading.Thread(target=server.serve_forever, daemon=True).start()
    app.API_BASE = f"http://127.0.0.1:{port}/matriculeCheck.php"
    app._RateLimiter.__init__.__defaults__ = (0.0,)

    client = app.app.test_client()

    # 1. upload
    with open(pdf, "rb") as f:
        r = client.post("/api/upload",
                        data={"files": (io.BytesIO(f.read()), "nursing.pdf"), "school_id": "10"},
                        content_type="multipart/form-data")
    job_id = r.json["job_id"]
    print("job:", job_id)

    # 2. parse until ready
    wait_for(client, f"/api/job/{job_id}", "status", "awaiting_confirmation")
    stu = client.get(f"/api/job/{job_id}/students").json["students"]
    print("students parsed:", len(stu))

    # 3. start matricule check (background) and poll
    r = client.post(f"/api/job/{job_id}/matricule-check", json={"years": ["2023/2024"]})
    assert r.json.get("running"), r.json
    d = wait_for(client, f"/api/job/{job_id}/matricule-check", "status", "done", tries=200, delay=0.4)
    assert d.get("years") == ["2023/2024"], d
    results = d.get("results", [])
    print("mat-check:", Counter(x["status"] for x in results), "| checked", d["checked"], "/", d["total"])

    # re-run must not start a second job while running
    assert client.post(f"/api/job/{job_id}/matricule-check", json={"years": ["2023/2024"]}).json["running"]

    # invalid years fall back to the default list
    assert client.get("/api/academic-years").json["years"][0] == "2025/2026"

    # 4. override a mismatch candidate if any
    target = next((x for x in results if x["status"] == "mismatch"), None)
    if target:
        r = client.post(f"/api/job/{job_id}/matricule-override",
                        json={"old_matricule": target["matricule"], "new_matricule": target["api_matricule"]})
        assert r.json.get("student"), r.json
        print("override ok:", target["matricule"], "->", target["api_matricule"])

    # 4a. bulk migrate matricules (applies to raw + student dicts)
    first = stu[0]["matricule"]
    r = client.post(f"/api/job/{job_id}/matricule-bulk",
                    json={"mappings": [{"old_matricule": first, "new_matricule": "LMU-ACCXXX"}]})
    assert r.json.get("applied") == 1, r.json
    assert client.get(f"/api/job/{job_id}/students").json["students"][0]["matricule"] == "LMU-ACCXXX"
    r = client.post(f"/api/job/{job_id}/matricule-bulk",
                    json={"mappings": [{"old_matricule": "LMU-NOPE", "new_matricule": "LMU-ACCYYY"}]})
    assert r.json["applied"] == 0 and len(r.json["failures"]) == 1, r.json
    print("bulk migrate ok:", first, "-> LMU-ACCXXX (+ failure path)")

    # 4a2. migrate a student by name (matricule was never detected)
    r = client.post(f"/api/job/{job_id}/matricule-bulk",
                    json={"mappings": [{"old_matricule": "", "old_name": stu[0]["name"],
                                        "new_matricule": "LMU-ACCZZZ"}]})
    assert r.json.get("applied") == 1, r.json
    assert client.get(f"/api/job/{job_id}/students").json["students"][0]["matricule"] == "LMU-ACCZZZ"
    print("name-based migrate ok (empty matricule)")

    # 4b. FORM B context + generation from uploaded transcripts
    fb = client.get(f"/api/job/{job_id}/form-b").json
    assert fb.get("catalogs"), fb
    assert fb.get("department") == "ACCOUNTING", fb
    assert fb.get("level") == "200", fb
    print("form-b dept/level:", fb.get("department"), fb.get("level"), "| catalogs:", [c["name"] for c in fb["catalogs"]])
    r = client.post(f"/api/job/{job_id}/form-b/generate", json={"catalog": ""})
    assert r.json.get("output_file"), r.json
    assert r.json["source"] == "uploaded", r.json
    fb_stats = r.json["stats"]
    assert fb_stats["catalog_count"] > 0, fb_stats
    assert fb_stats["student_count"] == 11, fb_stats
    print("form-b (from uploads):", fb_stats["catalog_count"], "courses,", fb_stats["student_count"], "students")
    dl = client.get(f"/api/form-b/{r.json['output_file']}/download")
    assert dl.status_code == 200 and dl.data[:2] == b"PK", "FORM B xlsx not produced"
    print("form-b xlsx built, size:", len(dl.data), "bytes")

    # with an official catalogue selected -> comparison report
    r2 = client.post(f"/api/job/{job_id}/form-b/generate", json={"catalog": fb["catalogs"][0]["name"]})
    assert r2.json.get("catalog_compare"), r2.json
    print("form-b catalogue compare:", r2.json["catalog_compare"]["name"])

    # 5. generate Excel
    r = client.post(f"/api/job/{job_id}/confirm", json={"school_id": 10, "overrides": {}})
    assert r.json.get("ok"), r.json
    d = wait_for(client, f"/api/job/{job_id}", "status", "done")
    dl = client.get(f"/api/job/{job_id}/download")
    assert dl.status_code == 200 and dl.data[:2] == b"PK", "xlsx not produced"
    print("excel built, size:", len(dl.data), "bytes")

    # 6. results workbook must include the Form B sheet
    import io as _io, openpyxl
    wb = openpyxl.load_workbook(_io.BytesIO(dl.data))
    assert "Form B" in wb.sheetnames, wb.sheetnames
    fws = wb["Form B"]
    assert fws.cell(1, 1).value == "Course Code"
    print("results workbook sheets:", wb.sheetnames)

    server.shutdown()
    print("\nAPI FLOW PASSED")


if __name__ == "__main__":
    main()
