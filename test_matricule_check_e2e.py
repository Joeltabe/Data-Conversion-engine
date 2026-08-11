# -*- coding: utf-8 -*-
"""End-to-end smoke test: runs the real checker against a mock portal server.
Run: python test_matricule_check_e2e.py
"""
import http.server
import os
import sys
import threading

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import app  # noqa: E402

PORTAL_DB = [
    # (matricule, name, matching keywords)
    ("LMU-24NUR001", "ASEH IYA MAKANE", ["MAKANE", "ASEH"]),
    ("LMU24NUR014", "NGWA LEM JANET", ["NGWA", "JANET"]),
    ("LMU24NUR016", "EYAMBE MBOTAKE DAUGHTER", ["EYAMBE", "DAUGHTER"]),
    ("LMU24NUR888", "NGWA JOEL ATANGA", ["NGWA", "JOEL"]),
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


def main():
    server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    port = server.server_address[1]
    threading.Thread(target=server.serve_forever, daemon=True).start()

    app.API_BASE = f"http://127.0.0.1:{port}/matriculeCheck.php"
    app.MC_RATE_INTERVAL = 0.0
    app._RateLimiter.__init__.__defaults__ = (0.0,)

    students = [
        {"matricule": "LMU-24NUR001", "name": "ASEH IYA MAKANE", "academic_year": "2023/2024"},
        {"matricule": "LMU24NUR014", "name": "NGWA LEM JANET", "academic_year": "2023/2024"},
        {"matricule": "24NURO16", "name": "EYAMBE MBOTAKE DAUGHTER", "academic_year": "2023/2024"},
        {"matricule": "LMU24NUR999", "name": "NGWA LEM JANET", "academic_year": "2023/2024"},
        {"matricule": "LMU24NUR777", "name": "FON INNOCENT TABE", "academic_year": "2023/2024"},
        {"matricule": "", "name": ""},
    ]

    results = app._run_matricule_check(students, years=["2023/2024"])

    # academic year list sanity
    import datetime
    now = datetime.datetime.now()
    cur = now.year if now.month >= 9 else now.year - 1
    years = app.academic_years()
    assert years and years[0] == f"{cur}/{cur+1}", years
    assert len(years) == 16, len(years)

    expected = ["verified", "verified", "verified", "mismatch", "not_found", "skipped"]
    failed = 0
    for got, want, s in zip(results, expected, students):
        mark = "OK " if got["status"] == want else "FAIL"
        if got["status"] != want:
            failed += 1
        print(f"  {mark} {s['name'] or '(no name)':28} -> {got['status']:10} (want {want})"
              f"{'  best=' + got['api_matricule'] if got['api_matricule'] else ''}")
    print(f"\n{'ALL PASSED' if not failed else str(failed) + ' FAILED'}")
    server.shutdown()
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
