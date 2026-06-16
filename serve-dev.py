"""
Sonia Outreach Dashboard — local dev server.

Serves index.html on http://localhost:62672 so you can test against the live
SharePoint lists before deploying to Azure. index.html is the single source of
truth — the same file Azure Static Web Apps serves. Port 62672 is fixed on
purpose: the exact page URL is the SPA redirect URI you register in the Entra
app, so Microsoft sign-in only works on this exact address.

Register this EXACT value as a Single-page application redirect URI on the
Entra app (da86319d-… — same app as the ITR Hub):

    http://localhost:62672/index.html

(Add it under Authentication -> Single-page application, alongside the hub's
http://localhost:62671/ entry. Do NOT add it under the "Web" platform.)

Usage:
    python serve-dev.py
    then open http://localhost:62672/index.html
    (append ?demo for synthetic data without signing in; Ctrl+C to stop)
"""
import http.server
import os
import socketserver
import sys

PORT = 62672
DIR = os.path.dirname(os.path.abspath(__file__))


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=DIR, **kw)

    def handle_one_request(self):
        self.connection.settimeout(30)
        super().handle_one_request()

    def end_headers(self):
        # Dev server: never cache, so edits to the HTML always load fresh.
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def log_message(self, fmt, *args):
        try:
            print(f"  {args[0]}  {args[1]}")
        except Exception:
            super().log_message(fmt, *args)


class ThreadingServer(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True
    allow_reuse_address = False


def main():
    if not os.path.exists(os.path.join(DIR, "index.html")):
        print("  ERROR: index.html not found next to this script.")
        sys.exit(1)
    if not os.path.exists(os.path.join(DIR, "msal-browser.min.js")):
        print("  WARNING: msal-browser.min.js not found next to this script — sign-in will fail.")

    print(f"\n  Outreach :  http://localhost:{PORT}/index.html")
    print(f"  Folder   :  {DIR}")
    print(f"  Redirect :  register http://localhost:{PORT}/index.html as a SPA redirect URI")
    print(f"  Demo     :  http://localhost:{PORT}/index.html?demo  (no sign-in)")
    print(f"  Ctrl+C to stop\n")

    try:
        ThreadingServer(("127.0.0.1", PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        print("\n  Stopped.")
    except OSError as e:
        print(f"\n  ERROR: could not bind port {PORT} — {e}")
        print(f"  Another process may already be using it. Close it and retry.")
        sys.exit(1)


if __name__ == "__main__":
    main()
