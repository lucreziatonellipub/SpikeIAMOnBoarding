"""
Explicit launcher for the Spike Admin Flask dashboard.

Put this file in the same folder as admin_app_solution_design_fixed_v2.py
and run:
    python -u run_dashboard.py

Optional:
    $env:ADMIN_MODULE="admin_dashboard"   # PowerShell
    python -u run_dashboard.py
"""

import importlib
import os
import sys
import traceback

MODULE_NAME = os.getenv("ADMIN_MODULE", "admin_dashboard")
HOST = os.getenv("HOST", "0.0.0.0")
PORT = int(os.getenv("PORT", "5001"))
DEBUG = os.getenv("FLASK_DEBUG", "1") == "1"

print("=" * 80, flush=True)
print("Starting Spike Admin dashboard launcher", flush=True)
print(f"Python executable: {sys.executable}", flush=True)
print(f"Working directory: {os.getcwd()}", flush=True)
print(f"Module to import: {MODULE_NAME}", flush=True)
print("=" * 80, flush=True)

try:
    module = importlib.import_module(MODULE_NAME)
    app = getattr(module, "app", None)

    if app is None:
        raise RuntimeError(f"Module '{MODULE_NAME}' was imported, but it does not expose a Flask variable named 'app'.")

    print("Flask app imported successfully.", flush=True)
    print("Registered routes:", flush=True)
    print(app.url_map, flush=True)
    print("=" * 80, flush=True)
    print(f"Open: http://localhost:{PORT}", flush=True)
    print(f"Listening on: {HOST}:{PORT}", flush=True)
    print("Press CTRL+C to stop.", flush=True)
    print("=" * 80, flush=True)

    app.run(host=HOST, port=PORT, debug=DEBUG, use_reloader=False)

except Exception:
    print("ERROR while starting the dashboard:", flush=True)
    traceback.print_exc()
    sys.exit(1)
