"""
Gold Procurement - Web UI (Flask)
===================================
Environment variables (required in production):
    APP_USERNAME   - login username  (default: admin)
    APP_PASSWORD   - login password  (default: titan@123)
    SECRET_KEY     - Flask session secret (generate a long random string)
"""

import os
import uuid
import tempfile
import traceback
import functools
from pathlib import Path

from flask import (
    Flask, request, jsonify, send_file, render_template,
    abort, session, redirect, url_for
)

import sys
sys.path.insert(0, str(Path(__file__).parent))
from parse_gold_report import parse_txt, write_excel_to_buffer

_EXCEL_CACHE: dict[str, tuple[bytes, str]] = {}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
app.secret_key = os.environ.get("SECRET_KEY", "change-me-in-production-use-a-long-random-string")

_USERNAME = os.environ.get("APP_USERNAME", "admin")
_PASSWORD = os.environ.get("APP_PASSWORD", "titan@123")


def login_required(f):
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            if request.method == "POST":
                return jsonify({"error": "Session expired. Please refresh and log in again."}), 401
            return redirect(url_for("login", next=request.path))
        return f(*args, **kwargs)
    return decorated


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if username == _USERNAME and password == _PASSWORD:
            session["logged_in"] = True
            session["username"]  = username
            next_url = request.args.get("next") or url_for("index")
            return redirect(next_url)
        error = "Invalid username or password."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def index():
    return render_template("index.html", username=session.get("username", ""))


@app.route("/upload", methods=["POST"])
@login_required
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file part in request"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "No file selected"}), 400
    if not f.filename.lower().endswith(".txt"):
        return jsonify({"error": "Only .txt report files are supported"}), 400

    suffix = Path(f.filename).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, mode="wb") as tmp:
        f.save(tmp)
        tmp_path = tmp.name

    try:
        df = parse_txt(tmp_path)
        if df.empty:
            return jsonify({"error": "No data rows could be parsed from this file"}), 422

        import re as _re
        month_label = _re.sub(r"['\s]+", "_", Path(f.filename).stem).upper()
        df["month"] = month_label

        excel_bytes, stats = write_excel_to_buffer(df)

        token     = str(uuid.uuid4())
        stem      = Path(f.filename).stem
        safe_name = f"GoldProcurement_{stem.replace(chr(39), '').replace(' ', '_')}.xlsx"
        _EXCEL_CACHE[token] = (excel_bytes, safe_name)

        if len(_EXCEL_CACHE) > 10:
            oldest = next(iter(_EXCEL_CACHE))
            del _EXCEL_CACHE[oldest]

        return jsonify({
            "token":          token,
            "filename":       safe_name,
            "rows":           stats["rows"],
            "purchase_value": stats["purchase_value"],
            "tax_amount":     stats["tax_amount"],
            "qty_gms":        stats["qty_gms"],
            "unique_vendors": stats["unique_vendors"],
            "month":          stats["month"],
            "vendor_types":   stats["vendor_types"],
        })

    except Exception as exc:
        traceback.print_exc()
        return jsonify({"error": str(exc)}), 500

    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


@app.route("/download/<token>")
@login_required
def download(token: str):
    if token not in _EXCEL_CACHE:
        abort(404)
    excel_bytes, filename = _EXCEL_CACHE[token]
    import io
    return send_file(
        io.BytesIO(excel_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port     = int(os.environ.get("PORT", 5000))
    is_local = os.environ.get("RAILWAY_ENVIRONMENT") is None

    if is_local:
        import webbrowser, threading
        def _open():
            import time; time.sleep(1)
            webbrowser.open(f"http://localhost:{port}")
        threading.Thread(target=_open, daemon=True).start()

    print(f"\n  Gold Procurement Web UI")
    print(f"  Running on port {port}")
    print(f"  Default login: admin / titan@123\n")
    app.run(host="0.0.0.0", port=port, debug=False)
