"""
Gold Procurement – Web UI (Flask)
===================================
Provides a browser-based interface to upload the monthly Oracle TXT report
and download a formatted Excel file.

Run:
    python app.py
Then open http://localhost:5000 in your browser.
"""

import os
import uuid
import tempfile
import traceback
from pathlib import Path

from flask import (
    Flask, request, jsonify, send_file, render_template, abort
)

# Import the core parser from the same directory
import sys
sys.path.insert(0, str(Path(__file__).parent))
from parse_gold_report import parse_txt, write_excel_to_buffer

# ─── in-memory store: token → (excel_bytes, download_filename) ───────────────
_EXCEL_CACHE: dict[str, tuple[bytes, str]] = {}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024   # 50 MB max upload


# ─── Routes ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file part in request"}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "No file selected"}), 400
    if not f.filename.lower().endswith(".txt"):
        return jsonify({"error": "Only .txt report files are supported"}), 400

    # Save to a temp file so parse_txt can read it (it needs a real path)
    suffix = Path(f.filename).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, mode="wb") as tmp:
        f.save(tmp)
        tmp_path = tmp.name

    try:
        # Parse the report
        df = parse_txt(tmp_path)

        if df.empty:
            return jsonify({"error": "No data rows could be parsed from this file"}), 422

        # Fix month label: use the original filename, not the temp path
        import re as _re
        month_label = _re.sub(r"['\s]+", "_", Path(f.filename).stem).upper()
        df["month"] = month_label

        # Generate Excel in memory
        excel_bytes, stats = write_excel_to_buffer(df)

        # Store in cache with a unique token
        token = str(uuid.uuid4())
        stem  = Path(f.filename).stem
        safe_name = f"GoldProcurement_{stem.replace(chr(39), '').replace(' ', '_')}.xlsx"
        _EXCEL_CACHE[token] = (excel_bytes, safe_name)

        # Keep cache small – drop old entries beyond last 10
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


# ─── Entry point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    is_local = port == 5000 and os.environ.get("RAILWAY_ENVIRONMENT") is None

    if is_local:
        import webbrowser, threading
        def _open():
            import time; time.sleep(1)
            webbrowser.open(f"http://localhost:{port}")
        threading.Thread(target=_open, daemon=True).start()

    print(f"\n  Gold Procurement Web UI")
    print(f"  ─────────────────────────")
    print(f"  Running on port {port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)
