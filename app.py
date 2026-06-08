"""
Gold Procurement - Web UI (Flask)
===================================
Local-only single page app. No authentication.
Run: python app.py
"""

import os
import uuid
import tempfile
import traceback
from pathlib import Path

from flask import Flask, request, jsonify, send_file, render_template, abort

import sys
sys.path.insert(0, str(Path(__file__).parent))
from parse_gold_report import parse_txt, write_excel_to_buffer

_EXCEL_CACHE: dict[str, tuple[bytes, str]] = {}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


# ─── JSON error helpers ────────────────────────────────────────────────────────

def _json_error(code, message):
    resp = jsonify({"error": message})
    resp.status_code = code
    return resp

@app.errorhandler(400)
def err_400(e): return _json_error(400, f"Bad request: {e.description}")

@app.errorhandler(404)
def err_404(e): return _json_error(404, "Not found.")

@app.errorhandler(413)
def err_413(e): return _json_error(413, "File too large. Maximum size is 50 MB.")

@app.errorhandler(500)
def err_500(e):
    traceback.print_exc()
    return _json_error(500, "Internal server error. Check server logs.")

@app.errorhandler(Exception)
def err_any(e):
    traceback.print_exc()
    return _json_error(500, str(e))


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
    port = int(os.environ.get("PORT", 5000))

    import webbrowser, threading
    def _open():
        import time; time.sleep(1)
        webbrowser.open(f"http://localhost:{port}")
    threading.Thread(target=_open, daemon=True).start()

    print(f"\n  Gold Procurement Web UI")
    print(f"  Running at: http://localhost:{port}\n")
    app.run(host="127.0.0.1", port=port, debug=False)
