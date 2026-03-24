import os
import functools
from io import BytesIO
from flask import (Flask, request, redirect, url_for, render_template,
                   flash, send_file, Response)
from datetime import datetime
from quote_engine import generate_pdf
from dxf_engine import parse_dxf_to_excel

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "msp-dev-key-change-in-production")
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB (DXF files can be large)
app.config["TEMPLATES_AUTO_RELOAD"] = True

# Set DEPLOYED_AT env var at deploy time, or fall back to process start time
_raw = os.environ.get("DEPLOYED_AT")
if _raw:
    try:
        DEPLOYED_AT = datetime.fromisoformat(_raw).strftime("%-d %B %Y at %H:%M")
    except ValueError:
        DEPLOYED_AT = _raw
else:
    DEPLOYED_AT = datetime.now().strftime("%-d %B %Y at %H:%M")


# ── Simple password protection ────────────────────────────────────────────────
def check_auth(username, password):
    expected_user = os.environ.get("APP_USER", "msp")
    expected_pass = os.environ.get("APP_PASSWORD", "")
    if not expected_pass:
        return True  # no password set → open (fine for localhost dev)
    return username == expected_user and password == expected_pass

def require_auth(f):
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        # If no APP_PASSWORD is set, allow through (localhost dev)
        if not os.environ.get("APP_PASSWORD"):
            return f(*args, **kwargs)
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return Response(
                "Please log in.",
                401,
                {"WWW-Authenticate": 'Basic realm="MSP Quote Generator"'},
            )
        return f(*args, **kwargs)
    return decorated


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
@require_auth
def index():
    return render_template("index.html", deployed_at=DEPLOYED_AT)


@app.route("/generate", methods=["POST"])
@require_auth
def generate():
    f = request.files.get("excel_file")

    if not f or f.filename == "":
        flash("No file selected.", "error")
        return redirect(url_for("index"))

    if not f.filename.lower().endswith(".xlsx"):
        flash("Please upload an .xlsx file.", "error")
        return redirect(url_for("index"))

    try:
        excel_bytes = BytesIO(f.read())
        pdf_buf, pdf_warnings = generate_pdf(excel_bytes)
    except KeyError as e:
        flash(f"Could not find sheet {e} in the workbook. "
              "Make sure the file has 'Client Summary (2)' and 'MSP LMS' tabs.", "error")
        return redirect(url_for("index"))
    except Exception as e:
        app.logger.exception("PDF generation failed")
        flash(f"Something went wrong generating the PDF: {e}", "error")
        return redirect(url_for("index"))

    stem = f.filename.rsplit(".", 1)[0]
    download_name = f"{stem}_quote_{datetime.today().strftime('%Y%m%d')}.pdf"

    import json
    response = send_file(
        pdf_buf,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=download_name,
    )
    if pdf_warnings:
        response.headers["X-PDF-Warnings"] = json.dumps(pdf_warnings)
    return response


@app.route("/dxf")
@require_auth
def dxf_index():
    return render_template("dxf.html", deployed_at=DEPLOYED_AT)


def _collect_dxf_files(uploaded_files):
    """Return list of (filename, bytes) from uploaded .dxf and/or .zip files."""
    import zipfile
    result = []
    for f in uploaded_files:
        name_lower = f.filename.lower()
        if name_lower.endswith(".dxf"):
            result.append((f.filename, f.read()))
        elif name_lower.endswith(".zip"):
            raw = f.read()
            try:
                with zipfile.ZipFile(BytesIO(raw)) as zf:
                    dxf_names = [n for n in zf.namelist()
                                 if n.lower().endswith(".dxf") and not n.startswith("__MACOSX")]
                    if not dxf_names:
                        raise ValueError(f"No .dxf files found inside '{f.filename}'.")
                    for name in dxf_names:
                        result.append((os.path.basename(name), zf.read(name)))
            except zipfile.BadZipFile:
                raise ValueError(f"'{f.filename}' is not a valid zip file.")
        else:
            raise ValueError(f"Unsupported file type: '{f.filename}'. Upload .dxf or .zip files.")
    return result


@app.route("/dxf/generate", methods=["POST"])
@require_auth
def dxf_generate():
    uploaded = [f for f in request.files.getlist("dxf_files") if f and f.filename]

    if not uploaded:
        flash("No files selected.", "error")
        return redirect(url_for("dxf_index"))

    try:
        files = _collect_dxf_files(uploaded)
        excel_buf = parse_dxf_to_excel(files)
    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("dxf_index"))
    except Exception as e:
        app.logger.exception("DXF parsing failed")
        flash(f"Something went wrong parsing the DXF: {e}", "error")
        return redirect(url_for("dxf_index"))

    stem = os.path.splitext(uploaded[0].filename)[0] if len(uploaded) == 1 else "drawings"
    download_name = f"{stem}_measurements_{datetime.today().strftime('%Y%m%d')}.xlsx"

    return send_file(
        excel_buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=download_name,
    )


@app.errorhandler(413)
def too_large(e):
    flash("File too large. Maximum upload size is 200 MB.", "error")
    dest = url_for("dxf_index") if "/dxf" in request.path else url_for("index")
    return redirect(dest), 413


if __name__ == "__main__":
    app.run(debug=False, port=3000)
