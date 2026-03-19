import os
import functools
from io import BytesIO
from flask import (Flask, request, redirect, url_for, render_template,
                   flash, send_file, Response)
from datetime import datetime
from quote_engine import generate_pdf

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "msp-dev-key-change-in-production")
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB


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
    return render_template("index.html")


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
        pdf_buf = generate_pdf(excel_bytes)
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

    return send_file(
        pdf_buf,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=download_name,
    )


@app.errorhandler(413)
def too_large(e):
    flash("File too large. Maximum upload size is 20 MB.", "error")
    return redirect(url_for("index")), 413


if __name__ == "__main__":
    app.run(debug=False, port=5000)
