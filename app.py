"""Dextra Document Translator - Flask web front end."""
import io
import os
import re
import secrets
import threading

from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

import docx_translate as dt

DEEPL_API_KEY = dt.load_api_key()  # from the environment or .env, never from code
if not DEEPL_API_KEY:
    raise SystemExit("DEEPL_API_KEY is not set. Add it to .env (see .env.example).")

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB uploads

client = dt.make_client(DEEPL_API_KEY)

DEFAULT_LANGUAGE = 'FR'

# Every target language DeepL offers, loaded once at start-up (restart the app
# to pick up newly added languages). Falls back to the built-in list offline.
try:
    LANGUAGES = dt.fetch_target_languages(client)
except Exception as e:
    app.logger.warning("Could not load DeepL language list (%s); using built-in list.", e)
    LANGUAGES = sorted(dt.TARGET_LANGUAGES, key=lambda cn: cn[1])
LANGUAGE_CODES = {code for code, _ in LANGUAGES}

# Progress per translation job, keyed by an id the browser generates, so
# simultaneous users don't see (or reset) each other's progress.
JOB_ID = re.compile(r"^[A-Za-z0-9-]{8,64}$")
jobs: dict[str, dict] = {}
jobs_lock = threading.Lock()


def set_progress(job_id, done, total):
    if job_id:
        with jobs_lock:
            jobs[job_id] = {"current": done, "total": total}


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_template('index.html', languages=LANGUAGES, default_language=DEFAULT_LANGUAGE)

    uploaded_file = request.files.get('file')
    lang = request.form.get('language', '')
    job_id = request.form.get('job_id', '')
    job_id = job_id if JOB_ID.match(job_id) else None

    if not uploaded_file or not uploaded_file.filename:
        flash("❌ Please upload a file.")
        return redirect(url_for('index'))
    base_name, ext = os.path.splitext(secure_filename(uploaded_file.filename) or "document.docx")
    if ext.lower() != '.docx':
        flash("❌ Only .docx files are supported.")
        return redirect(url_for('index'))
    if lang not in LANGUAGE_CODES:
        flash("❌ Please choose a target language.")
        return redirect(url_for('index'))

    source = io.BytesIO(uploaded_file.read())
    result = io.BytesIO()
    try:
        dt.translate_docx(source, result, lang, client,
                          progress=lambda done, total: set_progress(job_id, done, total))
    except Exception as e:
        app.logger.exception("Translation failed")
        flash(f"❌ Translation failed: {e}")
        return redirect(url_for('index'))
    finally:
        with jobs_lock:
            jobs.pop(job_id, None)

    result.seek(0)
    return send_file(
        result,
        as_attachment=True,
        download_name=f"{base_name}_{lang.lower()}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route('/progress/<job_id>')
def progress(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
    if job is None:  # not started yet, or already finished
        return jsonify({'total': 0, 'current': 0, 'active': False})
    return jsonify({**job, 'active': True})


@app.route('/quota')
def quota():
    try:
        usage = client.get_usage().character
        return jsonify({'used': usage.count, 'limit': usage.limit,
                        'remaining': usage.limit - usage.count})
    except Exception as e:
        app.logger.warning("Quota error: %s", e)
        return jsonify({'error': 'Unable to fetch quota.'})


if __name__ == '__main__':
    app.run(debug=os.getenv("FLASK_DEBUG") == "1")
