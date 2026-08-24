# app.py — Flask web server for the creditfile credit scoring pipeline

import io
import json
import os
import tempfile

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from creditfile.main import analyze_file

app = Flask(__name__, template_folder='frontend', static_folder='frontend/static')

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}


def _allowed(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# ── Routes ───────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/analyze', methods=['POST'])
def analyze():
    """
    Accept an Excel upload, run the full pipeline, return JSON results.
    The file is written to a temp path so os.path.getmtime can read the
    real last-modified timestamp, matching the original Colab behaviour.
    """
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided.'}), 400

    f = request.files['file']
    if f.filename == '':
        return jsonify({'error': 'No file selected.'}), 400
    if not _allowed(f.filename):
        return jsonify({'error': 'Only .xlsx / .xls files are supported.'}), 400

    filename = secure_filename(f.filename)

    # Write to a temp file so we can read its real mtime
    suffix = '.' + filename.rsplit('.', 1)[-1]
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=suffix)
    try:
        with os.fdopen(tmp_fd, 'wb') as tmp:
            f.save(tmp)

        with open(tmp_path, 'rb') as fh:
            result = analyze_file(fh, filename, tmp_path=tmp_path)

        return jsonify(result)

    except Exception as exc:
        return jsonify({'error': f'Processing failed: {exc}'}), 500

    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app.run(debug=True, port=5000)
