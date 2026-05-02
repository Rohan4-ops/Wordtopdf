from flask import Flask, request, send_file, jsonify, send_from_directory, after_this_request
from pathlib import Path
import tempfile
import subprocess
import shutil
import os

app = Flask(__name__, static_folder='static')

# Limit upload size (20MB)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024


# ==============================
# Home Route (UI)
# ==============================
@app.route("/")
def index():
    return send_from_directory('static', 'index.html')


# ==============================
# Conversion Functions
# ==============================

def convert_with_docx2pdf(input_path: Path, output_path: Path) -> bool:
    try:
        from docx2pdf import convert
        convert(str(input_path), str(output_path))
        return output_path.exists()
    except Exception as e:
        print("docx2pdf failed:", e)
        return False


def convert_with_libreoffice(input_path: Path, output_dir: Path) -> Path | None:
    try:
        result = subprocess.run(
            [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(output_dir),
                str(input_path)
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=30
        )

        if result.returncode != 0:
            print("LibreOffice error:", result.stderr.decode())
            return None

        output_file = output_dir / (input_path.stem + ".pdf")
        return output_file if output_file.exists() else None

    except Exception as e:
        print("LibreOffice failed:", e)
        return None


def convert_file(input_path: Path, output_dir: Path):
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / (input_path.stem + ".pdf")

    # Try docx2pdf first
    if convert_with_docx2pdf(input_path, output_path):
        return output_path

    # Fallback to LibreOffice
    return convert_with_libreoffice(input_path, output_dir)


# ==============================
# Convert Route
# ==============================

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']

    if not file.filename.lower().endswith('.docx'):
        return jsonify({"error": "Only .docx files allowed"}), 400

    # Create temp dir
    temp_dir = Path(tempfile.mkdtemp())
    input_path = temp_dir / file.filename
    file.save(input_path)

    output_path = convert_file(input_path, temp_dir)

    if not output_path or not output_path.exists():
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({"error": "Conversion failed"}), 500

    @after_this_request
    def cleanup(response):
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception as e:
            print("Cleanup error:", e)
        return response

    return send_file(output_path, as_attachment=True)


# ==============================
# Run App
# ==============================

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)