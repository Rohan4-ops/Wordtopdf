from flask import Flask, request, send_file, jsonify, send_from_directory
from pathlib import Path
import tempfile
import subprocess

app = Flask(__name__, static_folder='static')

# Limit upload size (20MB)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024

@app.route("/")
def home():
    return "App is running!"


# ==============================
# Conversion Functions
# ==============================

def convert_with_docx2pdf(input_path: Path, output_path: Path) -> bool:
    try:
        from docx2pdf import convert
        convert(str(input_path), str(output_path))
        return True
    except Exception as e:
        print("docx2pdf failed:", e)
        return False


def convert_with_libreoffice(input_path: Path, output_dir: Path) -> bool:
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
            stderr=subprocess.PIPE
        )

        if result.returncode != 0:
            print("LibreOffice error:", result.stderr.decode())

        return result.returncode == 0

    except Exception as e:
        print("LibreOffice failed:", e)
        return False


def convert_file(input_path: Path, output_dir: Path):
    input_path = Path(input_path)
    output_dir = Path(output_dir)

    if not input_path.exists():
        return None

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / (input_path.stem + ".pdf")

    # Try docx2pdf first
    if convert_with_docx2pdf(input_path, output_path):
        return output_path

    # Fallback to LibreOffice
    if convert_with_libreoffice(input_path, output_dir):
        return output_path

    return None


# ==============================
# Routes
# ==============================

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


@app.route('/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file uploaded"}), 400

        file = request.files['file']

        if not file.filename.lower().endswith('.docx'):
            return jsonify({"error": "Only .docx files allowed"}), 400

        # temp working dir
        temp_dir = tempfile.mkdtemp()

        input_path = Path(temp_dir) / file.filename
        file.save(input_path)

        # convert
        output_path = convert_file(input_path, temp_dir)

        if not output_path or not output_path.exists():
            return jsonify({"error": "Conversion failed"}), 500

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ==============================
# Run App
# ==============================

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)