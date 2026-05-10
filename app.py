from flask import Flask, request, send_file, jsonify, send_from_directory, after_this_request
from pathlib import Path
import tempfile
import subprocess
import shutil
import os
import logging

app = Flask(__name__, static_folder='static')

# Limit upload size (20MB)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024


# ==============================
# Logging Setup
# ==============================
LOG_DIR = "/opt/Wordtopdf/logs"
os.makedirs(LOG_DIR, exist_ok=True)

LOG_FILE = os.path.join(LOG_DIR, "app.log")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)
logger.info("Application started")


# ==============================
# Home Route (UI)
# ==============================
@app.route("/")
def index():
    logger.info("Home page accessed")
    return send_from_directory('static', 'index.html')


# ==============================
# Conversion Functions
# ==============================

def convert_with_docx2pdf(input_path: Path, output_path: Path) -> bool:
    try:
        from docx2pdf import convert
        convert(str(input_path), str(output_path))
        success = output_path.exists()
        if success:
            logger.info(f"docx2pdf success: {output_path}")
        return success
    except Exception as e:
        logger.error(f"docx2pdf failed: {e}")
        return False


def convert_with_libreoffice(input_path: Path, output_dir: Path) -> Path | None:
    try:
        logger.info("Trying LibreOffice conversion")

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
            logger.error(f"LibreOffice error: {result.stderr.decode()}")
            return None

        output_file = output_dir / (input_path.stem + ".pdf")

        if output_file.exists():
            logger.info(f"LibreOffice success: {output_file}")
            return output_file

        return None

    except Exception as e:
        logger.exception("LibreOffice conversion failed")
        return None


def convert_file(input_path: Path, output_dir: Path):
    logger.info(f"Starting conversion for: {input_path.name}")

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
    logger.info("Received conversion request")

    if 'file' not in request.files:
        logger.warning("No file uploaded")
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    logger.info(f"Uploaded file: {file.filename}")

    if not file.filename.lower().endswith('.docx'):
        logger.warning("Invalid file type")
        return jsonify({"error": "Only .docx files allowed"}), 400

    # Create temp dir
    temp_dir = Path(tempfile.mkdtemp())
    logger.info(f"Temp directory created: {temp_dir}")

    input_path = temp_dir / file.filename
    file.save(input_path)

    output_path = convert_file(input_path, temp_dir)

    if not output_path or not output_path.exists():
        logger.error("Conversion failed")
        shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({"error": "Conversion failed"}), 500

    @after_this_request
    def cleanup(response):
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
            logger.info(f"Cleaned temp directory: {temp_dir}")
        except Exception:
            logger.exception("Cleanup failed")
        return response

    logger.info(f"Conversion successful: {output_path}")
    return send_file(output_path, as_attachment=True)


# ==============================
# Run App
# ==============================

if __name__ == '__main__':
    logger.info("Starting Flask development server")
    app.run(host='0.0.0.0', port=6000, debug=True)