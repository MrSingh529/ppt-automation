
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify, Response
import os
import tempfile
import shutil
from werkzeug.utils import secure_filename
from datetime import datetime
import traceback
import time
import gc
import json
import threading
from io import BytesIO
import uuid

# Import your existing script functions
from main_script import main as generate_ppt

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('FLASK_SECRET_KEY', 'dev-key-change-in-production')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

ALLOWED_EXTENSIONS = {'xlsx', 'pptx'}

# Progress tracking globals (in-memory)
progress_data = {}
progress_lock = threading.Lock()

def update_progress(session_id, step, status='active', message='', file_path=None):
    """Update progress for a specific session"""
    with progress_lock:
        if session_id not in progress_data:
            progress_data[session_id] = {}
        progress_data[session_id] = {
            'step': step,
            'status': status,
            'message': message,
            'file_path': file_path,
            'timestamp': datetime.now().isoformat()
        }
        print(f"Progress update: Session {session_id}, Step {step}, Status {status}, Message: {message}")

def get_progress(session_id):
    """Get current progress for a session"""
    with progress_lock:
        return progress_data.get(session_id, {'step': 0, 'status': 'waiting', 'message': ''})

def cleanup_progress(session_id):
    """Clean up progress data for a session"""
    with progress_lock:
        if session_id in progress_data:
            del progress_data[session_id]

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    """Generate presentation using temporary files only"""
    excel_file = None
    ppt_file = None
    temp_dir = None

    try:
        # Validate files
        if 'excel_file' not in request.files or 'ppt_file' not in request.files:
            return jsonify({'error': 'Please select both Excel and PowerPoint files'}), 400

        excel_file = request.files['excel_file']
        ppt_file = request.files['ppt_file']

        if excel_file.filename == '' or ppt_file.filename == '':
            return jsonify({'error': 'Please select both files'}), 400

        if not (allowed_file(excel_file.filename) and allowed_file(ppt_file.filename)):
            return jsonify({'error': 'Please upload valid Excel (.xlsx) and PowerPoint (.pptx) files'}), 400

        # Create temporary directory
        temp_dir = tempfile.mkdtemp()

        # Generate unique filenames
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        session_id = str(uuid.uuid4())

        excel_filename = f"excel_{timestamp}_{secure_filename(excel_file.filename)}"
        ppt_filename = f"ppt_{timestamp}_{secure_filename(ppt_file.filename)}"
        output_filename = f"generated_presentation_{timestamp}.pptx"

        # Save files to temporary directory
        excel_path = os.path.join(temp_dir, excel_filename)
        ppt_path = os.path.join(temp_dir, ppt_filename)
        output_path = os.path.join(temp_dir, output_filename)

        excel_file.save(excel_path)
        ppt_file.save(ppt_path)

        print(f"Processing: {excel_path} + {ppt_path} -> {output_path}")

        # Call your main processing function
        generate_ppt(excel_path, ppt_path, output_path)

        # Verify the output file exists
        if not os.path.exists(output_path):
            raise Exception("Output file was not created")

        print(f"Generation completed successfully: {output_path}")

        # Read file into memory for sending
        with open(output_path, 'rb') as f:
            file_data = f.read()

        # Create BytesIO object
        file_io = BytesIO(file_data)
        file_io.seek(0)

        # Send the file from memory
        return send_file(
            file_io,
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        error_msg = f"Generation failed: {str(e)}"
        print(f"Error: {traceback.format_exc()}")
        return jsonify({'error': error_msg}), 500

    finally:
        # Clean up temporary directory
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                print(f"Cleaned up temporary directory: {temp_dir}")
            except Exception as e:
                print(f"Warning: Could not clean up temp directory: {e}")

@app.route('/health')
def health_check():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)