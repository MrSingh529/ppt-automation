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
from main_script import main as generate_ppt, set_progress_callback

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('FLASK_SECRET_KEY', 'dev-key-change-in-production')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# For cloud deployment, create temporary folders
UPLOAD_FOLDER = tempfile.mkdtemp()
OUTPUT_FOLDER = tempfile.mkdtemp()
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
        
        # CRITICAL DEBUG: Show the exact data being stored
        print(f"DEBUG: Progress data for {session_id}: {progress_data[session_id]}")

def get_progress(session_id):
    """Get current progress for a session"""
    with progress_lock:
        return progress_data.get(session_id, {'step': 0, 'status': 'waiting', 'message': ''})

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def safe_cleanup(file_paths):
    """Safely cleanup files with retry logic"""
    for file_path in file_paths:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                print(f"Successfully removed: {file_path}")
            except Exception as e:
                print(f"Could not cleanup {file_path}: {e}")

def generate_ppt_with_progress(excel_path, ppt_path, output_path, session_id):
    """Wrapper function to call main script with progress updates"""
    
    # Store session_id in a way that ensures it's captured correctly
    _session_id = session_id  # Create a local copy
    
    def progress_update_callback(step, status, message, file_path=None):
        """Callback function to update progress"""
        print(f"🔔 CALLBACK TRIGGERED: step={step}, status={status}, session={_session_id}")  # DEBUG LINE
        
        # Use output_path for completed status
        final_path = output_path if (status == 'completed' and step == 8) else file_path
        
        # Call the Flask update_progress function
        update_progress(_session_id, step, status, message, final_path)
        
        print(f"✅ UPDATE_PROGRESS CALLED for session {_session_id}")  # DEBUG LINE
    
    # Set the callback in main_script
    set_progress_callback(progress_update_callback)
    
    print(f"📌 Callback set for session: {_session_id}")  # DEBUG LINE
    
    # Call your main function
    generate_ppt(excel_path, ppt_path, output_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    """Generate presentation using temporary files only"""
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

@app.route('/generate-with-progress', methods=['POST'])
def generate_with_progress():
    """Generate presentation with real-time progress updates"""
    session_id = str(uuid.uuid4())
    
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

        # Initialize progress
        update_progress(session_id, 0, 'active', 'Processing files...')
        
        # Save files
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"excel_{timestamp}_{secure_filename(excel_file.filename)}"
        ppt_filename = f"ppt_{timestamp}_{secure_filename(ppt_file.filename)}"
        
        excel_path = os.path.join(UPLOAD_FOLDER, excel_filename)
        ppt_path = os.path.join(UPLOAD_FOLDER, ppt_filename)
        output_filename = f"generated_presentation_{timestamp}.pptx"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        excel_file.save(excel_path)
        ppt_file.save(ppt_path)
        
        def generate_with_real_progress():
            """Background generation function with real progress tracking"""
            try:
                # Step 1: Files already saved
                update_progress(session_id, 1, 'completed', 'Files saved successfully')

                # ✅ IMMEDIATE FIX: Direct callback instead of wrapper
                def direct_progress_callback(step, status, message, file_path=None):
                    print(f"🔔 DIRECT CALLBACK: step={step}, status={status}, session={session_id}")
                    final_path = output_path if (status == 'completed' and step == 8) else file_path
                    update_progress(session_id, step, status, message, final_path)
                    print(f"✅ DIRECT UPDATE_PROGRESS CALLED for session {session_id}")

                # Set callback directly in main_script
                set_progress_callback(direct_progress_callback)

                # Call main function directly
                generate_ppt(excel_path, ppt_path, output_path)
                
            except Exception as e:
                current_step = get_progress(session_id).get('step', 1)
                update_progress(session_id, current_step, 'error', f'Error: {str(e)}')
                print(f"Error in background generation: {traceback.format_exc()}")
                
            finally:
                # Cleanup files (delay for download)
                def delayed_cleanup():
                    time.sleep(30)  # Wait longer for download
                    files_to_cleanup = []
                    if excel_path and os.path.exists(excel_path):
                        files_to_cleanup.append(excel_path)
                    if ppt_path and os.path.exists(ppt_path):
                        files_to_cleanup.append(ppt_path)
                    safe_cleanup(files_to_cleanup)

                cleanup_thread = threading.Thread(target=delayed_cleanup)
                cleanup_thread.daemon = True
                cleanup_thread.start()

        # Start the generation process in a separate thread
        threading.Thread(target=generate_with_real_progress, daemon=True).start()
        
        # Return JSON response with redirect URL
        return jsonify({
            'status': 'success',
            'session_id': session_id,
            'redirect_url': f'/progress-view/{session_id}'
        })
        
    except Exception as e:
        return jsonify({'error': f'Upload failed: {str(e)}'}), 500

@app.route('/progress/<session_id>')
def get_progress_status(session_id):
    """Get progress status for real-time updates"""
    progress = get_progress(session_id)
    return jsonify(progress)

@app.route('/progress-view/<session_id>')
def progress_view(session_id):
    """Render the progress page"""
    return render_template('progress.html', session_id=session_id)

@app.route('/download-generated/<session_id>')
def download_generated_file(session_id):
    """Download the generated file directly"""
    try:
        progress = get_progress(session_id)
        if progress['status'] != 'completed' or progress['step'] != 8:
            return jsonify({'error': 'File not ready for download'}), 400

        file_path = progress.get('file_path')
        if file_path and os.path.exists(file_path):
            filename = os.path.basename(file_path)
            return send_file(
                file_path,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
        else:
            return jsonify({'error': 'Generated file not found'}), 404
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health_check():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)