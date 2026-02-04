import os
import sys
import datetime
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
import reconciliation_v2  # Import the existing logic

# Get the directory of the current file
base_dir = os.path.abspath(os.path.dirname(__file__))

# DEBUGGING: Print directory structure to logs
print(f"DEBUG: base_dir is {base_dir}")
try:
    print(f"DEBUG: Contents of base_dir: {os.listdir(base_dir)}")
    if os.path.exists(os.path.join(base_dir, 'templates')):
        print(f"DEBUG: Contents of templates: {os.listdir(os.path.join(base_dir, 'templates'))}")
    else:
        print("DEBUG: 'templates' folder NOT FOUND in base_dir")
except Exception as e:
    print(f"DEBUG: Error checking files: {e}")

app = Flask(__name__, 
            template_folder=os.path.join(base_dir, 'templates'), 
            static_folder=os.path.join(base_dir, 'static'))

# Configuration
UPLOAD_FOLDER = 'uploads'
RESULTS_FOLDER = 'results'
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'xlsm'}

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULTS_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULTS_FOLDER'] = RESULTS_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{timestamp}_{filename}")
        file.save(save_path)
        
        # Determine sheets if provided
        sheet1 = request.form.get('sheet1')
        sheet2 = request.form.get('sheet2')
        tolerance = float(request.form.get('tolerance', 1.0))
        
        try:
            # Process the file using the imported module
            # We need to capture the output path returned by process_reconciliation
            # Note: process_reconciliation accepts absolute paths usually, let's ensure we pass absolute
            abs_save_path = os.path.abspath(save_path)
            
            # Call the processing function
            # Note: We might need to modify reconciliation_v2 slightly if it prints too much or handle return values better
            # But relying on the existing 'return out_file' at the end of process_reconciliation
            
            output_file = reconciliation_v2.process_reconciliation(
                path=abs_save_path,
                sheet1=sheet1 if sheet1 else None, 
                sheet2=sheet2 if sheet2 else None,
                tolerance=tolerance
            )
            
            if output_file and os.path.exists(output_file):
                # Return the filename for the download link
                return jsonify({
                    'message': 'Reconciliation successful!',
                    'download_url': f'/download?file={os.path.basename(output_file)}'
                })
            else:
                 return jsonify({'error': 'Reconciliation failed to produce an output file.'}), 500

        except Exception as e:
            return jsonify({'error': str(e)}), 500
            
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/download')
def download():
    filename = request.args.get('file')
    if not filename:
         return "File not specified", 400
         
    # The reconciliation tool saves files in the same directory as input usually, 
    # OR we need to check where it saved. 
    # Based on reconciliation_v2.py: out_file = f"{os.path.splitext(path)[0]}_RECON_{ts}.xlsx"
    # So it should be in the UPLOAD_FOLDER
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    # Check if file exists there
    if not os.path.exists(file_path):
         # Try checking the root folder just in case (original script behavior)
        possible_path = os.path.abspath(filename)
        if os.path.exists(possible_path):
            file_path = possible_path
        else:
            return "File not found", 404

    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
