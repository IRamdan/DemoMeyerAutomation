from flask import Flask, render_template, request, jsonify
import pandas as pd
import os
import json
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Allowed file extensions
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/convert', methods=['POST'])
def convert():
    try:
        # Check if file is present
        if 'excel' not in request.files:
            return jsonify({"error": "Keine Datei ausgewählt"}), 400
        
        file = request.files['excel']
        
        # Check if file is selected
        if file.filename == '':
            return jsonify({"error": "Keine Datei ausgewählt"}), 400
        
        # Check file extension
        if not allowed_file(file.filename):
            return jsonify({"error": "Nur .xlsx und .xls Dateien sind erlaubt"}), 400
        
        # Read Excel file
        df = pd.read_excel(file)
        
        # Convert to JSON
        json_data = df.to_dict(orient='records')
        
        # Get basic file info
        file_info = {
            "filename": secure_filename(file.filename),
            "rows": len(df),
            "columns": len(df.columns),
            "column_names": list(df.columns)
        }
        
        return jsonify({
            "success": True,
            "data": json_data,
            "file_info": file_info
        })
        
    except Exception as e:
        return jsonify({"error": f"Fehler beim Verarbeiten der Datei: {str(e)}"}), 500

@app.route('/download', methods=['POST'])
def download():
    try:
        data = request.get_json()
        filename = data.get('filename', 'converted_data.json')
        json_data = data.get('data', {})
        
        # Format JSON nicely
        formatted_json = json.dumps(json_data, indent=2, ensure_ascii=False)
        
        return jsonify({
            "success": True,
            "json_content": formatted_json,
            "filename": filename
        })
        
    except Exception as e:
        return jsonify({"error": f"Fehler beim Erstellen der Download-Datei: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)