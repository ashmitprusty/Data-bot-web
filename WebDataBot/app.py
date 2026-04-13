from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
import requests
import io
import json
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)

# --- CONFIGURATION ---
API_KEY = 'helloworld' # Replace with your OCR.space API key
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'pdf', 'csv', 'xlsx', 'xls', 'json'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def process_file(filepath, filename):
    ext = filename.rsplit('.', 1)[1].lower()
    df = None

    try:
        if ext in ['xlsx', 'xls']:
            df = pd.read_excel(filepath, header=None)
        elif ext == 'csv':
            df = pd.read_csv(filepath, header=None, on_bad_lines='skip')
        elif ext == 'json':
            df = pd.read_json(filepath)
        elif ext in ['png', 'jpg', 'jpeg', 'pdf']:
            url = "https://api.ocr.space/parse/image"
            payload = {'apikey': API_KEY, 'isTable': True, 'scale': True}
            with open(filepath, 'rb') as f:
                response = requests.post(url, files={'file': f}, data=payload)
            
            result = response.json()
            if result.get('IsErroredOnProcessing'):
                return None, result.get('ErrorMessage')[0]
                
            text = result['ParsedResults'][0]['ParsedText']
            lines = [line.split('\t') for line in text.split('\r\n') if line.strip()]
            df = pd.DataFrame(lines)
        
        # --- CLEANING LOGIC ---
        if df is not None and not df.empty:
            df.dropna(how='all', inplace=True) 
            df.dropna(axis=1, how='all', inplace=True)
            valid_rows = df.dropna(thresh=min(3, len(df.columns))) 
            
            if not valid_rows.empty:
                header_idx = valid_rows.index[0]
                new_header = df.loc[header_idx] 
                df = df.loc[header_idx + 1:] 
                df.columns = new_header 
                
            df.reset_index(drop=True, inplace=True)
            df.columns = [str(col).strip().upper() for col in df.columns]
            df = df.fillna("") 
            return df, None
            
        return None, "Could not extract data."

    except Exception as e:
        return None, str(e)

# --- WEB ROUTES ---

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'})
        
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        temp_dir = tempfile.gettempdir()
        filepath = os.path.join(temp_dir, filename)
        file.save(filepath)
        
        df, error = process_file(filepath, filename)
        os.remove(filepath)
        
        if error:
            return jsonify({'error': error})
            
        table_html = df.to_html(classes='min-w-full text-left text-sm font-light', index=False, border=0)
        return jsonify({'message': 'Success', 'table': table_html, 'json_data': df.to_dict('records')})

    return jsonify({'error': 'Invalid file type'})

@app.route('/export', methods=['POST'])
def export_data():
    """ Handle New Excel / CSV Export """
    data = request.json.get('data')
    format_type = request.json.get('format')
    
    df = pd.DataFrame(data)
    output = io.BytesIO()
    
    if format_type == 'excel':
        df.to_excel(output, index=False, engine='openpyxl')
        mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        filename = 'Cleaned_Data.xlsx'
    else:
        df.to_csv(output, index=False)
        mimetype = 'text/csv'
        filename = 'Cleaned_Data.csv'
        
    output.seek(0)
    return send_file(output, mimetype=mimetype, as_attachment=True, download_name=filename)

@app.route('/append_export', methods=['POST'])
def append_export():
    """ Appends data to an existing uploaded Excel file """
    if 'existing_file' not in request.files:
        return jsonify({'error': 'No existing file provided'})
        
    existing_file = request.files['existing_file']
    new_data_json = request.form.get('new_data')
    
    if not new_data_json or existing_file.filename == '':
        return jsonify({'error': 'Missing data or file'})
        
    try:
        data_list = json.loads(new_data_json)
        df_new = pd.DataFrame(data_list)
        
        # Load the user's existing excel file
        wb = load_workbook(existing_file)
        ws = wb.active
        
        # Append rows without headers
        for r in dataframe_to_rows(df_new, index=False, header=False): 
            ws.append(r)
            
        # Save to memory and send back to user
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Prepend 'Updated_' to the original filename
        download_name = 'Updated_' + secure_filename(existing_file.filename)
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=download_name)
    except Exception as e:
        return jsonify({'error': f'Failed to append: {str(e)}'})

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
