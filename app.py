import os
import io
from flask import Flask, render_template, request, jsonify, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import webbrowser
import threading
import sys

if getattr(sys, 'frozen', False):
    bundle_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    app = Flask(
        __name__,
        template_folder=os.path.join(bundle_dir, 'templates'),
        static_folder=os.path.join(bundle_dir, 'static')
    )
    base_dir = os.path.dirname(sys.executable)
else:
    app = Flask(__name__)
    base_dir = os.path.dirname(os.path.abspath(__file__))

app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB file limit

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'app': 'Excel Web Pro', 'version': '2.0.0'})

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in the request'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    allowed_exts = ('.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.csv')
    if not file.filename.lower().endswith(allowed_exts):
        return jsonify({'error': 'Unsupported file format. Please upload an Excel (.xlsx, .xls) or CSV file.'}), 400
    
    try:
        if file.filename.lower().endswith('.csv'):
            import csv
            content = file.read().decode('utf-8-sig', errors='ignore')
            reader = csv.reader(io.StringIO(content))
            sheet_data = []
            for row in reader:
                sheet_data.append([c for c in row])
            if not sheet_data:
                sheet_data = [["" for _ in range(12)] for _ in range(25)]
            return jsonify({
                'filename': file.filename,
                'sheets': {'Sheet1': sheet_data},
                'sheetNames': ['Sheet1']
            })

        # Load workbook (data_only=False to preserve formulas)
        wb = openpyxl.load_workbook(file, data_only=False)
        
        sheets = {}
        styles = {}
        col_widths = {}
        
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            sheet_data = []
            sheet_styles = {}
            
            max_row = max(ws.max_row or 0, 25)
            max_col = max(ws.max_column or 0, 12)
            
            # Read cell values and styles
            for r in range(1, max_row + 1):
                row_data = []
                for c in range(1, max_col + 1):
                    cell = ws.cell(row=r, column=c)
                    val = cell.value
                    
                    if val is None:
                        val = ""
                    elif isinstance(val, (int, float)):
                        pass  # Keep numbers as is
                    else:
                        val = str(val)
                    row_data.append(val)
                    
                    # Extract basic styling if present
                    c_style = {}
                    if cell.font:
                        if cell.font.bold: c_style['bold'] = True
                        if cell.font.italic: c_style['italic'] = True
                        if cell.font.underline: c_style['underline'] = True
                        if cell.font.color and cell.font.color.rgb and isinstance(cell.font.color.rgb, str):
                            # rgb format AARRGGBB
                            rgb = cell.font.color.rgb
                            if len(rgb) == 8:
                                c_style['color'] = f"#{rgb[2:]}"
                            elif len(rgb) == 6:
                                c_style['color'] = f"#{rgb}"
                    if cell.fill and cell.fill.fill_type == 'solid' and cell.fill.start_color:
                        rgb = cell.fill.start_color.rgb
                        if rgb and isinstance(rgb, str):
                            if len(rgb) == 8:
                                c_style['bg'] = f"#{rgb[2:]}"
                            elif len(rgb) == 6:
                                c_style['bg'] = f"#{rgb}"
                    if cell.alignment:
                        if cell.alignment.horizontal:
                            c_style['align'] = cell.alignment.horizontal
                        if cell.alignment.vertical:
                            c_style['valign'] = cell.alignment.vertical
                    
                    if c_style:
                        sheet_styles[f"{r-1},{c-1}"] = c_style
                        
                sheet_data.append(row_data)
                
            sheets[sheet_name] = sheet_data
            styles[sheet_name] = sheet_styles
            
            # Column widths
            ws_widths = {}
            for col_letter, col_dim in ws.column_dimensions.items():
                if col_dim.width:
                    ws_widths[col_letter] = col_dim.width
            col_widths[sheet_name] = ws_widths
            
        return jsonify({
            'filename': file.filename,
            'sheets': sheets,
            'styles': styles,
            'colWidths': col_widths,
            'sheetNames': wb.sheetnames
        })
        
    except Exception as e:
        return jsonify({'error': f'Failed to process file: {str(e)}'}), 500

EXPORT_DIR = os.path.join(base_dir, 'exports')
os.makedirs(EXPORT_DIR, exist_ok=True)

def generate_openpyxl_workbook(data):
    filename = data.get('filename', 'export.xlsx')
    import re, unicodedata
    name_str = str(filename).replace('đ', 'd').replace('Đ', 'D')
    name_str = unicodedata.normalize('NFKD', name_str).encode('ascii', 'ignore').decode('utf-8')
    clean_filename = re.sub(r'[<>:"/\\|?*\x00-\x1F]', '_', name_str).rstrip('. ').strip() or 'export.xlsx'
    if not clean_filename.lower().endswith('.xlsx'):
        clean_filename += '.xlsx'
        
    styles_data = data.get('styles', {})
    col_widths_data = data.get('colWidths', {})
    merges_data = data.get('merges', {})
        
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet
    
    thin_side = Side(border_style="thin", color="D1D5DB")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    for sheet_name, sheet_data in data['sheets'].items():
        safe_title = "".join(ch for ch in sheet_name if ch not in r'\/*?:[]')[:31] or "Sheet"
        ws = wb.create_sheet(title=safe_title)
        sheet_styles = styles_data.get(sheet_name, {})
        
        for r_idx, row in enumerate(sheet_data):
            for c_idx, val in enumerate(row):
                cleaned_val = val
                if isinstance(val, str):
                    if val.startswith('='):
                        cleaned_val = val
                    elif val == "":
                        cleaned_val = None
                    else:
                        try:
                            if '.' in val:
                                cleaned_val = float(val)
                            else:
                                cleaned_val = int(val)
                        except ValueError:
                            cleaned_val = val
                
                cell = ws.cell(row=r_idx + 1, column=c_idx + 1, value=cleaned_val)
                
                cell_key = f"{r_idx},{c_idx}"
                if cell_key in sheet_styles:
                    c_style = sheet_styles[cell_key]
                    font_args = {}
                    if c_style.get('bold'): font_args['bold'] = True
                    if c_style.get('italic'): font_args['italic'] = True
                    if c_style.get('underline'): font_args['underline'] = 'single'
                    if c_style.get('fontSize'): font_args['size'] = int(c_style['fontSize'])
                    if c_style.get('fontFamily'): font_args['name'] = c_style['fontFamily']
                    if c_style.get('color'):
                        hex_col = c_style['color'].replace('#', '')
                        if len(hex_col) == 6:
                            font_args['color'] = hex_col
                    if font_args:
                        cell.font = Font(**font_args)
                        
                    if c_style.get('bg'):
                        hex_bg = c_style['bg'].replace('#', '')
                        if len(hex_bg) == 6:
                            cell.fill = PatternFill(start_color=hex_bg, end_color=hex_bg, fill_type='solid')
                            
                    align_args = {}
                    if c_style.get('align'): align_args['horizontal'] = c_style['align']
                    if c_style.get('valign'): align_args['vertical'] = c_style['valign']
                    if c_style.get('wrap'): align_args['wrap_text'] = True
                    if align_args:
                        cell.alignment = Alignment(**align_args)
                        
                    if c_style.get('border'):
                        cell.border = border_all

        if sheet_name in col_widths_data:
            for col_letter, width in col_widths_data[sheet_name].items():
                try:
                    ws.column_dimensions[col_letter].width = float(width)
                except Exception:
                    pass
                
        if sheet_name in merges_data:
            for merge_range in merges_data[sheet_name]:
                try:
                    ws.merge_cells(merge_range)
                except Exception:
                    pass
                    
    return wb, clean_filename

def save_workbook_to_downloads(wb, filename):
    try:
        downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
        if not os.path.exists(downloads_dir):
            downloads_dir = os.path.join(os.environ.get('USERPROFILE', 'C:\\Users\\Admin'), 'Downloads')
        os.makedirs(downloads_dir, exist_ok=True)
        dest_path = os.path.join(downloads_dir, filename)
        wb.save(dest_path)
        print(f"[Export] Saved directly to Downloads: {dest_path}")
        return dest_path
    except Exception as e:
        print(f"[Export] Failed to auto-save to Downloads: {e}")
        return None

@app.route('/download', methods=['POST'])
def download_file():
    data = None
    if request.is_json:
        data = request.get_json()
    else:
        import json
        raw = request.form.get('payload') or request.form.get('data')
        if raw:
            try:
                data = json.loads(raw)
            except Exception:
                data = None

    if not data or 'sheets' not in data:
        return jsonify({'error': 'Invalid request data'}), 400
    
    try:
        wb, filename = generate_openpyxl_workbook(data)
        saved_path = save_workbook_to_downloads(wb, filename)
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        resp = send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        if saved_path:
            resp.headers['X-Saved-Path'] = saved_path
        return resp
    except Exception as e:
        return jsonify({'error': f'Failed to generate Excel file: {str(e)}'}), 500

@app.route('/prepare_export', methods=['POST'])
def prepare_export():
    import json, uuid
    data = None
    if request.is_json:
        data = request.get_json()
    else:
        raw = request.form.get('payload') or request.form.get('data')
        if raw:
            try:
                data = json.loads(raw)
            except Exception:
                data = None

    if not data or 'sheets' not in data:
        return jsonify({'error': 'Invalid request data'}), 400
        
    try:
        wb, filename = generate_openpyxl_workbook(data)
        saved_path = save_workbook_to_downloads(wb, filename)
        token = uuid.uuid4().hex[:12]
        filepath = os.path.join(EXPORT_DIR, f"{token}_{filename}")
        wb.save(filepath)
        
        return jsonify({
            'success': True,
            'download_url': f'/get_file/{token}/{filename}',
            'filename': filename,
            'saved_to': saved_path or filepath
        })
    except Exception as e:
        return jsonify({'error': f'Failed to generate Excel file: {str(e)}'}), 500

@app.route('/get_file/<token>/<filename>', methods=['GET'])
def get_exported_file(token, filename):
    import glob
    matches = glob.glob(os.path.join(EXPORT_DIR, f"{token}_*"))
    if not matches:
        return "File not found or expired", 404
    filepath = matches[0]
    return send_file(
        filepath,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    def open_browser():
        webbrowser.open('http://127.0.0.1:5000')
    threading.Timer(1.0, open_browser).start()
    app.run(host='0.0.0.0', debug=False, port=5000)
