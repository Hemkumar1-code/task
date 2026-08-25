with open('e:/HEM/task/webapp/app.py', 'r', encoding='utf-8') as f:
    code = f.read()

api_routes = '''
@app.route('/api/upload-idfl', methods=['POST'])
def upload_idfl():
    if 'idfl_file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['idfl_file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
        
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'uploaded_idfl.xlsx')
    file.save(filepath)
    
    # Parse it
    try:
        import xlrd, openpyxl
        wb = openpyxl.load_workbook(filepath, data_only=True)
        stock = []
        if 'NON-IDFL' in wb.sheetnames:
            ws = wb['NON-IDFL']
            for i, row in enumerate(ws.iter_rows(min_row=5, values_only=True)):
                if not row[1] or not row[3]: continue
                tc = str(row[1]).strip()
                if tc.upper() == 'TC NUMBER': continue
                try: rem = float(row[2])
                except: continue
                prod = str(row[3]).strip()
                stock.append({
                    'id': f'non_idfl_{i}',
                    'sheet': 'NON-IDFL',
                    'tc_number': tc,
                    'remaining_weight': rem,
                    'products': prod,
                    'status': 'Active'
                })
        if 'IDFL' in wb.sheetnames:
            ws = wb['IDFL']
            for i, row in enumerate(ws.iter_rows(min_row=7, values_only=True)):
                if not row[2] or not row[6]: continue
                tc = str(row[2]).strip()
                if tc.upper() == 'TC NUMBER': continue
                try: rem = float(row[5])
                except: continue
                prod = str(row[6]).strip()
                stock.append({
                    'id': f'idfl_{i}',
                    'sheet': 'IDFL',
                    'tc_number': tc,
                    'remaining_weight': rem,
                    'products': prod,
                    'status': str(row[3]).strip() if row[3] else 'Active'
                })
        
        database.save_idfl_stock(stock)
        return jsonify({'status': 'success', 'records': len(stock)})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
'''

code = code.replace("if __name__ == '__main__':", api_routes + "\nif __name__ == '__main__':")

with open('e:/HEM/task/webapp/app.py', 'w', encoding='utf-8') as f:
    f.write(code)
