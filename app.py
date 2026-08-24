import os
import xlrd
import openpyxl
import tempfile
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from flask import Flask, render_template, request, send_file, flash, redirect

app = Flask(__name__)
app.secret_key = 'super_secret_key'
UPLOAD_FOLDER = tempfile.gettempdir()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def process_invoice_files(invoice_paths, output_path):
    styles_data = []
    style_weights = {}
    
    for invoice_path in invoice_paths:
        try:
            wb_in = xlrd.open_workbook(invoice_path)
        except Exception as e:
            continue
            
        # Try to extract Invoice No and Buyer from the first INV sheet
        inv_no = 'Unknown Invoice'
        buyer = 'Unknown Buyer'
        
        for si in range(wb_in.nsheets):
            ws_in = wb_in.sheet_by_index(si)
            if ws_in.name.startswith('INV'):
                for i in range(min(15, ws_in.nrows)):
                    r = [str(ws_in.cell_value(i,j)).strip() for j in range(ws_in.ncols)]
                    full = ' '.join(r)
                    if 'SKDT/' in full and inv_no == 'Unknown Invoice':
                        for v in r:
                            if 'SKDT/' in v and '/DT:' in v:
                                inv_no = v.split('/DT:')[0].strip()
                    if r[0].startswith('M/S.') and 'SREE KANAGA' not in r[0] and buyer == 'Unknown Buyer':
                        buyer = r[0]
                    
        for si in range(wb_in.nsheets):
            ws_in = wb_in.sheet_by_index(si)
            if ws_in.name == 'PL': continue
            
            current_quality = "(1) 100% Organic Cotton (RM0104) (40s VL, / INTERLOCK )"
            
            for i in range(ws_in.nrows):
                col0 = str(ws_in.cell_value(i, 0)).strip()
                if 'ORGANIC COTTON' in col0.upper():
                    current_quality = col0
                    
                r = [ws_in.cell_value(i,j) for j in range(ws_in.ncols)]
                try:
                    style = str(r[1]).strip()
                    style = re.sub(r'\(SIZE:.*?\)', '', style, flags=re.IGNORECASE).strip()
                    qty = float(r[5])
                    try:
                        net_wt = float(r[9])
                    except:
                        net_wt = None
                    
                    if style and qty > 0:
                        if net_wt and net_wt > 0:
                            style_weights[style] = net_wt
                            
                        styles_data.append({
                            'style': style,
                            'qty': qty,
                            'inv_no': inv_no,
                            'buyer': buyer,
                            'quality': current_quality
                        })
                except:
                    pass
                    
    if not styles_data:
        raise ValueError("No valid style data found in the uploaded invoices.")
        
    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = 'Mass Balance Sheet'
    
    bold_font = Font(bold=True)
    header_font = Font(bold=True, size=10)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    
    ws_out.merge_cells('A1:W1')
    c = ws_out['A1']
    c.value = "Mass Balance Sheet"
    c.font = Font(bold=True, size=14)
    c.alignment = center_align
    
    ws_out.merge_cells('B3:I3'); ws_out['B3'] = "Production Capacity :"
    ws_out.merge_cells('L3:S3'); ws_out['L3'] = "Storage Capacity :"
    
    ws_out.merge_cells('B4:I4'); ws_out['B4'] = "Inword Data"
    ws_out.merge_cells('J4:U4'); ws_out['J4'] = "Outward Data"
    ws_out.merge_cells('V4:W4'); ws_out['V4'] = "Stock Details"
    
    headers = [
        "Sr.No ", "Suppliers Name ", "Product Name and Quality", "TC No(IDFL or Other CB)", 
        "Certified Weight(Kg)", "Net Wt (Kg.)", "Gross Weight (Kg.)", "Lot No/Batch No", 
        "Open Stock in Kgs.", "Raw Material used in Kg", "Product Name", "Loss(%)", 
        "Buyers Name", "Invoice No.", "Certified Weight(Kg)", "Net Wt(Kg)", "Gross Weight(Kg)", 
        "Supplementary Wt (Kg)", "Transport Details(BL No/Challan No)", "Standard", 
        "IDFL TC No.", "Raw Material (Kg)", " Finished Product (kg)"
    ]
    
    ws_out.append(headers)
    for col_idx in range(1, len(headers)+1):
        cell = ws_out.cell(row=5, column=col_idx)
        cell.font = header_font
        cell.border = border
        cell.alignment = center_align
        if 2 <= col_idx <= 9: cell.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        elif 10 <= col_idx <= 21: cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        elif 22 <= col_idx <= 23: cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        else: cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")

    for i in range(1, len(headers)+1):
        ws_out.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 15
    ws_out.column_dimensions['B'].width = 30
    ws_out.column_dimensions['C'].width = 40
    ws_out.column_dimensions['J'].width = 30
    ws_out.column_dimensions['K'].width = 35

    row_num = 6
    for idx, data in enumerate(styles_data, 1):
        style = data['style']
        qty = data['qty']
        buyer = data['buyer']
        inv_no = data['inv_no']
        quality = data.get('quality', "(1) 100% Organic Cotton (RM0104) (40s VL, / INTERLOCK )")
        
        net_wt = style_weights.get(style, None)
        
        if net_wt is not None:
            supp_wt = net_wt * 0.10
            cert_wt = net_wt - supp_wt
            finished_prod = net_wt * qty
            loss_pct = 0.15 
            raw_used = finished_prod * (1 + loss_pct)
            
            raw_val = round(raw_used, 3)
            cert_val = round(cert_wt, 3)
            net_val = round(net_wt, 3)
            supp_val = round(supp_wt, 3)
            fin_val = round(finished_prod, 3)
        else:
            raw_val = cert_val = net_val = supp_val = fin_val = ""
        
        row_values = [
            idx,
            "Sri Shanmugavel Mills Private Limited Knitting Division",
            quality,
            "", "", "", "", "", "",
            raw_val,
            style,
            "15.000%",
            buyer,
            inv_no,
            cert_val,
            net_val,
            net_val,
            supp_val,
            "S00055408" if idx == 1 else "",
            "GOTS" if idx == 1 else "",
            "551193" if idx == 1 else "",
            "",
            fin_val
        ]
        
        for col_idx, val in enumerate(row_values, 1):
            cell = ws_out.cell(row=row_num, column=col_idx)
            cell.value = val
            cell.border = border
            if type(val) in [int, float]:
                cell.alignment = Alignment(horizontal='right')
            else:
                cell.alignment = Alignment(horizontal='center', wrap_text=True)
                
        row_num += 1

    wb_out.save(output_path)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'invoice_file' not in request.files:
        flash('No file uploaded.')
        return redirect(request.url)
    
    files = request.files.getlist('invoice_file')
    if not files or files[0].filename == '':
        flash('No files selected.')
        return redirect(request.url)
        
    filepaths = []
    for file in files:
        if file and file.filename.endswith('.xls'):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            filepaths.append(filepath)
            
    if not filepaths:
        flash('Please upload valid .xls invoice files.')
        return redirect('/')
        
    output_filename = "MASS_BALANCE_COMBINED.xlsx"
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
    
    try:
        process_invoice_files(filepaths, output_path)
        return send_file(output_path, as_attachment=True)
    except Exception as e:
        flash(f"Error processing files: {str(e)}")
        return redirect('/')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
