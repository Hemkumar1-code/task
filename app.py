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
    style_weights = {}
    
    # Cache workbooks in memory to avoid reading from disk multiple times
    workbooks = []
    
    for invoice_path in invoice_paths:
        try:
            wb_in = xlrd.open_workbook(invoice_path, on_demand=True)
            workbooks.append(wb_in)
        except Exception as e:
            continue
            
        for si in range(wb_in.nsheets):
            ws_in = wb_in.sheet_by_index(si)
            if ws_in.name == 'PL': continue
            for i in range(ws_in.nrows):
                r = [ws_in.cell_value(i,j) for j in range(ws_in.ncols)]
                try:
                    style = str(r[1]).strip()
                    style = re.sub(r'\(SIZE:.*?\)', '', style, flags=re.IGNORECASE).strip()
                    qty = float(r[5])
                    net_wt = float(r[9])
                    if style and qty > 0 and net_wt > 0:
                        existing_wt = style_weights.get(style, 0)
                        style_weights[style] = max(existing_wt, net_wt)
                except:
                    pass

    styles_data = []
    
    pl_gross_weights = {}
    pl_net_weights = {}
    
    for wb_in in workbooks:
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
        
        # Extract PL weights for this invoice
        for si in range(wb_in.nsheets):
            ws_in = wb_in.sheet_by_index(si)
            if ws_in.name == 'PL':
                for i in range(ws_in.nrows):
                    row_str = ' '.join([str(x).strip() for x in ws_in.row_values(i)]).upper()
                    if 'TTL GROSS WEIGHT' in row_str:
                        for v in ws_in.row_values(i):
                            try:
                                val = float(v)
                                if val > 0: pl_gross_weights[inv_no] = val; break
                            except: pass
                    elif 'TTL NET WEIGHT' in row_str:
                        for v in ws_in.row_values(i):
                            try:
                                val = float(v)
                                if val > 0: pl_net_weights[inv_no] = val; break
                            except: pass

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
                    
                    if style and qty > 0:
                        styles_data.append({
                            'style': style,
                            'qty': qty,
                            'inv_no': inv_no,
                            'buyer': buyer,
                            'quality': current_quality
                        })
                except:
                    pass

    # Calculate total finished product for each invoice
    invoice_totals = {}
    for data in styles_data:
        inv_no_curr = data['inv_no']
        style_curr = data['style']
        qty_curr = data['qty']
        single_piece_wt = style_weights.get(style_curr, 0)
        if single_piece_wt:
            invoice_totals[inv_no_curr] = invoice_totals.get(inv_no_curr, 0) + (single_piece_wt * qty_curr)

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
        "IDFL TC No.", "Raw Material (Kg)", " Finished Product (kg)", "Difference"
    ]
    
    ws_out.append(headers)
    for col_idx in range(1, len(headers)+1):
        cell = ws_out.cell(row=5, column=col_idx)
        cell.font = header_font
        cell.border = border
        cell.alignment = center_align
        if 2 <= col_idx <= 9: cell.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        elif 10 <= col_idx <= 21: cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        elif 22 <= col_idx <= 24: cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
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
        
        single_piece_wt = style_weights.get(style, None)
        
        # Calculate ratio for this invoice
        ratio = 1.0
        inv_total_fin = invoice_totals.get(inv_no, 0)
        pl_net = pl_net_weights.get(inv_no, None)
        if pl_net is not None and inv_total_fin > 0:
            ratio = pl_net / inv_total_fin
        
        if single_piece_wt is not None:
            raw_total_net_wt = single_piece_wt * qty
            total_net_wt = raw_total_net_wt * ratio
            
            supp_wt = total_net_wt * 0.10
            cert_wt = total_net_wt - supp_wt
            finished_prod = total_net_wt
            loss_pct = 0.15 
            raw_used = finished_prod * (1 + loss_pct)
            
            raw_val = round(raw_used, 3)
            cert_val = round(cert_wt, 3)
            net_val = round(total_net_wt, 3)
            supp_val = round(supp_wt, 3)
            fin_val = round(finished_prod, 3)
        else:
            raw_val = cert_val = net_val = supp_val = fin_val = ""
            
        # Calculate Difference (PL Gross Weight - Invoice Total Finished Weight)
        diff_val = ""
        pl_gross = pl_gross_weights.get(inv_no, None)
        pl_net = pl_net_weights.get(inv_no, None)
        inv_total_fin = pl_net if pl_net is not None else invoice_totals.get(inv_no, None)
        
        if pl_gross is not None and inv_total_fin is not None:
            diff_val = round(pl_gross - inv_total_fin, 3)
        
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
            fin_val,
            diff_val
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
