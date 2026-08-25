import re

with open('e:/HEM/task/webapp/app.py', 'r', encoding='utf-8') as f:
    code = f.read()

optimized = '''def process_invoice_files(invoice_paths, output_path):
    style_weights = {}
    
    # Cache workbooks in memory to avoid reading from disk multiple times
    workbooks = []
    
    for invoice_path in invoice_paths:
        try:
            wb_in = xlrd.open_workbook(invoice_path)
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
            invoice_totals[inv_no_curr] = invoice_totals.get(inv_no_curr, 0) + (single_piece_wt * qty_curr)'''

old_func_pattern = re.compile(r'def process_invoice_files\(invoice_paths, output_path\):.*?invoice_totals\[inv_no_curr\] \= invoice_totals\.get\(inv_no_curr\, 0\) \+ \(single_piece_wt \* qty_curr\)', re.DOTALL)

code = old_func_pattern.sub(optimized, code)

with open('e:/HEM/task/webapp/app.py', 'w', encoding='utf-8') as f:
    f.write(code)
