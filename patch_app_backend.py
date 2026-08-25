import re

with open('e:/HEM/task/webapp/app.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Make sure we import database
if 'import database' not in code:
    code = code.replace('import openpyxl', 'import openpyxl\nimport database')

# --- 1. Update style_weights extraction to use DB ---
# Find style_weights initialization
code = code.replace('    style_weights = {}', '    style_weights = database.get_style_weights()')

# After extraction, save style_weights
code = code.replace('    styles_data = []', '    database.save_style_weights(style_weights)\n\n    styles_data = []')

# --- 2. Remove PL extraction completely because we don't need difference column anymore ---
pl_pattern = re.compile(r'# Extract PL weights for this invoice.*?elif \'TTL NET WEIGHT\' in row_str:.*?pass', re.DOTALL)
code = pl_pattern.sub('', code)

# --- 3. Replace Invoice Totals and Difference Column logic with IDFL Stock matching ---
invoice_totals_pattern = re.compile(r'# Calculate total finished product for each invoice.*?wb_out = openpyxl\.Workbook\(\)', re.DOTALL)

stock_logic = '''
    idfl_stock = database.get_idfl_stock()
    
    def find_matching_stock(quality, required_weight):
        # Determine target product string and sheet based on quality
        q = quality.upper()
        
        target_prod = None
        target_sheet = None
        
        if 'WOVEN' in q and '100%' in q:
            target_prod = 'Woven Fabrics' # Partial match
            target_sheet = 'IDFL'
        elif '95%' in q and '5%' in q:
            target_prod = '95%'
            target_sheet = None # Search all, or NON-IDFL
        elif '100% ORGANIC COTTON' in q:
            target_prod = '100% Organic Cotton'
            target_sheet = 'NON-IDFL'
            
        for s in idfl_stock:
            # Skip if exhausted or not enough weight
            if s.get('status') == 'Exhausted' or s.get('remaining_weight', 0) < required_weight:
                continue
                
            # Sheet filter
            if target_sheet and s.get('sheet') != target_sheet:
                continue
                
            prod_str = s.get('products', '').upper()
            if target_prod and target_prod.upper() in prod_str:
                return s
                
            # If no target prod identified but it's 100% organic cotton
            if not target_prod and '100% ORGANIC COTTON' in prod_str:
                return s
                
        return None

    wb_out = openpyxl.Workbook()
'''
code = invoice_totals_pattern.sub(stock_logic.strip() + '\n', code)

# --- 4. Update Headers ---
old_headers = '''
    headers = [
        "S.No", "Production Unit", "Product Name and Quality", "Standard", 
        "TC No.", "TC Date", "IDFL TC No.", "IDFL TC Date", "Invoice Number", 
        "Raw Material Used (Kg)", "Style / Item", "Finished Product (Kg)", 
        "Difference", "Supplementary Wt (Kg)", "Certified Weight (Kg)", "Net Wt (Kg)", 
        "Gross Wt (Kg)", "Buyer"
    ]
'''
new_headers = '''
    headers = [
        "S.No", "Production Unit", "Product Name and Quality", "Standard", 
        "TC No.", "TC Date", "TC No(IDFL or Other CB)", "IDFL TC Date", "Invoice Number", 
        "Raw Material Used (Kg)", "Style / Item", "Finished Product (Kg)", 
        "Supplementary Wt (Kg)", "Certified Weight (Kg)", "Net Wt (Kg)", 
        "Gross Wt (Kg)", "Transport Details(BL No/Challan No)", "Buyer"
    ]
'''
code = code.replace(old_headers.strip(), new_headers.strip())

# Update column colors and widths
old_colors = '''
        if 2 <= col_idx <= 9: cell.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        elif 10 <= col_idx <= 21: cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        elif 22 <= col_idx <= 24: cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        else: cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
'''
new_colors = '''
        if 2 <= col_idx <= 9: cell.fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        elif 10 <= col_idx <= 21: cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        else: cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
'''
code = code.replace(old_colors.strip(), new_colors.strip())


# --- 5. Update Row Calculation Logic ---
row_logic_pattern = re.compile(r'# Calculate ratio for this invoice.*?for col_idx, val in enumerate\(row_values, 1\):', re.DOTALL)

new_row_logic = '''
        if single_piece_wt is not None:
            finished_prod = single_piece_wt * qty
            
            # Loss percentage increased to 21%
            loss_pct = 0.21
            raw_used = finished_prod * (1 + loss_pct)
            
            supp_wt = finished_prod * 0.10
            cert_wt = finished_prod - supp_wt
            
            raw_val = round(raw_used, 3)
            cert_val = round(cert_wt, 3)
            supp_val = round(supp_wt, 3)
            fin_val = round(finished_prod, 3)
            
            # FIFO Stock Logic
            matched_stock = find_matching_stock(quality, finished_prod)
            
            if matched_stock:
                net_val = round(matched_stock['remaining_weight'], 3)
                tc_number = matched_stock['tc_number']
                
                # Deduct weight
                matched_stock['remaining_weight'] -= finished_prod
            else:
                net_val = round(finished_prod, 3) # Fallback
                tc_number = ""
                
        else:
            raw_val = cert_val = net_val = supp_val = fin_val = ""
            tc_number = ""
            
        buyer = "M/S. DUNS"
        standard = "GOTS"
        
        row_values = [
            idx,
            "Sri Shanmugavel Mills Private Limited Knitting Division",
            quality,
            standard, 
            "", 
            "", 
            tc_number, # TC No(IDFL or Other CB)
            "", 
            inv_no,
            raw_val,
            style,
            fin_val,
            supp_val,
            cert_val,
            net_val, # Net Wt
            net_val, # Gross Wt
            "", # Transport Details
            buyer
        ]
        
        for col_idx, val in enumerate(row_values, 1):
'''

code = row_logic_pattern.sub(new_row_logic.strip() + '\n', code)

# Save IDFL stock at the end of the function
save_pattern = re.compile(r'    wb_out\.save\(output_path\)')
code = save_pattern.sub('    database.save_idfl_stock(idfl_stock)\n    wb_out.save(output_path)', code)

with open('e:/HEM/task/webapp/app.py', 'w', encoding='utf-8') as f:
    f.write(code)
