import json
import os
import tempfile

# Define the base directory for DB files
DB_DIR = os.path.join(tempfile.gettempdir(), 'app_data')
if not os.path.exists(DB_DIR):
    os.makedirs(DB_DIR, exist_ok=True)

STYLE_WEIGHTS_FILE = os.path.join(DB_DIR, 'style_weights.json')
IDFL_STOCK_FILE = os.path.join(DB_DIR, 'idfl_stock.json')

def load_json(filepath, default):
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return default
    return default

def save_json(filepath, data):
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4)

def get_style_weights():
    return load_json(STYLE_WEIGHTS_FILE, {})

def save_style_weights(weights):
    save_json(STYLE_WEIGHTS_FILE, weights)

def get_idfl_stock():
    # List of stock objects: {"id": "row_index", "tc_number": "", "products": "", "remaining_weight": float, "sheet": "NON-IDFL"}
    return load_json(IDFL_STOCK_FILE, [])

def save_idfl_stock(stock):
    save_json(IDFL_STOCK_FILE, stock)
