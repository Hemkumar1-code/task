with open('e:/HEM/task/webapp/app.py', 'r', encoding='utf-8') as f:
    code = f.read()

if 'jsonify' not in code:
    code = code.replace('from flask import Flask, request', 'from flask import Flask, request, jsonify')

api_routes = '''
@app.route('/api/weights', methods=['GET', 'POST'])
def handle_weights():
    if request.method == 'POST':
        data = request.json
        database.save_style_weights(data)
        return jsonify({'status': 'success'})
    return jsonify(database.get_style_weights())

@app.route('/api/idfl-stock', methods=['GET', 'POST'])
def handle_idfl_stock():
    if request.method == 'POST':
        data = request.json
        database.save_idfl_stock(data)
        return jsonify({'status': 'success'})
    return jsonify(database.get_idfl_stock())
'''

code = code.replace('if __name__ == "__main__":', api_routes + '\nif __name__ == "__main__":')
code = code.replace("if __name__ == '__main__':", api_routes + "\nif __name__ == '__main__':")

with open('e:/HEM/task/webapp/app.py', 'w', encoding='utf-8') as f:
    f.write(code)
