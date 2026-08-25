with open('e:/HEM/task/webapp/templates/index.html', 'r', encoding='utf-8') as f:
    html = f.read()

old_idfl_header = '''<h2 class="text-2xl font-semibold mb-1">IDFL TC Stock Tracker</h2>
                        <p class="text-slate-400 text-sm">Live remaining balances of your certified stock</p>'''

new_idfl_header = '''<h2 class="text-2xl font-semibold mb-1">IDFL TC Stock Tracker</h2>
                        <p class="text-slate-400 text-sm">Live remaining balances of your certified stock</p>
                    </div>
                    <div class="flex gap-3">
                        <label for="idflUpload" class="cursor-pointer px-4 py-2 rounded-lg bg-blue-600 hover:bg-blue-500 text-white text-sm transition-colors shadow-lg shadow-blue-500/25 flex items-center gap-2">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"></path></svg>
                            Upload Original IDFL Stock
                        </label>
                        <input type="file" id="idflUpload" accept=".xlsx,.xls" class="hidden" onchange="uploadIdflStock(event)">'''

html = html.replace(old_idfl_header, new_idfl_header)

js_func = '''
        async function uploadIdflStock(event) {
            const file = event.target.files[0];
            if (!file) return;
            
            const formData = new FormData();
            formData.append('idfl_file', file);
            
            try {
                alert('Uploading and resetting IDFL Database. Please wait...');
                
                const res = await fetch('/api/upload-idfl', {
                    method: 'POST',
                    body: formData
                });
                
                if (res.ok) {
                    alert('IDFL Stock Database successfully reset with original data!');
                    loadIdflStock();
                } else {
                    alert('Error uploading IDFL Stock file.');
                }
            } catch (e) {
                console.error(e);
                alert('Network error while uploading.');
            }
            event.target.value = '';
        }
'''

html = html.replace('// --- API Integration Logic ---', '// --- API Integration Logic ---\n' + js_func)

with open('e:/HEM/task/webapp/templates/index.html', 'w', encoding='utf-8') as f:
    f.write(html)
