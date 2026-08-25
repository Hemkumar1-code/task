with open('e:/HEM/task/webapp/templates/index.html', 'r', encoding='utf-8') as f:
    html = f.read()

# 1. Make Edit button always visible
html = html.replace('class="opacity-0 group-hover:opacity-100 px-3 py-1 rounded bg-slate-700', 'class="px-3 py-1 rounded bg-blue-600 hover:bg-blue-500 text-white shadow-lg')

# 2. Add Download Button to IDFL Tab
idfl_refresh_btn = '''<button onclick="loadIdflStock()" class="p-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors">'''
idfl_download_btn = '''<button onclick="downloadCSV('idfl')" class="px-4 py-2 rounded-lg bg-green-600 hover:bg-green-500 text-white text-sm transition-colors shadow-lg shadow-green-500/25 flex items-center gap-2">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
                            Download CSV
                        </button>
                        <button onclick="loadIdflStock()" class="p-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors">'''
html = html.replace(idfl_refresh_btn, idfl_download_btn)

# 3. Add Download Button to Weights Tab
weights_header = '''<p class="text-slate-400 text-sm">Single piece weights extracted from invoices</p>
                    </div>
                </div>'''
weights_header_new = '''<p class="text-slate-400 text-sm">Single piece weights extracted from invoices</p>
                    </div>
                    <div class="flex gap-3">
                        <button onclick="downloadCSV('weights')" class="px-4 py-2 rounded-lg bg-green-600 hover:bg-green-500 text-white text-sm transition-colors shadow-lg shadow-green-500/25 flex items-center gap-2">
                            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
                            Download CSV
                        </button>
                    </div>
                </div>'''
html = html.replace(weights_header, weights_header_new)

# 4. Add CSV Download JS Logic
csv_logic = '''
        // --- CSV Download Logic ---
        function downloadCSV(type) {
            let csvContent = "data:text/csv;charset=utf-8,";
            
            if (type === 'idfl') {
                csvContent += "Sheet,TC Number,Products,Remaining (Kg),Status\\n";
                document.querySelectorAll('#idflTableBody tr').forEach(row => {
                    const cols = row.querySelectorAll('td');
                    if (cols.length === 5) {
                        const sheet = cols[0].innerText.trim();
                        const tc = cols[1].innerText.trim();
                        const products = '"' + cols[2].innerText.trim().replace(/"/g, '""') + '"';
                        const remaining = cols[3].innerText.trim();
                        const status = cols[4].innerText.trim();
                        csvContent += `${sheet},${tc},${products},${remaining},${status}\\n`;
                    }
                });
            } else if (type === 'weights') {
                csvContent += "Style / Item Name,Single Piece Weight (Kg)\\n";
                document.querySelectorAll('#weightsTableBody tr').forEach(row => {
                    const cols = row.querySelectorAll('td');
                    if (cols.length === 3) {
                        const style = '"' + cols[0].innerText.trim().replace(/"/g, '""') + '"';
                        const weight = cols[1].innerText.trim();
                        csvContent += `${style},${weight}\\n`;
                    }
                });
            }
            
            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", `${type}_export_${new Date().toISOString().split('T')[0]}.csv`);
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
'''

html = html.replace('// --- Modal Logic ---', csv_logic + '\n        // --- Modal Logic ---')

with open('e:/HEM/task/webapp/templates/index.html', 'w', encoding='utf-8') as f:
    f.write(html)
