html_content = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Mass Balance Tracker & Database</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=Inter:wght@400;500;600&display=swap" rel="stylesheet">
    <style>
        body {
            background: #0f172a;
            background-image: 
                radial-gradient(at 0% 0%, hsla(253,16%,7%,1) 0, transparent 50%), 
                radial-gradient(at 50% 0%, hsla(225,39%,30%,1) 0, transparent 50%), 
                radial-gradient(at 100% 0%, hsla(339,49%,30%,1) 0, transparent 50%);
            background-attachment: fixed;
            font-family: 'Inter', sans-serif;
            color: #f8fafc;
            min-height: 100vh;
        }
        h1, h2, h3, .nav-item {
            font-family: 'Outfit', sans-serif;
        }
        .glass-panel {
            background: rgba(30, 41, 59, 0.7);
            backdrop-filter: blur(16px);
            -webkit-backdrop-filter: blur(16px);
            border: 1px solid rgba(255, 255, 255, 0.1);
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
        }
        .upload-area {
            border: 2px dashed rgba(148, 163, 184, 0.5);
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .upload-area:hover, .upload-area.dragover {
            border-color: #3b82f6;
            background-color: rgba(59, 130, 246, 0.1);
            transform: translateY(-2px);
        }
        .nav-item {
            position: relative;
            transition: all 0.3s;
        }
        .nav-item.active {
            color: #60a5fa;
            font-weight: 600;
        }
        .nav-item.active::after {
            content: '';
            position: absolute;
            bottom: -8px;
            left: 0;
            width: 100%;
            height: 3px;
            background: #3b82f6;
            border-radius: 4px;
            box-shadow: 0 0 10px rgba(59, 130, 246, 0.5);
        }
        .table-container::-webkit-scrollbar {
            height: 8px;
            width: 8px;
        }
        .table-container::-webkit-scrollbar-track {
            background: rgba(0,0,0,0.1);
            border-radius: 4px;
        }
        .table-container::-webkit-scrollbar-thumb {
            background: rgba(255,255,255,0.2);
            border-radius: 4px;
        }
        .table-container::-webkit-scrollbar-thumb:hover {
            background: rgba(255,255,255,0.3);
        }
    </style>
</head>
<body class="p-4 md:p-8">

    <div class="max-w-6xl mx-auto">
        <!-- Header & Nav -->
        <header class="flex flex-col md:flex-row justify-between items-center mb-8 glass-panel p-6 rounded-2xl">
            <div class="flex items-center gap-4 mb-4 md:mb-0">
                <div class="w-12 h-12 rounded-xl bg-gradient-to-br from-blue-500 to-purple-600 flex items-center justify-center shadow-lg shadow-blue-500/30">
                    <svg class="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 002-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10"></path></svg>
                </div>
                <div>
                    <h1 class="text-2xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-blue-400 to-purple-400">Mass Balance Hub</h1>
                    <p class="text-sm text-slate-400">Automated TC Stock & Document Generation</p>
                </div>
            </div>
            
            <nav class="flex gap-8 text-slate-300">
                <button onclick="switchTab('generator')" id="nav-generator" class="nav-item active hover:text-white">Generator</button>
                <button onclick="switchTab('idfl')" id="nav-idfl" class="nav-item hover:text-white">IDFL Stock DB</button>
                <button onclick="switchTab('weights')" id="nav-weights" class="nav-item hover:text-white">Style Weights DB</button>
            </nav>
        </header>

        <!-- Generator Tab -->
        <div id="tab-generator" class="tab-content transition-opacity duration-300">
            <div class="glass-panel rounded-2xl p-8 max-w-2xl mx-auto">
                <div class="text-center mb-8">
                    <h2 class="text-2xl font-semibold mb-2">Generate Mass Balance Sheet</h2>
                    <p class="text-slate-400">Upload your Invoice (.xls) files. The system will automatically deduct stock from the IDFL database.</p>
                </div>

                {% with messages = get_flashed_messages() %}
                  {% if messages %}
                    <div class="bg-red-500/10 border border-red-500/50 text-red-400 px-4 py-3 rounded-xl mb-6">
                      {% for message in messages %}
                        <span class="block sm:inline">{{ message }}</span>
                      {% endfor %}
                    </div>
                  {% endif %}
                {% endwith %}

                <form action="/upload" method="post" enctype="multipart/form-data" id="uploadForm">
                    <div class="upload-area rounded-2xl p-12 flex flex-col items-center justify-center cursor-pointer mb-6" id="dropZone" onclick="document.getElementById('fileInput').click()">
                        <div class="w-16 h-16 rounded-full bg-blue-500/10 flex items-center justify-center mb-4">
                            <svg class="w-8 h-8 text-blue-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path>
                            </svg>
                        </div>
                        <p class="text-lg font-medium mb-1">Click or drag invoices here</p>
                        <p class="text-sm text-slate-400">Supports multiple .xls files</p>
                        <input type="file" name="invoice_file" id="fileInput" class="hidden" accept=".xls" multiple onchange="fileSelected()">
                    </div>
                    
                    <div id="fileInfo" class="hidden bg-slate-800/50 border border-slate-700 rounded-xl p-4 mb-6 flex justify-between items-center">
                        <div class="flex items-center text-slate-300">
                            <svg class="w-6 h-6 text-blue-400 mr-3" fill="currentColor" viewBox="0 0 20 20"><path fill-rule="evenodd" d="M4 4a2 2 0 012-2h4.586A2 2 0 0112 2.586L15.414 6A2 2 0 0116 7.414V16a2 2 0 01-2 2H6a2 2 0 01-2-2V4zm2 6a1 1 0 011-1h6a1 1 0 110 2H7a1 1 0 01-1-1zm1 3a1 1 0 100 2h6a1 1 0 100-2H7z" clip-rule="evenodd"></path></svg>
                            <span id="fileName" class="font-medium"></span>
                        </div>
                    </div>

                    <button type="submit" id="submitBtn" class="w-full bg-gradient-to-r from-blue-600 to-blue-500 hover:from-blue-500 hover:to-blue-400 text-white font-bold py-4 px-6 rounded-xl shadow-lg shadow-blue-500/25 transition-all transform hover:-translate-y-1 opacity-50 cursor-not-allowed" disabled>
                        Process & Generate Sheet
                    </button>
                </form>
            </div>
        </div>

        <!-- IDFL Stock Tab -->
        <div id="tab-idfl" class="tab-content hidden transition-opacity duration-300">
            <div class="glass-panel rounded-2xl p-8">
                <div class="flex justify-between items-center mb-6">
                    <div>
                        <h2 class="text-2xl font-semibold mb-1">IDFL TC Stock Tracker</h2>
                        <p class="text-slate-400 text-sm">Live remaining balances of your certified stock</p>
                    </div>
                    <button onclick="loadIdflStock()" class="p-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors">
                        <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"></path></svg>
                    </button>
                </div>
                
                <div class="table-container overflow-x-auto rounded-xl border border-slate-700">
                    <table class="w-full text-left border-collapse">
                        <thead>
                            <tr class="bg-slate-800/80 text-slate-300 text-sm">
                                <th class="p-4 font-medium border-b border-slate-700">Sheet</th>
                                <th class="p-4 font-medium border-b border-slate-700">TC Number</th>
                                <th class="p-4 font-medium border-b border-slate-700">Products</th>
                                <th class="p-4 font-medium border-b border-slate-700 text-right">Remaining (Kg)</th>
                                <th class="p-4 font-medium border-b border-slate-700">Status</th>
                            </tr>
                        </thead>
                        <tbody id="idflTableBody" class="text-sm divide-y divide-slate-800">
                            <!-- Populated via JS -->
                            <tr><td colspan="5" class="p-8 text-center text-slate-500">Loading stock data...</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Style Weights Tab -->
        <div id="tab-weights" class="tab-content hidden transition-opacity duration-300">
            <div class="glass-panel rounded-2xl p-8">
                <div class="flex justify-between items-center mb-6">
                    <div>
                        <h2 class="text-2xl font-semibold mb-1">Style Weight Database</h2>
                        <p class="text-slate-400 text-sm">Single piece weights extracted from invoices</p>
                    </div>
                </div>
                
                <div class="table-container overflow-x-auto rounded-xl border border-slate-700">
                    <table class="w-full text-left border-collapse">
                        <thead>
                            <tr class="bg-slate-800/80 text-slate-300 text-sm">
                                <th class="p-4 font-medium border-b border-slate-700 w-2/3">Style / Item Name</th>
                                <th class="p-4 font-medium border-b border-slate-700">Single Piece Weight (Kg)</th>
                                <th class="p-4 font-medium border-b border-slate-700 text-right">Actions</th>
                            </tr>
                        </thead>
                        <tbody id="weightsTableBody" class="text-sm divide-y divide-slate-800">
                            <!-- Populated via JS -->
                            <tr><td colspan="3" class="p-8 text-center text-slate-500">Loading style weights...</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <!-- Edit Weight Modal -->
    <div id="editModal" class="fixed inset-0 bg-black/60 backdrop-blur-sm hidden items-center justify-center z-50">
        <div class="glass-panel rounded-2xl p-6 w-full max-w-md mx-4">
            <h3 class="text-xl font-semibold mb-4" id="modalTitle">Edit Weight</h3>
            <div class="mb-4">
                <label class="block text-sm text-slate-400 mb-1">Style</label>
                <input type="text" id="editStyle" class="w-full bg-slate-800 border border-slate-700 rounded-lg px-4 py-2 text-white outline-none focus:border-blue-500" readonly>
            </div>
            <div class="mb-6">
                <label class="block text-sm text-slate-400 mb-1">Weight (Kg)</label>
                <input type="number" step="0.001" id="editWeight" class="w-full bg-slate-800 border border-slate-700 rounded-lg px-4 py-2 text-white outline-none focus:border-blue-500">
            </div>
            <div class="flex justify-end gap-3">
                <button onclick="closeModal()" class="px-4 py-2 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 transition-colors">Cancel</button>
                <button onclick="saveWeight()" class="px-4 py-2 rounded-lg bg-blue-600 hover:bg-blue-500 text-white shadow-lg shadow-blue-500/25 transition-colors">Save</button>
            </div>
        </div>
    </div>

    <script>
        // --- Tab Switching Logic ---
        function switchTab(tabId) {
            document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
            document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
            
            document.getElementById('tab-' + tabId).classList.remove('hidden');
            document.getElementById('nav-' + tabId).classList.add('active');
            
            if (tabId === 'idfl') loadIdflStock();
            if (tabId === 'weights') loadStyleWeights();
        }

        // --- File Upload UI Logic ---
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const submitBtn = document.getElementById('submitBtn');
        const fileInfo = document.getElementById('fileInfo');
        const fileNameDisplay = document.getElementById('fileName');

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, e => {
                e.preventDefault(); e.stopPropagation();
            });
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'));
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'));
        });

        dropZone.addEventListener('drop', (e) => {
            if (e.dataTransfer.files.length > 0) {
                fileInput.files = e.dataTransfer.files;
                fileSelected();
            }
        });

        function fileSelected() {
            if (fileInput.files.length > 0) {
                fileNameDisplay.textContent = fileInput.files.length === 1 
                    ? fileInput.files[0].name 
                    : fileInput.files.length + " files selected";
                fileInfo.classList.remove('hidden');
                submitBtn.disabled = false;
                submitBtn.classList.remove('opacity-50', 'cursor-not-allowed');
            }
        }

        // --- API Integration Logic ---
        let currentWeights = {};

        async function loadIdflStock() {
            try {
                const res = await fetch('/api/idfl-stock');
                const data = await res.json();
                const tbody = document.getElementById('idflTableBody');
                
                if (data.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="5" class="p-8 text-center text-slate-500">No stock data available. Upload the IDFL Excel file to initialize.</td></tr>';
                    return;
                }
                
                tbody.innerHTML = data.map(item => `
                    <tr class="hover:bg-slate-800/50 transition-colors">
                        <td class="p-4"><span class="px-2 py-1 rounded bg-slate-800 text-xs text-slate-300 border border-slate-700">${item.sheet}</span></td>
                        <td class="p-4 font-medium text-slate-200">${item.tc_number}</td>
                        <td class="p-4 text-slate-400 truncate max-w-xs" title="${item.products}">${item.products}</td>
                        <td class="p-4 text-right font-mono ${item.remaining_weight > 0 ? 'text-green-400' : 'text-red-400'}">${item.remaining_weight.toFixed(3)}</td>
                        <td class="p-4">
                            <span class="px-2 py-1 rounded-full text-xs ${item.status === 'Exhausted' || item.remaining_weight <= 0 ? 'bg-red-500/10 text-red-400' : 'bg-green-500/10 text-green-400'}">
                                ${item.remaining_weight <= 0 ? 'Exhausted' : item.status}
                            </span>
                        </td>
                    </tr>
                `).join('');
            } catch (e) {
                console.error(e);
            }
        }

        async function loadStyleWeights() {
            try {
                const res = await fetch('/api/weights');
                currentWeights = await res.json();
                const tbody = document.getElementById('weightsTableBody');
                
                const entries = Object.entries(currentWeights);
                if (entries.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="3" class="p-8 text-center text-slate-500">No styles found. Upload invoices to build the database.</td></tr>';
                    return;
                }
                
                tbody.innerHTML = entries.map(([style, weight]) => `
                    <tr class="hover:bg-slate-800/50 transition-colors group">
                        <td class="p-4 text-slate-300">${style}</td>
                        <td class="p-4 font-mono text-blue-400">${weight.toFixed(3)}</td>
                        <td class="p-4 text-right">
                            <button onclick="openModal('${style.replace(/'/g, "\\'")}', ${weight})" class="opacity-0 group-hover:opacity-100 px-3 py-1 rounded bg-slate-700 hover:bg-slate-600 text-xs transition-all">Edit</button>
                        </td>
                    </tr>
                `).join('');
            } catch (e) {
                console.error(e);
            }
        }

        // --- Modal Logic ---
        function openModal(style, weight) {
            document.getElementById('editStyle').value = style;
            document.getElementById('editWeight').value = weight;
            document.getElementById('editModal').classList.remove('hidden');
            document.getElementById('editModal').classList.add('flex');
        }

        function closeModal() {
            document.getElementById('editModal').classList.add('hidden');
            document.getElementById('editModal').classList.remove('flex');
        }

        async function saveWeight() {
            const style = document.getElementById('editStyle').value;
            const weight = parseFloat(document.getElementById('editWeight').value);
            
            if (style && !isNaN(weight)) {
                currentWeights[style] = weight;
                try {
                    await fetch('/api/weights', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(currentWeights)
                    });
                    closeModal();
                    loadStyleWeights();
                } catch (e) {
                    alert('Error saving weight');
                }
            }
        }
    </script>
</body>
</html>
"""

with open('e:/HEM/task/webapp/templates/index.html', 'w', encoding='utf-8') as f:
    f.write(html_content)
