<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Papan Pengumuman - Pembangunan MZAK Hari Ini</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- SheetJS untuk Export Excel -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        /* Tema Futuristik Sci-Fi HUD */
        body {
            background-color: #030712;
            background-image: 
                radial-gradient(circle at 50% 50%, rgba(6, 182, 212, 0.05) 0%, transparent 80%),
                linear-gradient(to right, rgba(15, 23, 42, 0.4) 1px, transparent 1px),
                linear-gradient(to bottom, rgba(15, 23, 42, 0.4) 1px, transparent 1px);
            background-size: 100% 100%, 20px 20px, 20px 20px;
            color: #94a3b8;
            font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            font-size: 13px;
        }
        .hud-card {
            background: rgba(13, 20, 36, 0.75);
            backdrop-filter: blur(12px);
            border: 1px solid rgba(6, 182, 212, 0.3);
            box-shadow: 0 0 15px rgba(6, 182, 212, 0.1), inset 0 0 15px rgba(6, 182, 212, 0.03);
            border-radius: 8px;
        }
        .hud-header {
            background: linear-gradient(90deg, rgba(13, 20, 36, 0.9) 0%, rgba(6, 182, 212, 0.15) 100%);
            border-bottom: 1px solid rgba(6, 182, 212, 0.4);
        }
        
        /* Efek Teks Hologram Lembut, Tipis, dan Pelan (Khusus Judul Utama) */
        .hologram-title {
            position: relative;
            overflow: hidden;
            background: linear-gradient(90deg, #d946ef, #06b6d4, #facc15, #d946ef);
            background-size: 300% auto;
            color: transparent;
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            animation: hologramScroll 8s linear infinite;
            font-weight: 600;
            letter-spacing: 0.05em;
        }
        @keyframes hologramScroll {
            0% { background-position: 0% 50%; }
            100% { background-position: 300% 50%; }
        }

        /* Gaya Judul Tabel Standar yang Nyaman di Mata */
        .table-title {
            color: #38bdf8;
            font-weight: 600;
            letter-spacing: 0.025em;
        }

        .hud-input {
            background: rgba(3, 7, 18, 0.6) !important;
            border: 1px solid rgba(51, 65, 85, 0.8) !important;
            color: #e2e8f0 !important;
            border-radius: 4px;
            padding: 4px 8px;
            transition: all 0.2s ease;
        }
        .hud-input:focus {
            border-color: #06b6d4 !important;
            box-shadow: 0 0 8px rgba(6, 182, 212, 0.4);
            outline: none;
        }
        .hud-btn {
            background: rgba(6, 182, 212, 0.15);
            border: 1px solid #06b6d4;
            color: #22d3ee;
            transition: all 0.2s;
        }
        .hud-btn:hover {
            background: #06b6d4;
            color: #030712;
            box-shadow: 0 0 10px rgba(6, 182, 212, 0.6);
        }
        /* Custom Scrollbar */
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #030712; }
        ::-webkit-scrollbar-thumb { background: #1e293b; border-radius: 3px; border: 1px solid rgba(6,182,212,0.2); }
        ::-webkit-scrollbar-thumb:hover { background: #06b6d4; }
    </style>
</head>
<body class="min-h-screen p-4 flex flex-col gap-4">

    <!-- Header / Navigasi HUD -->
    <header class="hud-card hud-header p-4 flex flex-col justify-center items-center text-center gap-1 sticky top-4 z-50">
        <div>
            <p class="text-xs text-cyan-500 uppercase tracking-[0.2em]">Papan Pengumuman</p>
            <h1 class="text-2xl hologram-title tracking-tight">Pembangunan MZAK Hari Ini</h1>
            <p id="datetime" class="text-[10px] text-cyan-600 font-mono mt-1"></p>
        </div>
    </header>

    <!-- Main Content Grid -->
    <main class="grid grid-cols-1 md:grid-cols-2 gap-4 flex-grow">

        <!-- 1. Jumlah orang di lokasi MZAK Hari Ini -->
        <section class="hud-card p-4 flex flex-col">
            <div class="flex justify-between items-center mb-3 pb-2 border-b border-cyan-950">
                <h2 class="uppercase text-xs tracking-wider table-title text-sm">
                    1. Jumlah orang di lokasi MZAK Hari Ini
                </h2>
                <button onclick="addRow('table-orang')" class="hud-btn w-6 h-6 rounded flex items-center justify-center font-bold text-sm">+</button>
            </div>
            <div class="overflow-x-auto flex-grow flex flex-col justify-between">
                <table id="table-orang" class="w-full text-left border-collapse">
                    <thead>
                        <tr class="text-cyan-600 text-[11px] border-b border-cyan-950/60 uppercase">
                            <th class="p-1.5 w-10 text-center">No</th>
                            <th class="p-1.5">Unsur</th>
                            <th class="p-1.5 w-16 text-center">Pagi</th>
                            <th class="p-1.5 w-16 text-center">Siang</th>
                            <th class="p-1.5 w-16 text-center">Sore</th>
                            <th class="p-1.5 w-8 text-center">#</th>
                        </tr>
                    </thead>
                    <tbody class="text-xs">
                        <tr class="border-b border-slate-900/50 hover:bg-cyan-950/20">
                            <td class="p-1.5 text-center font-mono text-cyan-600">1</td>
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="Pekerja Utama" oninput="calculateTotals(); saveAllData();"></td>
                            <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="0" oninput="calculateTotals(); saveAllData();"></td>
                            <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="0" oninput="calculateTotals(); saveAllData();"></td>
                            <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="0" oninput="calculateTotals(); saveAllData();"></td>
                            <td class="p-1.5 text-center"><button onclick="deleteRow(this)" class="text-red-400 hover:text-red-300 font-bold">×</button></td>
                        </tr>
                    </tbody>
                </table>

                <!-- Baris Total -->
                <div class="w-full pt-3 mt-2 border-t border-orange-500/30 text-xs font-bold text-orange-400">
                    <table class="w-full border-collapse">
                        <tr>
                            <td class="w-10"></td>
                            <td class="text-right pr-3 uppercase tracking-wider">Total :</td>
                            <td id="total-pagi" class="w-16 text-center font-mono">0</td>
                            <td id="total-siang" class="w-16 text-center font-mono">0</td>
                            <td id="total-sore" class="w-16 text-center font-mono">0</td>
                            <td class="w-8"></td>
                        </tr>
                    </table>
                </div>
            </div>
        </section>

        <!-- 2. Hubungi Panitia di Lokasi MZAK -->
        <section class="hud-card p-4 flex flex-col">
            <div class="flex justify-between items-center mb-3 pb-2 border-b border-cyan-950">
                <h2 class="uppercase text-xs tracking-wider table-title text-sm">
                    2. Hubungi Panitia di Lokasi MZAK
                </h2>
                <button onclick="addRow('table-panitia')" class="hud-btn w-6 h-6 rounded flex items-center justify-center font-bold text-sm">+</button>
            </div>
            <div class="overflow-x-auto flex-grow">
                <table id="table-panitia" class="w-full text-left border-collapse">
                    <thead>
                        <tr class="text-cyan-600 text-[11px] border-b border-cyan-950/60 uppercase">
                            <th class="p-1.5 w-10 text-center">No</th>
                            <th class="p-1.5">Bidang</th>
                            <th class="p-1.5">PIC</th>
                            <th class="p-1.5 w-8 text-center">#</th>
                        </tr>
                    </thead>
                    <tbody class="text-xs">
                        <tr class="border-b border-slate-900/50 hover:bg-cyan-950/20">
                            <td class="p-1.5 text-center font-mono text-cyan-600">1</td>
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="Keamanan & Logistik" oninput="saveAllData()"></td>
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="Ahmad" oninput="saveAllData()"></td>
                            <td class="p-1.5 text-center"><button onclick="deleteRow(this)" class="text-red-400 hover:text-red-300 font-bold">×</button></td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </section>

        <!-- 3. Daftar Pekerjaan hari ini -->
        <section class="hud-card p-4 flex flex-col">
            <div class="flex justify-between items-center mb-3 pb-2 border-b border-cyan-950">
                <h2 class="uppercase text-xs tracking-wider table-title text-sm">
                    3. Daftar Pekerjaan Hari Ini
                </h2>
                <button onclick="addRow('table-pekerjaan')" class="hud-btn w-6 h-6 rounded flex items-center justify-center font-bold text-sm">+</button>
            </div>
            <div class="overflow-x-auto flex-grow">
                <table id="table-pekerjaan" class="w-full text-left border-collapse">
                    <thead>
                        <tr class="text-cyan-600 text-[11px] border-b border-cyan-950/60 uppercase">
                            <th class="p-1.5 w-10 text-center">No</th>
                            <th class="p-1.5">Item Pekerjaan</th>
                            <th class="p-1.5">Nama Tukang</th>
                            <th class="p-1.5 w-20 text-center">Jml Helper</th>
                            <th class="p-1.5 w-8 text-center">#</th>
                        </tr>
                    </thead>
                    <tbody class="text-xs">
                        <tr class="border-b border-slate-900/50 hover:bg-cyan-950/20">
                            <td class="p-1.5 text-center font-mono text-cyan-600">1</td>
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="Pemasangan Keramik" oninput="saveAllData()"></td>
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="Budi" oninput="saveAllData()"></td>
                            <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="2" oninput="saveAllData()"></td>
                            <td class="p-1.5 text-center"><button onclick="deleteRow(this)" class="text-red-400 hover:text-red-300 font-bold">×</button></td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </section>

        <!-- 4. Tabel Informasi Input Manual -->
        <section class="hud-card p-4 flex flex-col">
            <div class="flex justify-between items-center mb-3 pb-2 border-b border-cyan-950">
                <h2 class="uppercase text-xs tracking-wider table-title text-sm">
                    4. Informasi Input Manual
                </h2>
                <button onclick="addRow('table-informasi')" class="hud-btn w-6 h-6 rounded flex items-center justify-center font-bold text-sm">+</button>
            </div>
            <div class="overflow-x-auto flex-grow">
                <table id="table-informasi" class="w-full text-left border-collapse">
                    <thead>
                        <tr class="text-cyan-600 text-[11px] border-b border-cyan-950/60 uppercase">
                            <th class="p-1.5 w-10 text-center">No</th>
                            <th class="p-1.5">Keterangan / Informasi</th>
                            <th class="p-1.5">Catatan</th>
                            <th class="p-1.5 w-8 text-center">#</th>
                        </tr>
                    </thead>
                    <tbody class="text-xs">
                        <tr class="border-b border-slate-900/50 hover:bg-cyan-950/20">
                            <td class="p-1.5 text-center font-mono text-cyan-600">1</td>
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="Status Proyek" oninput="saveAllData()"></td>
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="Berjalan lancar" oninput="saveAllData()"></td>
                            <td class="p-1.5 text-center"><button onclick="deleteRow(this)" class="text-red-400 hover:text-red-300 font-bold">×</button></td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </section>

    </main>

    <!-- Tombol Aksi di Paling Bawah Halaman -->
    <div class="hud-card p-4 flex flex-wrap justify-center items-center gap-3">
        <button onclick="shareEditor()" class="hud-btn px-4 py-2 rounded text-xs font-bold uppercase tracking-wider flex items-center gap-2">
            <span>🔗</span> Share Editor <span class="text-[10px] bg-cyan-950 px-1.5 py-0.5 rounded border border-cyan-700 text-cyan-300">PIN: 513513</span>
        </button>
        <button onclick="exportExcel()" class="bg-cyan-500 hover:bg-cyan-400 text-slate-950 px-4 py-2 rounded text-xs font-bold uppercase tracking-wider transition shadow-[0_0_10px_rgba(6,182,212,0.4)] flex items-center gap-2">
            <span>📊</span> Export Excel (Bulanan)
        </button>
        <button onclick="syncToGoogleDrive()" class="hud-btn px-4 py-2 rounded text-xs font-bold uppercase tracking-wider flex items-center gap-2">
            <span>☁️</span> Sync ke Google Drive
        </button>
    </div>

    <!-- Footer ringan -->
    <footer class="text-center text-[11px] text-slate-600 font-mono py-1">
        MZAK Command Center System • Drive: pembangunanmzak@gmail.com • PIN: 513513
    </footer>

    <!-- Skrip JavaScript -->
    <script>
        // Tanggal dan Jam Otomatis
        function updateDateTime() {
            const now = new Date();
            const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit' };
            document.getElementById('datetime').innerText = now.toLocaleDateString('id-ID', options) + " WIB";
        }
        setInterval(updateDateTime, 1000);
        updateDateTime();

        // Hitung Otomatis Total Pagi, Siang, Sore
        function calculateTotals() {
            const tbody = document.getElementById('table-orang').getElementsByTagName('tbody')[0];
            let totalPagi = 0;
            let totalSiang = 0;
            let totalSore = 0;

            for (let i = 0; i < tbody.rows.length; i++) {
                const inputs = tbody.rows[i].getElementsByTagName('input');
                if (inputs.length >= 4) {
                    totalPagi += Number(inputs[1].value) || 0;
                    totalSiang += Number(inputs[2].value) || 0;
                    totalSore += Number(inputs[3].value) || 0;
                }
            }

            document.getElementById('total-pagi').innerText = totalPagi;
            document.getElementById('total-siang').innerText = totalSiang;
            document.getElementById('total-sore').innerText = totalSore;
        }

        // --- SISTEM LOCAL STORAGE (DATABASE 1) ---
        function saveAllData() {
            const tables = ['table-orang', 'table-panitia', 'table-pekerjaan', 'table-informasi'];
            const dataToSave = {};

            tables.forEach(tableId => {
                const tbody = document.getElementById(tableId).getElementsByTagName('tbody')[0];
                let rowsContent = [];
                for (let i = 0; i < tbody.rows.length; i++) {
                    let rowInputs = [];
                    let inputs = tbody.rows[i].getElementsByTagName('input');
                    for (let j = 0; j < inputs.length; j++) {
                        rowInputs.push(inputs[j].value);
                    }
                    rowsContent.push(rowInputs);
                }
                dataToSave[tableId] = rowsContent;
            });

            localStorage.setItem('mzak_app_data', JSON.stringify(dataToSave));
        }

        function loadAllData() {
            const savedData = localStorage.getItem('mzak_app_data');
            if (!savedData) return;

            const parsedData = JSON.parse(savedData);
            Object.keys(parsedData).forEach(tableId => {
                const tableElement = document.getElementById(tableId);
                if (!tableElement) return;
                const tbody = tableElement.getElementsByTagName('tbody')[0];
                tbody.innerHTML = ""; 

                parsedData[tableId].forEach((rowValues, index) => {
                    const rowCount = index + 1;
                    let newRow = tbody.insertRow();
                    newRow.className = "border-b border-slate-900/50 hover:bg-cyan-950/20";
                    let newRowHTML = `<td class="p-1.5 text-center font-mono text-cyan-600">${rowCount}</td>`;

                    if (tableId === 'table-orang') {
                        newRowHTML += `
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="${rowValues[0] || ''}" oninput="calculateTotals(); saveAllData();"></td>
                            <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="${rowValues[1] || 0}" oninput="calculateTotals(); saveAllData();"></td>
                            <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="${rowValues[2] || 0}" oninput="calculateTotals(); saveAllData();"></td>
                            <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="${rowValues[3] || 0}" oninput="calculateTotals(); saveAllData();"></td>
                        `;
                    } else if (tableId === 'table-panitia') {
                        newRowHTML += `
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="${rowValues[0] || ''}" oninput="saveAllData()"></td>
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="${rowValues[1] || ''}" oninput="saveAllData()"></td>
                        `;
                    } else if (tableId === 'table-pekerjaan') {
                        newRowHTML += `
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="${rowValues[0] || ''}" oninput="saveAllData()"></td>
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="${rowValues[1] || ''}" oninput="saveAllData()"></td>
                            <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="${rowValues[2] || 0}" oninput="saveAllData()"></td>
                        `;
                    } else if (tableId === 'table-informasi') {
                        newRowHTML += `
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="${rowValues[0] || ''}" oninput="saveAllData()"></td>
                            <td class="p-1.5"><input type="text" class="hud-input w-full" value="${rowValues[1] || ''}" oninput="saveAllData()"></td>
                        `;
                    }

                    newRowHTML += `<td class="p-1.5 text-center"><button onclick="deleteRow(this)" class="text-red-400 hover:text-red-300 font-bold">×</button></td>`;
                    newRow.innerHTML = newRowHTML;
                });
            });

            calculateTotals();
        }

        window.onload = function() {
            loadAllData();
        };

        function addRow(tableId) {
            const tableElement = document.getElementById(tableId);
            const tbody = tableElement.getElementsByTagName('tbody')[0];
            const rowCount = tbody.rows.length + 1;
            let newRowHTML = '';

            if (tableId === 'table-orang') {
                newRowHTML = `
                    <td class="p-1.5 text-center font-mono text-cyan-600">${rowCount}</td>
                    <td class="p-1.5"><input type="text" class="hud-input w-full" placeholder="Unsur..." oninput="calculateTotals(); saveAllData();"></td>
                    <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="0" oninput="calculateTotals(); saveAllData();"></td>
                    <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="0" oninput="calculateTotals(); saveAllData();"></td>
                    <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="0" oninput="calculateTotals(); saveAllData();"></td>
                    <td class="p-1.5 text-center"><button onclick="deleteRow(this)" class="text-red-400 hover:text-red-300 font-bold">×</button></td>
                `;
            } else if (tableId === 'table-panitia') {
                newRowHTML = `
                    <td class="p-1.5 text-center font-mono text-cyan-600">${rowCount}</td>
                    <td class="p-1.5"><input type="text" class="hud-input w-full" placeholder="Bidang..." oninput="saveAllData()"></td>
                    <td class="p-1.5"><input type="text" class="hud-input w-full" placeholder="Nama PIC..." oninput="saveAllData()"></td>
                    <td class="p-1.5 text-center"><button onclick="deleteRow(this)" class="text-red-400 hover:text-red-300 font-bold">×</button></td>
                `;
            } else if (tableId === 'table-pekerjaan') {
                newRowHTML = `
                    <td class="p-1.5 text-center font-mono text-cyan-600">${rowCount}</td>
                    <td class="p-1.5"><input type="text" class="hud-input w-full" placeholder="Item Pekerjaan..." oninput="saveAllData()"></td>
                    <td class="p-1.5"><input type="text" class="hud-input w-full" placeholder="Nama Tukang..." oninput="saveAllData()"></td>
                    <td class="p-1.5 text-center"><input type="number" class="hud-input w-full text-center" value="0" oninput="saveAllData()"></td>
                    <td class="p-1.5 text-center"><button onclick="deleteRow(this)" class="text-red-400 hover:text-red-300 font-bold">×</button></td>
                `;
            } else if (tableId === 'table-informasi') {
                newRowHTML = `
                    <td class="p-1.5 text-center font-mono text-cyan-600">${rowCount}</td>
                    <td class="p-1.5"><input type="text" class="hud-input w-full" placeholder="Keterangan..." oninput="saveAllData()"></td>
                    <td class="p-1.5"><input type="text" class="hud-input w-full" placeholder="Catatan..." oninput="saveAllData()"></td>
                    <td class="p-1.5 text-center"><button onclick="deleteRow(this)" class="text-red-400 hover:text-red-300 font-bold">×</button></td>
                `;
            }

            const newRow = tbody.insertRow();
            newRow.className = "border-b border-slate-900/50 hover:bg-cyan-950/20";
            newRow.innerHTML = newRowHTML;

            if (tableId === 'table-orang') {
                calculateTotals();
            }
            saveAllData();
        }

        function deleteRow(btn) {
            const row = btn.parentNode.parentNode;
            const tableId = row.closest('table').id;
            row.parentNode.removeChild(row);

            const tbody = document.getElementById(tableId).getElementsByTagName('tbody')[0];
            for (let i = 0; i < tbody.rows.length; i++) {
                tbody.rows[i].cells[0].innerText = i + 1;
            }

            if (tableId === 'table-orang') {
                calculateTotals();
            }
            saveAllData();
        }

        function shareEditor() {
            const pin = "513513";
            const currentUrl = window.location.href;
            prompt(`Salin tautan kolaborasi sistem ini (Google Drive: pembangunanmzak@gmail.com):\n\nPIN KEAMANAN ADMIN: ${pin}`, currentUrl);
        }

        function exportExcel() {
            const wb = XLSX.utils.book_new();
            const tables = ['table-orang', 'table-panitia', 'table-pekerjaan', 'table-informasi'];
            const sheetNames = ['Jml Orang MZAK', 'Panitia Jaga', 'Pekerjaan', 'Informasi'];

            tables.forEach((tblId, index) => {
                const tableElement = document.getElementById(tblId);
                const cloneTable = tableElement.cloneNode(true);
                
                const inputs = cloneTable.querySelectorAll('input');
                inputs.forEach(input => {
                    const span = document.createElement('span');
                    span.textContent = input.value;
                    input.parentNode.replaceChild(span, input);
                });

                const rows = cloneTable.querySelectorAll('tr');
                rows.forEach(row => {
                    if (row.cells.length > 0) {
                        row.deleteCell(row.cells.length - 1);
                    }
                });

                const ws = XLSX.utils.table_to_sheet(cloneTable);
                XLSX.utils.book_append_sheet(wb, ws, sheetNames[index]);
            });

            const dateObj = new Date();
            const year = dateObj.getFullYear();
            const month = String(dateObj.getMonth() + 1).padStart(2, '0');
            const lastDay = new Date(year, dateObj.getMonth() + 1, 0).getDate();
            const filename = `Laporan_Bulanan_MZAK_Tgl_01_s_d_${lastDay}_${month}_${year}.xlsx`;

            XLSX.writeFile(wb, filename);
        }

        // --- SISTEM GOOGLE DRIVE (DATABASE 2) ---
        function syncToGoogleDrive() {
            const webAppUrl = "MASUKKAN_URL_WEB_APP_ANDA_DI_SINI"; // Ganti dengan URL deployment Google Apps Script Anda
            
            if(webAppUrl.includes("MASUKKAN_URL")) {
                alert("Harap masukkan URL Web App Google Apps Script terlebih dahulu pada kode fungsi syncToGoogleDrive().");
                return;
            }

            const tableElement = document.getElementById('table-orang');
            let rowsData = [];
            
            let headerRow = [];
            for (let th of tableElement.querySelectorAll('thead th')) {
                if(th.innerText !== "#") headerRow.push(th.innerText);
            }
            rowsData.push(headerRow);

            for (let tr of tableElement.querySelectorAll('tbody tr')) {
                let rowData = [];
                let inputs = tr.querySelectorAll('input');
                inputs.forEach(input => rowData.push(input.value));
                if(rowData.length > 0) rowsData.push(rowData);
            }

            fetch(webAppUrl, {
                method: "POST",
                mode: "no-cors",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    sheetName: "Jml Orang MZAK",
                    rows: rowsData
                })
            })
            .then(() => {
                alert("Perintah sinkronisasi terkirim ke Google Drive pembangunanmzak@gmail.com!");
            })
            .catch(error => {
                console.error("Gagal:", error);
                alert("Terjadi kesalahan saat menyinkronkan data.");
            });
        }
    </script>
</body>
</html>
