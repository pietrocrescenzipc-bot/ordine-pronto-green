# <!DOCTYPE html>
<html lang="it">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Pronto Green - Inserimento Rapido</title>
    <!-- Tailwind CSS per una grafica moderna e pulita -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- SheetJS per leggere e scrivere i file Excel (.xlsx) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>
<body class="bg-slate-100 text-slate-800 font-sans min-h-screen flex flex-col">

    <!-- Header -->
    <header class="bg-emerald-700 text-white shadow-md py-4 px-6 flex flex-wrap justify-between items-center gap-4">
        <div>
            <h1 class="text-2xl font-bold flex items-center gap-2">🌱 Pronto Green <span class="text-xs bg-emerald-800 px-2 py-1 rounded">Modalità Semplificata</span></h1>
            <p class="text-emerald-100 text-sm">Aggiornamento rapido magazzino senza computer</p>
        </div>
        <div class="flex items-center gap-3">
            <label for="excelFile" class="bg-white text-emerald-800 hover:bg-emerald-50 font-bold px-4 py-2 rounded-xl shadow cursor-pointer transition flex items-center gap-2 text-sm">
                📂 Carica Excel Originale
                <input type="file" id="excelFile" accept=".xlsx, .xls" class="hidden">
            </label>
        </div>
    </header>

    <!-- Main Container -->
    <main class="flex-1 max-w-5xl w-full mx-auto p-4 md:p-6 flex flex-col gap-6">

        <!-- Sezione Istruzioni / Stato File -->
        <div id="statusCard" class="bg-amber-50 border-l-4 border-amber-500 p-4 rounded-r-xl shadow-sm">
            <div class="flex items-center justify-between flex-wrap gap-2">
                <div>
                    <h2 class="font-bold text-amber-900 text-base">⚠️ Attenzione: Nessun file caricato</h2>
                    <p class="text-amber-700 text-sm">Clicca su "Carica Excel Originale" in alto a destra per iniziare, oppure usa la lista di esempio precaricata.</p>
                </div>
                <button id="loadDefaultBtn" class="bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold px-3 py-2 rounded-lg transition">
                    Usa Dati Incorporati
                </button>
            </div>
        </div>

        <!-- Controlli: Selezione Data e Ricerca -->
        <div class="bg-white p-4 rounded-2xl shadow-sm flex flex-col md:flex-row gap-4 items-center justify-between">
            <div class="w-full md:w-1/2 flex flex-col gap-1">
                <label for="dateSelect" class="text-xs font-bold text-slate-500 uppercase">📅 Seleziona Data / Settimana:</label>
                <select id="dateSelect" class="w-full p-3 border border-slate-300 rounded-xl font-semibold text-lg bg-slate-50 focus:ring-2 focus:ring-emerald-500 outline-none">
                    <option value="">-- Carica prima un file Excel --</option>
                </select>
            </div>
            <div class="w-full md:w-1/2 flex flex-col gap-1">
                <label for="searchInput" class="text-xs font-bold text-slate-500 uppercase">🔍 Cerca Prodotto:</label>
                <input type="text" id="searchInput" placeholder="Scrivi il nome del prodotto (es. sesamo, patate...)" class="w-full p-3 border border-slate-300 rounded-xl text-lg bg-slate-50 focus:ring-2 focus:ring-emerald-500 outline-none">
            </div>
        </div>

        <!-- Lista Prodotti -->
        <div class="bg-white rounded-2xl shadow-sm overflow-hidden flex-1">
            <div class="bg-slate-50 px-6 py-3 border-b border-slate-200 flex justify-between items-center">
                <span class="font-bold text-slate-700 text-sm uppercase">Prodotti Disponibili</span>
                <span id="productCount" class="text-xs bg-slate-200 text-slate-700 font-bold px-2.5 py-1 rounded-full">0 prodotti</span>
            </div>
            <div id="productsContainer" class="divide-y divide-slate-100 max-h-[550px] overflow-y-auto p-4 grid grid-cols-1 gap-3">
                <!-- I prodotti verranno inseriti qui dinamicamente -->
                <div class="text-center py-12 text-slate-400">
                    Nessun prodotto disponibile. Carica il file Excel o premi "Usa Dati Incorporati".
                </div>
            </div>
        </div>

        <!-- Barra di Esportazione -->
        <div class="sticky bottom-4 bg-white/90 backdrop-blur p-4 rounded-2xl shadow-lg border border-slate-200 flex justify-between items-center gap-4">
            <div>
                <span id="modifiedStatus" class="text-sm font-semibold text-slate-600">Nessuna modifica effettuata</span>
            </div>
            <button id="exportBtn" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-6 py-3 rounded-xl shadow-md transition flex items-center gap-2 text-base cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed" disabled>
                💾 Scarica Excel Aggiornato
            </button>
        </div>

    </main>

    <!-- Script di Gestione -->
    <script>
        let workbookData = null; // Memorizza il workbook caricato
        let currentSheetName = 'Foglio6';
        let productsList = []; // Lista dei nomi dei prodotti
        let dateColumns = []; // Indici delle colonne delle date e etichette
        let rawSheetData = []; // Matrice completa del foglio excel

        const excelFileInput = document.getElementById('excelFile');
        const dateSelect = document.getElementById('dateSelect');
        const searchInput = document.getElementById('searchInput');
        const productsContainer = document.getElementById('productsContainer');
        const productCount = document.getElementById('productCount');
        const exportBtn = document.getElementById('exportBtn');
        const statusCard = document.getElementById('statusCard');
        const modifiedStatus = document.getElementById('modifiedStatus');
        const loadDefaultBtn = document.getElementById('loadDefaultBtn');

        let currentSelectedDateCol = null;
        let productValuesMap = {}; // Mappa { productName: value } per la data selezionata

        // Gestione caricamento file Excel
        excelFileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    workbookData = XLSX.read(data, { type: 'array', cellDates: true });
                    processWorkbook();
                } catch (err) {
                    alert("Errore nella lettura del file Excel: " + err.message);
                }
            };
            reader.readAsArrayBuffer(file);
        });

        function processWorkbook() {
            if (!workbookData) return;
            currentSheetName = workbookData.SheetNames.includes('Foglio6') ? 'Foglio6' : workbookData.SheetNames[0];
            const worksheet = workbookData.Sheets[currentSheetName];
            rawSheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, defval: "" });

            if (rawSheetData.length < 6) {
                alert("Il formato del file Excel non è valido.");
                return;
            }

            // Estrai le date dalla riga 1 (indice 1)
            dateColumns = [];
            const headerDatesRow = rawSheetData[1];
            for (let colIdx = 0; colIdx < headerDatesRow.length; colIdx++) {
                let cellVal = headerDatesRow[colIdx];
                if (cellVal && cellVal !== 'DATA') {
                    dateColumns.push({ colIndex: colIdx, label: String(cellVal) });
                }
            }

            // Estrai i prodotti dalla colonna 0 (indice 0), partendo dalla riga 5 (indice 5) fino alla fine
            productsList = [];
            for (let rowIdx = 5; rowIdx < rawSheetData.length; rowIdx++) {
                let prodName = rawSheetData[rowIdx][0];
                if (prodName && String(prodName).trim() !== "") {
                    productsList.push({ name: String(prodName).trim(), rowIndex: rowIdx });
                }
            }

            // Popola il menu a tendina delle date
            dateSelect.innerHTML = '<option value="">-- Seleziona una data --</option>';
            dateColumns.forEach(d => {
                let opt = document.createElement('option');
                opt.value = d.colIndex;
                opt.textContent = d.label;
                dateSelect.appendChild(opt);
            });

            statusCard.className = "bg-emerald-50 border-l-4 border-emerald-500 p-4 rounded-r-xl shadow-sm";
            statusCard.innerHTML = `<h2 class="font-bold text-emerald-900 text-base">✅ File caricato con successo!</h2><p class="text-emerald-700 text-sm">Trovati ${productsList.items ? productsList.items.length : productsList.length} prodotti e ${dateColumns.length} date. Seleziona una data per procedere.</p>`;
            
            if (dateColumns.length > 0) {
                dateSelect.value = dateColumns[0].colIndex;
                loadDataForDate(dateColumns[0].colIndex);
            }
        }

        // Carica i dati della data selezionata
        dateSelect.addEventListener('change', (e) => {
            const colIdx = e.target.value;
            if (colIdx !== "") {
                loadDataForDate(parseInt(colIdx));
            } else {
                productsContainer.innerHTML = '<div class="text-center py-12 text-slate-400">Seleziona una data per visualizzare i prodotti.</div>';
                exportBtn.disabled = true;
            }
        });

        function loadDataForDate(colIdx) {
            currentSelectedDateCol = colIdx;
            productValuesMap = {};

            productsList.forEach(p => {
                let val = rawSheetData[p.rowIndex][colIdx];
                productValuesMap[p.name] = val !== undefined && val !== null ? val : "";
            });

            renderProducts();
            exportBtn.disabled = false;
            modifiedStatus.textContent = "Pronto per la modifica";
        }

        // Renderizza la lista dei prodotti con pulsanti grandi
        function renderProducts() {
            const filter = searchInput.value.toLowerCase();
            const filtered = productsList.filter(p => p.name.toLowerCase().includes(filter));

            productCount.textContent = `${filtered.length} prodotti`;

            if (filtered.length === 0) {
                productsContainer.innerHTML = '<div class="text-center py-8 text-slate-400">Nessun prodotto trovato con questa ricerca.</div>';
                return;
            }

            let html = '';
            filtered.forEach(p => {
                let val = productValuesMap[p.name] !== undefined ? productValuesMap[p.name] : "";
                html += `
                    <div class="flex items-center justify-between p-3 bg-white hover:bg-slate-50 rounded-xl border border-slate-100 gap-4 transition">
                        <div class="font-bold text-slate-800 text-lg flex-1">${p.name}</div>
                        <div class="flex items-center gap-2">
                            <button onclick="changeVal('${p.name}', -1)" class="w-12 h-12 bg-rose-100 text-rose-700 font-extrabold text-2xl rounded-xl hover:bg-rose-200 active:scale-95 transition flex items-center justify-center">-</button>
                            <input type="text" value="${val}" id="input_${p.name.replace(/[^a-zA-Z0-9]/g, '_')}" onchange="updateVal('${p.name}', this.value)" class="w-24 h-12 text-center text-xl font-bold border-2 border-slate-200 rounded-xl bg-slate-50 focus:border-emerald-500 outline-none">
                            <button onclick="changeVal('${p.name}', 1)" class="w-12 h-12 bg-emerald-100 text-emerald-700 font-extrabold text-2xl rounded-xl hover:bg-emerald-200 active:scale-95 transition flex items-center justify-center">+</button>
                        </div>
                    </div>
                `;
            });
            productsContainer.innerHTML = html;
        }

        window.changeVal = function(prodName, delta) {
            let current = parseFloat(productValuesMap[prodName]);
            if (isNaN(current)) current = 0;
            let newVal = current + delta;
            if (newVal < 0) newVal = 0;
            productValuesMap[prodName] = newVal;

            // Aggiorna input visivo senza perdere focus se possibile
            let inputId = `input_${prodName.replace(/[^a-zA-Z0-9]/g, '_')}`;
            let el = document.getElementById(inputId);
            if (el) el.value = newVal;

            modifiedStatus.textContent = "Modifiche non salvate";
            modifiedStatus.className = "text-sm font-bold text-amber-600";
        }

        window.updateVal = function(prodName, val) {
            productValuesMap[prodName] = val;
            modifiedStatus.textContent = "Modifiche non salvate";
            modifiedStatus.className = "text-sm font-bold text-amber-600";
        }

        searchInput.addEventListener('input', () => {
            renderProducts();
        });

        // Pulsante di Esportazione Excel
        exportBtn.addEventListener('click', () => {
            if (!workbookData || currentSelectedDateCol === null) return;

            const worksheet = workbookData.Sheets[currentSheetName];

            // Aggiorna i valori nel foglio raw
            productsList.forEach(p => {
                let val = productValuesMap[p.name];
                let cellAddress = XLSX.utils.encode_cell({ r: p.rowIndex, c: currentSelectedDateCol });
                if (!worksheet[cellAddress]) {
                    worksheet[cellAddress] = { t: 's', v: val };
                } else {
                    worksheet[cellAddress].v = val;
                    worksheet[cellAddress].t = typeof val === 'number' ? 'n' : 's';
                }
            });

            // Genera il file e scaricalo
            const wbout = XLSX.write(workbookData, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([wbout], { type: 'application/octet-stream' });
            
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = "PRONTO_GREEN_Aggiornato.xlsx";
            link.click();

            modifiedStatus.textContent = "File esportato con successo!";
            modifiedStatus.className = "text-sm font-bold text-emerald-600";
        });
    </script>
</body>
</html>
