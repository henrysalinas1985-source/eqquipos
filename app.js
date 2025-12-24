document.addEventListener('DOMContentLoaded', () => {
    // Elementos principales
    const fileInput = document.getElementById('fileInput');
    const fileLabel = document.getElementById('fileLabel');
    const processBtn = document.getElementById('processBtn');
    const loadingDiv = document.getElementById('loading');
    const resultsArea = document.getElementById('resultsArea');
    const exportBtn = document.getElementById('exportBtn');

    // Variables Globales
    let globalDataRaw = [];
    let globalHeaders = [];
    let globalWorkbook = null;
    let globalFirstSheetName = "";

    // Elementos de Edición
    const editPanel = document.getElementById('editPanel');
    const matchInfo = document.getElementById('matchInfo');
    const dateInput = document.getElementById('dateInput');
    const updateBtn = document.getElementById('updateBtn');
    let currentMatchIndex = -1;

    // Elementos de Registro
    const registerSerieBtn = document.getElementById('registerSerieBtn');
    const registerModal = document.getElementById('registerModal');
    const regSerieInput = document.getElementById('regSerieInput');
    const regLocationSelect = document.getElementById('regLocationSelect');
    const confirmRegBtn = document.getElementById('confirmRegBtn');
    const cancelRegBtn = document.getElementById('cancelRegBtn');
    const regFeedback = document.getElementById('regFeedback');

    // QR Elements
    const startScanBtn = document.getElementById('startScanBtn');
    const stopScanBtn = document.getElementById('stopScanBtn');
    const readerDiv = document.getElementById('reader');
    const scanResult = document.getElementById('scanResult');
    let html5QrcodeScanner = null;

    // --- ARCHIVO ---
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            fileLabel.textContent = e.target.files[0].name;
            fileLabel.style.color = '#00d9ff';
        }
    });

    processBtn.addEventListener('click', () => {
        const file = fileInput.files[0];
        if (!file) {
            alert('Selecciona un archivo Excel primero.');
            return;
        }

        loadingDiv.classList.remove('hidden');
        processBtn.disabled = true;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                globalWorkbook = XLSX.read(data, { type: 'array', cellDates: true });
                globalFirstSheetName = globalWorkbook.SheetNames[0];
                const worksheet = globalWorkbook.Sheets[globalFirstSheetName];
                globalDataRaw = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

                if (globalDataRaw.length > 0) {
                    globalHeaders = Object.keys(globalDataRaw[0]);
                    console.log("Headers:", globalHeaders);
                }

                renderTable();
                populateLocations();

                loadingDiv.classList.add('hidden');
                resultsArea.classList.remove('hidden');
                exportBtn.disabled = false;
                registerSerieBtn.classList.remove('disabled');

            } catch (error) {
                console.error(error);
                alert('Error al leer el archivo.');
                loadingDiv.classList.add('hidden');
            }
            processBtn.disabled = false;
        };
        reader.readAsArrayBuffer(file);
    });

    // --- EXPORTAR ---
    exportBtn.addEventListener('click', () => {
        if (globalDataRaw.length === 0) return;
        try {
            const newSheet = XLSX.utils.json_to_sheet(globalDataRaw);
            const newWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWb, newSheet, globalFirstSheetName || "Sheet1");
            XLSX.writeFile(newWb, "Equipos_Actualizados.xlsx");
        } catch (err) {
            alert("Error al exportar.");
        }
    });

    // --- HELPERS ---
    function getColumnKey(name) {
        // Normalizar: quitar tildes y convertir a minúsculas
        const normalize = str => str.toLowerCase()
            .normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // Quita tildes
        
        const searchTerm = normalize(name);
        
        return globalHeaders.find(h => normalize(h).includes(searchTerm)) || null;
    }

    function formatDate(val) {
        if (val instanceof Date) return val.toLocaleDateString('es-ES');
        if (typeof val === 'string') {
            const d = new Date(val);
            if (!isNaN(d.getTime())) return d.toLocaleDateString('es-ES');
        }
        return val;
    }

    function formatDateForInput(val) {
        if (!val) return "";
        let d = val instanceof Date ? val : new Date(val);
        if (isNaN(d.getTime())) return "";
        return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    }

    function renderTable() {
        const thead = document.querySelector('#tableMain thead');
        const tbody = document.querySelector('#tableMain tbody');
        thead.innerHTML = '';
        tbody.innerHTML = '';

        if (globalDataRaw.length === 0) return;

        const headerRow = document.createElement('tr');
        globalHeaders.forEach(h => {
            const th = document.createElement('th');
            th.textContent = h;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);

        globalDataRaw.forEach(row => {
            const tr = document.createElement('tr');
            globalHeaders.forEach(h => {
                const td = document.createElement('td');
                let val = row[h];
                if (val instanceof Date) val = formatDate(val);
                td.textContent = val !== undefined && val !== null ? val : '';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    }

    function populateLocations() {
        // Buscar columna que contenga "ubicacion" o "tecnica" o "location"
        const locKey = getColumnKey('ubicacion') || getColumnKey('tecnica') || getColumnKey('location');
        
        console.log("Headers disponibles:", globalHeaders);
        console.log("Columna ubicación encontrada:", locKey);
        
        if (!locKey) {
            console.warn("No se encontró columna de ubicación. Headers:", globalHeaders);
            // Mostrar todas las columnas en el select para debug
            regLocationSelect.innerHTML = '<option value="">-- No se encontró columna ubicación --</option>';
            return;
        }

        const locSet = new Set();
        globalDataRaw.forEach(row => {
            const val = row[locKey];
            if (val && String(val).trim()) {
                locSet.add(String(val).trim());
            }
        });

        console.log("Ubicaciones encontradas:", Array.from(locSet));

        regLocationSelect.innerHTML = '<option value="">-- Seleccionar Ubicación --</option>';
        Array.from(locSet).sort().forEach(loc => {
            const opt = document.createElement('option');
            opt.value = loc;
            opt.textContent = loc;
            regLocationSelect.appendChild(opt);
        });
    }

    // --- REGISTRO SERIE ---
    registerSerieBtn.addEventListener('click', () => {
        if (registerSerieBtn.classList.contains('disabled')) return;
        if (globalDataRaw.length === 0) {
            alert("Primero carga un archivo Excel.");
            return;
        }
        regSerieInput.value = '';
        regLocationSelect.value = '';
        regFeedback.classList.add('hidden');
        registerModal.classList.remove('hidden');
        regSerieInput.focus();
    });

    cancelRegBtn.addEventListener('click', () => {
        registerModal.classList.add('hidden');
    });

    // Filtro en tiempo real mientras escribe la serie
    regSerieInput.addEventListener('input', () => {
        const searchVal = regSerieInput.value.trim().toUpperCase();
        
        if (searchVal.length < 2) {
            regFeedback.classList.add('hidden');
            return;
        }

        const serieKey = getColumnKey('serie');
        if (!serieKey) return;

        const normalize = s => String(s || '').trim().toUpperCase();
        
        // Buscar coincidencia exacta
        const exactMatch = globalDataRaw.findIndex(row => normalize(row[serieKey]) === searchVal);
        
        if (exactMatch !== -1) {
            const row = globalDataRaw[exactMatch];
            const locKey = getColumnKey('ubicacion') || getColumnKey('tecnica');
            regFeedback.textContent = `⚠️ Serie existe (fila ${exactMatch + 2}) - Ubicación: ${locKey ? row[locKey] : 'N/A'}`;
            regFeedback.className = 'feedback warning';
            regFeedback.classList.remove('hidden');
        } else {
            // Buscar coincidencias parciales
            const partialMatches = globalDataRaw.filter(row => normalize(row[serieKey]).includes(searchVal));
            
            if (partialMatches.length > 0 && partialMatches.length <= 5) {
                const matches = partialMatches.map(r => r[serieKey]).join(', ');
                regFeedback.textContent = `🔍 Similares: ${matches}`;
                regFeedback.className = 'feedback warning';
                regFeedback.classList.remove('hidden');
            } else if (partialMatches.length === 0) {
                regFeedback.textContent = `✅ Serie disponible para registrar`;
                regFeedback.className = 'feedback success';
                regFeedback.classList.remove('hidden');
            } else {
                regFeedback.classList.add('hidden');
            }
        }
    });

    confirmRegBtn.addEventListener('click', () => {
        const serieVal = regSerieInput.value.trim().toUpperCase();
        const locVal = regLocationSelect.value;

        if (!serieVal) {
            alert("Ingresa un número de serie.");
            return;
        }
        if (!locVal) {
            alert("Selecciona una ubicación.");
            return;
        }

        const serieKey = getColumnKey('serie');
        const locKey = getColumnKey('ubicacion') || getColumnKey('tecnica');

        if (!serieKey) {
            alert("No se encontró columna 'Serie' en el Excel. Columnas disponibles: " + globalHeaders.join(', '));
            return;
        }

        // Verificar duplicado
        const normalize = s => String(s || '').trim().toUpperCase();
        const exists = globalDataRaw.findIndex(row => normalize(row[serieKey]) === serieVal);

        if (exists !== -1) {
            regFeedback.textContent = `⚠️ Serie "${serieVal}" ya existe (fila ${exists + 2})`;
            regFeedback.className = 'feedback error';
            regFeedback.classList.remove('hidden');
            return;
        }

        // Crear nueva fila
        const newRow = {};
        globalHeaders.forEach(h => newRow[h] = "");
        newRow[serieKey] = serieVal;
        if (locKey) newRow[locKey] = locVal;

        const dateKey = getColumnKey('calibracion') || getColumnKey('fecha');
        if (dateKey) newRow[dateKey] = new Date();

        globalDataRaw.push(newRow);
        renderTable();

        regFeedback.textContent = `✅ Serie "${serieVal}" registrada`;
        regFeedback.className = 'feedback success';
        regFeedback.classList.remove('hidden');

        setTimeout(() => {
            registerModal.classList.add('hidden');
        }, 1200);
    });

    // --- QR ---
    startScanBtn.addEventListener('click', () => {
        if (globalDataRaw.length === 0) {
            alert("Primero carga un archivo Excel.");
            return;
        }

        readerDiv.classList.remove('hidden');
        startScanBtn.classList.add('hidden');
        stopScanBtn.classList.remove('hidden');
        editPanel.classList.add('hidden');
        scanResult.classList.add('hidden');

        html5QrcodeScanner = new Html5Qrcode("reader");
        html5QrcodeScanner.start(
            { facingMode: "environment" },
            { fps: 15, qrbox: { width: 250, height: 250 } },
            onScanSuccess
        ).catch(err => {
            alert("Error cámara: " + err);
            stopScanning();
        });
    });

    stopScanBtn.addEventListener('click', stopScanning);

    function stopScanning() {
        if (html5QrcodeScanner) {
            html5QrcodeScanner.stop().then(() => {
                readerDiv.classList.add('hidden');
                startScanBtn.classList.remove('hidden');
                stopScanBtn.classList.add('hidden');
                html5QrcodeScanner.clear();
            }).catch(console.error);
        }
    }

    function onScanSuccess(decodedText) {
        let finalValue = decodedText;
        if (decodedText.includes('%2F') || decodedText.includes('%2f')) {
            const parts = decodedText.split(/%2F|%2f/);
            finalValue = parts[parts.length - 1].replace(/%/g, '-');
        }

        stopScanning();
        findOrRegister(finalValue);
    }

    function findOrRegister(scannedValue) {
        const normalize = s => String(s || '').trim().toUpperCase();
        const target = normalize(scannedValue);

        const idKey = globalHeaders[0];
        const serieKey = getColumnKey('serie');

        // Buscar por ID o Serie
        let index = globalDataRaw.findIndex(row => normalize(row[idKey]) === target);
        if (index === -1 && serieKey) {
            index = globalDataRaw.findIndex(row => normalize(row[serieKey]) === target);
        }

        if (index !== -1) {
            // ENCONTRADO
            currentMatchIndex = index;
            const row = globalDataRaw[index];

            scanResult.textContent = `✅ Encontrado en fila ${index + 2}`;
            scanResult.className = 'feedback success';
            scanResult.classList.remove('hidden');

            matchInfo.innerHTML = `<strong>ID:</strong> ${row[idKey] || 'N/A'} | <strong>Serie:</strong> ${serieKey ? row[serieKey] : 'N/A'}`;
            
            const dateKey = getColumnKey('calibracion') || getColumnKey('fecha');
            if (dateKey) {
                dateInput.value = formatDateForInput(row[dateKey]);
                dateInput.dataset.targetKey = dateKey;
            }

            editPanel.classList.remove('hidden');

        } else {
            // NO ENCONTRADO - Abrir registro
            scanResult.textContent = `⚠️ "${scannedValue}" no encontrado. Registrando...`;
            scanResult.className = 'feedback warning';
            scanResult.classList.remove('hidden');

            regSerieInput.value = scannedValue;
            regLocationSelect.value = '';
            regFeedback.classList.add('hidden');
            registerModal.classList.remove('hidden');
        }
    }

    // --- ACTUALIZAR FECHA ---
    updateBtn.addEventListener('click', () => {
        if (currentMatchIndex === -1) return;

        const newDate = dateInput.value;
        if (!newDate) {
            alert("Selecciona una fecha.");
            return;
        }

        const targetKey = dateInput.dataset.targetKey;
        if (!targetKey) return;

        const [y, m, d] = newDate.split('-').map(Number);
        globalDataRaw[currentMatchIndex][targetKey] = new Date(y, m - 1, d);

        alert(`✅ Fecha actualizada (fila ${currentMatchIndex + 2})`);
        editPanel.classList.add('hidden');
        scanResult.classList.add('hidden');
        renderTable();
    });
});
