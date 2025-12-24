document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const fileLabelText = document.querySelector('.file-label .text');
    const processBtn = document.getElementById('processBtn');
    const loadingDiv = document.getElementById('loading');
    const resultsArea = document.getElementById('resultsArea');
    const exportBtn = document.getElementById('exportBtn');

    // Variables Globales
    let globalDataRaw = [];
    let globalHeaders = [];
    let globalWorkbook = null;
    let globalFirstSheetName = "";

    // Elementos de Verificación / Edición
    const editPanel = document.getElementById('editPanel');
    const matchInfo = document.getElementById('matchInfo');
    const dateInput = document.getElementById('dateInput');
    const updateBtn = document.getElementById('updateBtn');
    let currentMatchIndex = -1;

    // Elementos de Registro de Serie
    const registerSerieBtn = document.getElementById('registerSerieBtn');
    const registerModal = document.getElementById('registerModal');
    const regSerieInput = document.getElementById('regSerieInput');
    const regLocationSelect = document.getElementById('regLocationSelect');
    const confirmRegBtn = document.getElementById('confirmRegBtn');
    const cancelRegBtn = document.getElementById('cancelRegBtn');
    const regFeedback = document.getElementById('regFeedback');

    // --- MANEJO DE ARCHIVOS ---
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            fileLabelText.textContent = `Archivo seleccionado: ${e.target.files[0].name}`;
            fileLabelText.style.color = '#818cf8';
        }
    });

    processBtn.addEventListener('click', () => {
        const file = fileInput.files[0];
        if (!file) {
            alert('Por favor selecciona un archivo Excel primero.');
            return;
        }

        loadingDiv.classList.remove('hidden');
        processBtn.disabled = true;
        processBtn.textContent = 'Procesando...';
        resultsArea.classList.add('hidden');
        exportBtn.disabled = true;

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

                processData(globalDataRaw);
                populateLocations();

                loadingDiv.classList.add('hidden');
                resultsArea.classList.remove('hidden');

                if (globalDataRaw.length > 0) {
                    exportBtn.disabled = false;
                    registerSerieBtn.disabled = false;
                }

            } catch (error) {
                console.error(error);
                alert('Error al leer el archivo Excel.');
            } finally {
                processBtn.disabled = false;
                processBtn.textContent = 'Procesar Archivo';
            }
        };

        reader.readAsArrayBuffer(file);
    });

    // --- EXPORTAR EXCEL ---
    exportBtn.addEventListener('click', () => {
        if (!globalDataRaw || globalDataRaw.length === 0) {
            alert("No hay datos para exportar.");
            return;
        }

        try {
            const newSheet = XLSX.utils.json_to_sheet(globalDataRaw);
            const newWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWb, newSheet, globalFirstSheetName || "Sheet1");
            XLSX.writeFile(newWb, "Equipos_Actualizados.xlsx");
        } catch (err) {
            console.error(err);
            alert("Error al exportar el archivo.");
        }
    });

    // --- HELPERS ---
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
        let d = val;
        if (!(d instanceof Date)) d = new Date(val);
        if (isNaN(d.getTime())) return "";

        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
    }

    function getColumnKey(columnName) {
        // Buscar header que contenga el nombre (case insensitive)
        return globalHeaders.find(h => h.toLowerCase().includes(columnName.toLowerCase())) || null;
    }

    function processData(rawData) {
        if (rawData.length === 0) return;

        const serieKey = getColumnKey('serie');
        
        const sortedData = [...rawData].sort((a, b) => {
            const idA = (a[globalHeaders[0]] || '').toString();
            const idB = (b[globalHeaders[0]] || '').toString();
            return idA.localeCompare(idB);
        });

        renderMainTable('tableMain', sortedData);
    }

    function renderMainTable(tableId, dataList) {
        const table = document.getElementById(tableId);
        const thead = table.querySelector('thead');
        const tbody = table.querySelector('tbody');

        thead.innerHTML = '';
        tbody.innerHTML = '';

        if (dataList.length === 0) return;

        const headers = Object.keys(dataList[0]);
        const headerRow = document.createElement('tr');
        headers.forEach(h => {
            const th = document.createElement('th');
            th.textContent = h;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);

        dataList.forEach(row => {
            const tr = document.createElement('tr');
            headers.forEach(h => {
                const td = document.createElement('td');
                let val = row[h];
                if (val instanceof Date) val = formatDate(val);
                td.textContent = (val !== undefined && val !== null) ? val : '';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    }

    // --- POBLAR SELECTOR DE UBICACIONES ---
    function populateLocations() {
        const locKey = getColumnKey('ubicacion');
        if (!locKey) {
            console.warn("No se encontró columna de ubicación");
            return;
        }

        const locSet = new Set();
        globalDataRaw.forEach(row => {
            const val = row[locKey];
            if (val && typeof val === 'string' && val.trim() !== '') {
                locSet.add(val.trim());
            }
        });

        const locations = Array.from(locSet).sort();

        regLocationSelect.innerHTML = '<option value="">-- Seleccionar Ubicación --</option>';
        locations.forEach(loc => {
            const opt = document.createElement('option');
            opt.value = loc;
            opt.textContent = loc;
            regLocationSelect.appendChild(opt);
        });
    }

    // --- REGISTRO DE NUEVA SERIE ---
    registerSerieBtn.addEventListener('click', () => {
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

    confirmRegBtn.addEventListener('click', () => {
        const serieVal = regSerieInput.value.trim().toUpperCase();
        const locVal = regLocationSelect.value;

        if (!serieVal) {
            alert("Ingresa un número de serie.");
            regSerieInput.focus();
            return;
        }

        if (!locVal) {
            alert("Selecciona una ubicación técnica.");
            return;
        }

        const serieKey = getColumnKey('serie');
        const locKey = getColumnKey('ubicacion');

        if (!serieKey) {
            alert("No se encontró la columna 'Serie' en el Excel.");
            return;
        }

        // Verificar si la serie ya existe
        const normalize = s => (s || '').toString().trim().toUpperCase();
        const existingIndex = globalDataRaw.findIndex(row => normalize(row[serieKey]) === serieVal);

        if (existingIndex !== -1) {
            regFeedback.textContent = `⚠️ La serie "${serieVal}" ya existe en la fila ${existingIndex + 2}.`;
            regFeedback.style.color = '#f87171';
            regFeedback.classList.remove('hidden');
            return;
        }

        // Crear nueva fila
        const newRow = {};
        globalHeaders.forEach(h => newRow[h] = "");

        newRow[serieKey] = serieVal;
        if (locKey) newRow[locKey] = locVal;

        // Agregar fecha actual si existe columna de fecha/calibración
        const dateKey = getColumnKey('calibracion') || getColumnKey('fecha');
        if (dateKey) newRow[dateKey] = new Date();

        globalDataRaw.push(newRow);

        // Actualizar tabla
        processData(globalDataRaw);

        regFeedback.textContent = `✅ Serie "${serieVal}" registrada correctamente.`;
        regFeedback.style.color = '#4ade80';
        regFeedback.classList.remove('hidden');

        // Limpiar inputs
        regSerieInput.value = '';
        regLocationSelect.value = '';

        setTimeout(() => {
            registerModal.classList.add('hidden');
            regFeedback.classList.add('hidden');
        }, 1500);
    });

    // --- QR LOGIC ---
    const startScanBtn = document.getElementById('startScanBtn');
    const stopScanBtn = document.getElementById('stopScanBtn');
    const qrInput = document.getElementById('qrInput');
    const scanResult = document.getElementById('scanResult');
    const readerDiv = document.getElementById('reader');

    let html5QrcodeScanner = null;

    startScanBtn.addEventListener('click', () => {
        readerDiv.classList.remove('hidden');
        startScanBtn.classList.add('hidden');
        stopScanBtn.classList.remove('hidden');
        editPanel.classList.add('hidden');
        scanResult.classList.add('hidden');

        html5QrcodeScanner = new Html5Qrcode("reader");
        const config = { fps: 30, qrbox: { width: 250, height: 250 }, aspectRatio: 1.0 };

        html5QrcodeScanner.start({ facingMode: "environment" }, config, onScanSuccess)
            .catch(err => {
                console.error("Camera Error", err);
                alert("Error iniciando la cámara: " + err);
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
        console.log(`Scan: ${decodedText}`);

        let finalValue = decodedText;
        try {
            if (decodedText.includes('%2F') || decodedText.includes('%2f')) {
                const parts = decodedText.split(/%2F|%2f/);
                finalValue = parts[parts.length - 1].replace(/%/g, '-');
            }
        } catch (e) { console.error(e); }

        qrInput.value = finalValue;
        stopScanning();

        if (globalDataRaw.length > 0 && globalHeaders.length > 0) {
            findAndSetupEdit(finalValue);
        } else {
            showScanFeedback(`Escaneado: ${finalValue} (⚠️ Carga un Excel para buscar)`, false);
        }
    }

    function showScanFeedback(msg, isSuccess) {
        scanResult.classList.remove('hidden');
        scanResult.innerHTML = msg;
        scanResult.style.color = isSuccess ? '#4ade80' : '#f87171';
        scanResult.style.borderColor = isSuccess ? '#4ade80' : '#f87171';
    }

    function findAndSetupEdit(scannedId) {
        const idKey = globalHeaders[0];
        const serieKey = getColumnKey('serie');
        const normalize = (s) => (s || '').toString().trim().toUpperCase();
        const target = normalize(scannedId);

        // Buscar primero por ID (columna 1)
        let index = globalDataRaw.findIndex(row => normalize(row[idKey]) === target);
        
        // Si no encuentra por ID, buscar por Serie
        if (index === -1 && serieKey) {
            index = globalDataRaw.findIndex(row => normalize(row[serieKey]) === target);
        }

        if (index !== -1) {
            currentMatchIndex = index;
            const rowData = globalDataRaw[index];

            showScanFeedback(`✅ ¡Encontrado en fila ${index + 2}!`, true);

            editPanel.classList.remove('hidden');
            matchInfo.innerHTML = `<strong>ID:</strong> ${rowData[idKey]}<br><strong>Serie:</strong> ${serieKey ? rowData[serieKey] : 'N/A'}<br><strong>Fila Excel:</strong> ${index + 2}`;

            const dateKey = getColumnKey('calibracion') || getColumnKey('fecha');
            if (dateKey) {
                const currentDateVal = rowData[dateKey];
                dateInput.value = formatDateForInput(currentDateVal);
                dateInput.dataset.targetKey = dateKey;
                dateInput.disabled = false;
                updateBtn.disabled = false;
            } else {
                dateInput.disabled = true;
                updateBtn.disabled = true;
            }

        } else {
            // NO ENCONTRADO - Abrir modal de registro con el código escaneado
            currentMatchIndex = -1;
            editPanel.classList.add('hidden');
            showScanFeedback(`⚠️ "${scannedId}" no encontrado. Puedes registrarlo como nueva serie.`, false);

            // Abrir modal de registro con el valor escaneado
            regSerieInput.value = scannedId;
            regLocationSelect.value = '';
            regFeedback.classList.add('hidden');
            registerModal.classList.remove('hidden');
        }
    }

    // --- ACTUALIZAR FECHA ---
    updateBtn.addEventListener('click', () => {
        if (currentMatchIndex === -1) return;

        const newDateVal = dateInput.value;
        if (!newDateVal) {
            alert("Selecciona una fecha válida.");
            return;
        }

        const targetKey = dateInput.dataset.targetKey;
        if (!targetKey) return;

        const [y, m, d] = newDateVal.split('-').map(Number);
        const dateObj = new Date(y, m - 1, d);

        globalDataRaw[currentMatchIndex][targetKey] = dateObj;

        alert(`¡Fecha actualizada para fila ${currentMatchIndex + 2}!`);
        editPanel.classList.add('hidden');
        qrInput.value = "";

        processData(globalDataRaw);
    });

});
