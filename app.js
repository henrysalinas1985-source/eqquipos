document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const fileLabelText = document.querySelector('.file-label .text');
    const processBtn = document.getElementById('processBtn');
    const loadingDiv = document.getElementById('loading');
    const resultsArea = document.getElementById('resultsArea');

    // Variables Globales para Estado
    let globalDataRaw = [];
    let globalHeaders = [];
    let globalWorkbook = null;
    let globalFirstSheetName = "";

    // Elementos UI adicionales
    const exportBtn = document.getElementById('exportBtn');

    // Elementos de la sección Verificación / Edición
    const editPanel = document.getElementById('editPanel');
    const matchInfo = document.getElementById('matchInfo');
    const dateInput = document.getElementById('dateInput');
    const updateBtn = document.getElementById('updateBtn');

    // Estado de la selección actual
    let currentMatchIndex = -1; // Índice en globalDataRaw

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

                // Leer JSON con defval ""
                globalDataRaw = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

                if (globalDataRaw.length > 0) {
                    globalHeaders = Object.keys(globalDataRaw[0]);
                    console.log("Headers Globales:", globalHeaders);
                }

                processData(globalDataRaw);

                loadingDiv.classList.add('hidden');
                resultsArea.classList.remove('hidden');

                // Habilitar Exportar si hay datos
                if (globalDataRaw.length > 0) {
                    exportBtn.disabled = false;
                }

            } catch (error) {
                console.error(error);
                alert('Error al leer el archivo Excel. Asegúrate de que sea un archivo válido.');
            } finally {
                processBtn.disabled = false;
                processBtn.textContent = 'Procesar Archivo';
            }
        };

        reader.readAsArrayBuffer(file);
    });

    // --- ACCIÓN: EXPORTAR EXCEL ---
    exportBtn.addEventListener('click', () => {
        if (!globalDataRaw || globalDataRaw.length === 0) {
            alert("No hay datos para exportar.");
            return;
        }

        try {
            // Convertir datos actuales (que pueden haber sido modificados) a hoja
            const newSheet = XLSX.utils.json_to_sheet(globalDataRaw);

            // Crear nuevo libro
            const newWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWb, newSheet, globalFirstSheetName || "Sheet1");

            // Descargar
            XLSX.writeFile(newWb, "Equipos_Actualizados.xlsx");
        } catch (err) {
            console.error(err);
            alert("Error al exportar el archivo.");
        }
    });


    // --- HELPERS ---
    function findKey(row, keywords) {
        if (!row) return null;
        const keys = Object.keys(row);
        const normalize = (s) => s.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

        return keys.find(k => {
            const normKey = normalize(k);
            return keywords.some(kw => normKey.includes(normalize(kw)));
        });
    }

    function formatDate(val) {
        if (val instanceof Date) return val.toLocaleDateString('es-ES'); // Ej: 19/12/2025

        if (typeof val === 'string') {
            const d = new Date(val);
            if (!isNaN(d.getTime())) {
                return d.toLocaleDateString('es-ES');
            }
        }
        return val;
    }

    // Ajuste para formatear fechas para input type="date" (YYYY-MM-DD)
    function formatDateForInput(val) {
        if (!val) return "";
        let d = val;
        if (!(d instanceof Date)) {
            d = new Date(val);
        }
        if (isNaN(d.getTime())) return "";

        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
    }

    function processData(rawData) {
        if (rawData.length === 0) return;

        const firstRow = rawData[0];
        const keyUnidad = findKey(firstRow, ['unidad', 'unit', 'ubicacion']) || 'UNIDAD';
        const keyRango = findKey(firstRow, ['rango', 'range']) || 'RANGO';
        const keyFecha = findKey(firstRow, ['fecha', 'date', 'calib']) || 'FECHA DE CALIBRACION';

        // Lógica de conteo para Resumen
        const counts = {};
        rawData.forEach(row => {
            const unidad = row[keyUnidad] || 'Sin Unidad';
            const rango = row[keyRango] || 'Sin Rango';
            const compositeKey = `${unidad}|||${rango}`;

            if (!counts[compositeKey]) {
                counts[compositeKey] = { unidad, rango, cantidad: 0 };
            }
            counts[compositeKey].cantidad++;
        });

        const summaryList = Object.values(counts).sort((a, b) => {
            if (a.unidad < b.unidad) return -1;
            if (a.unidad > b.unidad) return 1;
            if (a.rango < b.rango) return -1;
            if (a.rango > b.rango) return 1;
            return 0;
        });

        const sortedData = [...rawData].sort((a, b) => {
            // Orden simple para visualización
            const uA = (a[keyUnidad] || '').toString();
            const uB = (b[keyUnidad] || '').toString();
            if (uA < uB) return -1;
            if (uA > uB) return 1;
            return 0;
        });

        renderResults(summaryList, sortedData, keyFecha);
    }

    function renderResults(summaryList, sortedData, keyFecha) {
        if (summaryList) renderSummaryTable('tableResumen', summaryList);
        if (sortedData) renderMainTable('tableMain', sortedData, keyFecha);
    }

    function renderSummaryTable(tableId, list) {
        const tbody = document.querySelector(`#${tableId} tbody`);
        tbody.innerHTML = '';
        list.forEach(item => {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td>${item.unidad}</td><td>${item.rango}</td><td>${item.cantidad}</td>`;
            tbody.appendChild(tr);
        });
    }

    function renderMainTable(tableId, dataList, keyFecha) {
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
                if (h === keyFecha || val instanceof Date) {
                    val = formatDate(val);
                }
                td.textContent = (val !== undefined && val !== null) ? val : '';
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    }

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

        // Reset UI de edición
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

    function onScanSuccess(decodedText, decodedResult) {
        console.log(`Scan: ${decodedText}`);

        // --- 1. PROCESAR FORMATO QR ---
        let finalValue = decodedText;
        try {
            // Estrategia 1: Sharepoint encoded (CEN%2D...)
            if (decodedText.includes('%2F') || decodedText.includes('%2f')) {
                const parts = decodedText.split(/%2F|%2f/);
                const lastSegment = parts[parts.length - 1];
                // CEN%2D123 -> CEN-2D123
                finalValue = lastSegment.replace(/%/g, '-');

            } else if (decodedText.includes('%')) {
                // Estrategia 2: Desactivada anteriormente, solo replace simple si se desea
                finalValue = decodedText; // Dejar tal cual o aplicar lógica si es necesario
            }
        } catch (e) { console.error(e); }

        qrInput.value = finalValue;
        stopScanning();

        // --- 2. BUSCAR EN EXCEL (Columna 1) ---
        if (globalDataRaw.length > 0 && globalHeaders.length > 0) {
            findAndSetupEdit(finalValue);
        } else {
            showScanFeedback(`Escaneado: ${finalValue} (⚠️ Carga un Excel para buscar)`, false);
        }
    }

    function showScanFeedback(msg, isSuccess) {
        scanResult.classList.remove('hidden');
        scanResult.innerHTML = msg;
        scanResult.style.color = isSuccess ? '#4ade80' : '#f87171'; // Green : Red
        scanResult.style.borderColor = isSuccess ? '#4ade80' : '#f87171';
    }

    function findAndSetupEdit(scannedId) {
        // Asumiendo Columna 1 = index 0
        const idKey = globalHeaders[0];

        // Búsqueda (Case Insensitive trim)
        const normalize = (s) => (s || '').toString().trim().toUpperCase();
        const target = normalize(scannedId);

        const index = globalDataRaw.findIndex(row => normalize(row[idKey]) === target);

        if (index !== -1) {
            currentMatchIndex = index;
            const rowData = globalDataRaw[index];

            showScanFeedback(`✅ ¡Encontrado en fila ${index + 2}!`, true);

            // Setup Edit Panel
            editPanel.classList.remove('hidden');
            matchInfo.innerHTML = `<strong>ID:</strong> ${rowData[idKey]}<br><strong>Fila Excel:</strong> ${index + 2}`;

            // Asumiendo Columna 5 = index 4 para la Fecha
            // Verificar si existe esa columna
            if (globalHeaders.length >= 5) {
                const dateKey = globalHeaders[4];
                const currentDateVal = rowData[dateKey];

                // Poner valor actual en input date
                dateInput.value = formatDateForInput(currentDateVal);

                // Guardar key para uso en update
                dateInput.dataset.targetKey = dateKey;
            } else {
                alert("El Excel tiene menos de 5 columnas. No se puede actualizar la columna 5.");
                dateInput.disabled = true;
                updateBtn.disabled = true;
            }

        } else {
            currentMatchIndex = -1;
            editPanel.classList.add('hidden');
            showScanFeedback(`❌ ID "${scannedId}" no encontrado en Columna "${idKey}".`, false);
        }
    }

    // --- ACCIÓN: ACTUALIZAR FECHA ---
    updateBtn.addEventListener('click', () => {
        if (currentMatchIndex === -1) return;

        const newDateVal = dateInput.value; // YYYY-MM-DD string
        if (!newDateVal) {
            alert("Selecciona una fecha válida.");
            return;
        }

        const targetKey = dateInput.dataset.targetKey;
        if (!targetKey) return;

        // Actualizar datos globales
        // XLSX sheet_to_json usa fechas nativas si cellDates: true, 
        // pero para homogeneizar podemos guardar string o date. 
        // Al escribir de nuevo, si es Date object es mejor.

        // Crear fecha local (el input date da yyyy-mm-dd UTC a veces, mejor parsing simple)
        const [y, m, d] = newDateVal.split('-').map(Number);
        const dateObj = new Date(y, m - 1, d); // Mes 0-index

        globalDataRaw[currentMatchIndex][targetKey] = dateObj;

        // Feedback UI
        alert(`¡Fecha actualizada para fila ${currentMatchIndex + 2}!`);
        editPanel.classList.add('hidden');
        qrInput.value = "";

        // Refrescar tabla visualmente (importante para que el usuario crea que pasó algo)
        // Re-ejecutar processData es pesado pero seguro
        processData(globalDataRaw);
    });

});
