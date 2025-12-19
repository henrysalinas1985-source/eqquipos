document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const fileLabelText = document.querySelector('.file-label .text');
    const processBtn = document.getElementById('processBtn');
    const loadingDiv = document.getElementById('loading');
    const resultsArea = document.getElementById('resultsArea');

    // --- DEBUG LOGGER ---
    const debugConsole = document.getElementById('debugConsole');
    const toggleDebug = document.getElementById('toggleDebug');

    toggleDebug.addEventListener('click', () => {
        debugConsole.style.display = debugConsole.style.display === 'none' ? 'block' : 'none';
    });

    function logToScreen(msg, type = 'INFO') {
        const line = document.createElement('div');
        line.style.color = type === 'ERROR' ? '#f87171' : '#4ade80';
        line.style.borderBottom = '1px solid #333';
        line.textContent = `[${new Date().toLocaleTimeString()}] [${type}] ${msg}`;
        debugConsole.appendChild(line);
        debugConsole.scrollTop = debugConsole.scrollHeight;
    }

    // Sobreescribir console para capturar errores de Tesseract/Cámara
    const originalLog = console.log;
    const originalError = console.error;

    console.log = (...args) => {
        originalLog(...args);
        logToScreen(args.join(' '));
    };

    console.error = (...args) => {
        originalError(...args);
        logToScreen(args.join(' '), 'ERROR');
    };

    // Variables Globales para Estado
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

        // Detectar columnas estrictas por índice (0 y 4)
        // Guardamos las keys reales para usarlas
        const keys = Object.keys(rawData[0]);
        const keyId = keys[0]; // Columna 1 -> ID
        const keyCalib = (keys.length > 4) ? keys[4] : null; // Columna 5 -> Calibración

        console.log("Columnas Estrictas detectadas:", { keyId, keyCalib });

        // Ordenar datos principales (Detalle) solo por ID
        const sortedData = [...rawData].sort((a, b) => {
            const idA = (a[keyId] || '').toString();
            const idB = (b[keyId] || '').toString();
            if (idA < idB) return -1;
            if (idA > idB) return 1;
            return 0;
        });

        renderMainTable('tableMain', sortedData, keyCalib);
    }

    // Ya no se usa renderSummaryTable

    function renderMainTable(tableId, dataList, keyCalib) {
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

                // Si es la columna de calibración (Col 5)
                if (h === keyCalib) {
                    val = formatDate(val);
                } else if (val instanceof Date) {
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

    // --- OCR LOGIC ---
    const ocrModal = document.getElementById('ocrModal');
    const ocrVideo = document.getElementById('ocrVideo');
    const ocrCanvas = document.getElementById('ocrCanvas');
    const captureBtn = document.getElementById('captureBtn');
    const cancelOcrBtn = document.getElementById('cancelOcrBtn');
    const ocrProcessing = document.getElementById('ocrProcessing');
    const ocrControls = document.getElementById('ocrControls');
    const ocrBar = document.getElementById('ocrBar');

    let ocrStream = null;
    let pendingQrCode = ""; // Guardar el QR que no se encontró

    function startOcrCamera() {
        ocrModal.classList.remove('hidden');
        ocrControls.classList.remove('hidden');
        ocrProcessing.classList.add('hidden');

        navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } })
            .then(stream => {
                ocrStream = stream;
                ocrVideo.srcObject = stream;
            })
            .catch(err => {
                console.error("Error cámara OCR:", err);
                alert("No se pudo acceder a la cámara para OCR.");
                closeOcrModal();
            });
    }

    function closeOcrModal() {
        if (ocrStream) {
            ocrStream.getTracks().forEach(track => track.stop());
            ocrStream = null;
        }
        ocrModal.classList.add('hidden');
    }

    cancelOcrBtn.addEventListener('click', closeOcrModal);

    captureBtn.addEventListener('click', () => {
        // Capturar frame
        const w = ocrVideo.videoWidth;
        const h = ocrVideo.videoHeight;
        ocrCanvas.width = w;
        ocrCanvas.height = h;
        const ctx = ocrCanvas.getContext('2d');
        ctx.drawImage(ocrVideo, 0, 0, w, h);

        // Detener cámara temporalmente para procesar
        if (ocrStream) {
            ocrStream.getTracks().forEach(track => track.stop());
        }

        // Mostrar UI de proceso
        ocrControls.classList.add('hidden');
        ocrProcessing.classList.remove('hidden');
        ocrBar.style.width = "0%";

        // Iniciar Tesseract
        Tesseract.recognize(
            ocrCanvas.toDataURL('image/png'),
            'eng', // Usar inglés para mejor reconocimiento de letras/números (vs spa)
            {
                logger: m => {
                    if (m.status === 'recognizing text') {
                        ocrBar.style.width = `${Math.floor(m.progress * 100)}%`;
                    }
                }
            }
        ).then(({ data: { text } }) => {
            console.log("OCR Result:", text);

            // Limpieza básica: Alfanuméricos y guiones
            const cleanedSerial = text.replace(/[^A-Za-z0-9-]/g, '').trim();

            if (!cleanedSerial) {
                alert("No se detectó texto legible. Inténtalo de nuevo.");
                closeOcrModal();
                return;
            }

            // Confirmar resultado
            const confirmedSerial = prompt(`Texto detectado: ${cleanedSerial} \n\n¿Es correcto el Número de Serie? (Edítalo si es necesario)`, cleanedSerial);

            if (confirmedSerial) {
                createNewRow(pendingQrCode, confirmedSerial);
            }
            closeOcrModal();
        }).catch(err => {
            console.error(err);
            alert("Error al procesar la imagen: " + err.message);
            closeOcrModal();
        });
    });

    function createNewRow(qrCode, serialNumber) {
        if (!globalHeaders || globalHeaders.length < 5) {
            alert("Error: El Excel no tiene estructura suficiente para agregar filas.");
            return;
        }

        const newRow = {};
        // Llenar con strings vacíos por defecto
        globalHeaders.forEach(h => newRow[h] = "");

        // Asignar valores específicos
        // Columna 1 (Index 0): Equipo (QR)
        newRow[globalHeaders[0]] = qrCode;

        // Columna 3 (Index 2): Serie
        if (globalHeaders.length > 2) {
            newRow[globalHeaders[2]] = serialNumber;
        }

        // Columna 5 (Index 4): Fecha Actual (por defecto)
        if (globalHeaders.length > 4) {
            newRow[globalHeaders[4]] = new Date();
        }

        // Agregar a datos globales
        globalDataRaw.push(newRow);

        // Refrescar tabla
        processData(globalDataRaw);

        // Feedback
        alert(`✅ ¡Nuevo equipo agregado!\n\nID: ${qrCode}\nSerie: ${serialNumber}`);
        qrInput.value = "";
        editPanel.classList.add('hidden');
        scanResult.classList.add('hidden');
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
            if (globalHeaders.length >= 5) {
                const dateKey = globalHeaders[4];
                const currentDateVal = rowData[dateKey];

                dateInput.value = formatDateForInput(currentDateVal);
                dateInput.dataset.targetKey = dateKey;
                dateInput.disabled = false;
                updateBtn.disabled = false;
            } else {
                alert("El Excel tiene menos de 5 columnas. No se puede actualizar la columna 5.");
                dateInput.disabled = true;
                updateBtn.disabled = true;
            }

        } else {
            // NO ENCONTRADO
            currentMatchIndex = -1;
            editPanel.classList.add('hidden');

            // Feedback con opción de OCR
            pendingQrCode = scannedId;
            scanResult.classList.remove('hidden');
            scanResult.style.color = '#f87171';
            scanResult.style.borderColor = '#f87171';
            scanResult.innerHTML = `
                ❌ ID no encontrado.<br>
                <button id="startOcrBtn" class="secondary-btn" style="margin-top:10px; font-size: 0.9em; background: #3b82f6; border-color: #3b82f6; color: white;">
                    📷 Capturar Serie (OCR)
                </button>
            `;

            // Bind click del botón dinámico
            document.getElementById('startOcrBtn').addEventListener('click', startOcrCamera);
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

        const [y, m, d] = newDateVal.split('-').map(Number);
        const dateObj = new Date(y, m - 1, d); // Mes 0-index

        globalDataRaw[currentMatchIndex][targetKey] = dateObj;

        alert(`¡Fecha actualizada para fila ${currentMatchIndex + 2}!`);
        editPanel.classList.add('hidden');
        qrInput.value = "";

        processData(globalDataRaw);
    });

});
