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

    const startDirectOcrBtn = document.getElementById('startDirectOcrBtn');

    let ocrStream = null;
    let pendingQrCode = ""; // Guardar el QR que no se encontró

    startDirectOcrBtn.addEventListener('click', () => {
        pendingQrCode = ""; // Limpiar porque es un registro directo
        startOcrCamera();
    });

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

    // --- REGISTRATION MODAL ELEMENTS ---
    const registerModal = document.getElementById('registerModal');
    const regIdInput = document.getElementById('regIdInput');
    const regSerieInput = document.getElementById('regSerieInput');
    const regLocationSelect = document.getElementById('regLocationSelect');
    const confirmRegBtn = document.getElementById('confirmRegBtn');
    const cancelRegBtn = document.getElementById('cancelRegBtn');

    let globalLocations = []; // Cache list of unique locations

    // Populate Location Dropdown
    function populateLocations() {
        if (!globalDataRaw || !globalHeaders || globalHeaders.length < 4) return;

        const locKey = globalHeaders[3]; // Columna 4 (index 3)
        // Extract unique, non-empty locations
        const locSet = new Set();
        globalDataRaw.forEach(row => {
            const val = row[locKey];
            if (val && typeof val === 'string' && val.trim() !== '') {
                locSet.add(val.trim());
            }
        });

        globalLocations = Array.from(locSet).sort();

        // Render options
        regLocationSelect.innerHTML = '<option value="">-- Seleccionar --</option>';
        globalLocations.forEach(loc => {
            const opt = document.createElement('option');
            opt.value = loc;
            opt.textContent = loc;
            regLocationSelect.appendChild(opt);
        });
    }

    // Call this after processing data
    function refreshLocations() {
        populateLocations();
    }

    // Modal Control
    cancelRegBtn.addEventListener('click', () => {
        registerModal.classList.add('hidden');
    });

    confirmRegBtn.addEventListener('click', () => {
        const idVal = regIdInput.value.trim();
        const serieVal = regSerieInput.value.trim();
        const locVal = regLocationSelect.value;

        if (!idVal) { alert("Debes ingresar un ID (Col. 1)."); return; }
        if (!serieVal) { alert("Debes tener un N° Serie (OCR)."); return; }

        // --- VALIDACIÓN DE DUPLICADOS (AHORA AQUÍ) ---
        if (globalHeaders.length > 2) {
            const serieKey = globalHeaders[2];
            const normalize = s => (s || '').toString().trim().toUpperCase();
            const targetSerial = normalize(serieVal);

            const duplicateIndex = globalDataRaw.findIndex(row => normalize(row[serieKey]) === targetSerial);

            if (duplicateIndex !== -1) {
                const proceed = confirm(`⚠️ ¡ADVERTENCIA DE DUPLICADO!\n\nEl N° Serie "${serieVal}" ya existe en la fila ${duplicateIndex + 2}.\n\n¿Quieres guardarlo de todas formas?`);
                if (!proceed) return;
            }
        }

        saveNewEquipment(idVal, serieVal, locVal);
        registerModal.classList.add('hidden');
    });

    function saveNewEquipment(id, serie, location) {
        if (!globalHeaders || globalHeaders.length < 5) {
            alert("Error: Estructura de Excel insuficiente.");
            return;
        }

        const newRow = {};
        globalHeaders.forEach(h => newRow[h] = ""); // Init defaults

        newRow[globalHeaders[0]] = id;      // Col 1: ID
        if (globalHeaders.length > 2) newRow[globalHeaders[2]] = serie;    // Col 3: Serie
        if (globalHeaders.length > 3) newRow[globalHeaders[3]] = location; // Col 4: Ubicacion
        if (globalHeaders.length > 4) newRow[globalHeaders[4]] = new Date(); // Col 5: Fecha

        globalDataRaw.push(newRow);

        processData(globalDataRaw);
        refreshLocations(); // Update list in case we want to reuse logic later for adding new locations (not implemented yet but good practice)

        alert(`✅ Equipo guardado!`);
        qrInput.value = "";
        editPanel.classList.add('hidden');
        scanResult.classList.add('hidden');
    }

    // --- UPDATED OCR CALLBACK ---
    // --- UPDATED OCR CALLBACK (NATIVE + TESSERACT FALLBACK) ---
    captureBtn.addEventListener('click', () => {
        // 1. Validar que el video esté listo
        if (ocrVideo.readyState !== ocrVideo.HAVE_ENOUGH_DATA) {
            alert("La cámara aún no está lista. Por favor espera unos segundos e intenta de nuevo.");
            return;
        }

        const w = ocrVideo.videoWidth;
        const h = ocrVideo.videoHeight;

        if (!w || !h) {
            alert("Error: Dimensiones de video inválidas. Reinicia la cámara.");
            return;
        }

        ocrCanvas.width = w;
        ocrCanvas.height = h;
        const ctx = ocrCanvas.getContext('2d');
        ctx.drawImage(ocrVideo, 0, 0, w, h);

        // --- NUEVO: PRE-PROCESAMIENTO DE IMAGEN ---
        // Convertir a escala de grises y aumentar contraste para mejorar OCR
        applyImagePreprocessing(ctx, w, h);

        // Detener video para ahorrar recursos mientras procesa
        if (ocrStream) ocrStream.getTracks().forEach(track => track.stop());

        ocrControls.classList.add('hidden');
        ocrProcessing.classList.remove('hidden');
        ocrBar.style.width = "10%";

        startSmartOCR(ocrCanvas);
    });

    function applyImagePreprocessing(ctx, w, h) {
        const imgData = ctx.getImageData(0, 0, w, h);
        const d = imgData.data;

        // === PASO 1: Convertir a escala de grises y construir histograma ===
        const histogram = new Array(256).fill(0);
        const grays = new Uint8Array(d.length / 4);

        for (let i = 0; i < d.length; i += 4) {
            const r = d[i];
            const g = d[i + 1];
            const b = d[i + 2];
            // Fórmula ITU-R BT.601 para luminosidad
            const gray = Math.round(0.299 * r + 0.587 * g + 0.114 * b);
            grays[i / 4] = gray;
            histogram[gray]++;
        }

        // === PASO 2: Calcular umbral óptimo con MÉTODO DE OTSU ===
        const totalPixels = d.length / 4;
        let sum = 0;
        for (let i = 0; i < 256; i++) sum += i * histogram[i];

        let sumB = 0;
        let wB = 0; // Weight Background
        let wF = 0; // Weight Foreground

        let varMax = 0;
        let threshold = 128; // Fallback default

        for (let t = 0; t < 256; t++) {
            wB += histogram[t];
            if (wB === 0) continue;

            wF = totalPixels - wB;
            if (wF === 0) break;

            sumB += t * histogram[t];

            const mB = sumB / wB;            // Mean Background
            const mF = (sum - sumB) / wF;    // Mean Foreground

            // Varianza entre clases
            const varBetween = wB * wF * (mB - mF) * (mB - mF);

            if (varBetween > varMax) {
                varMax = varBetween;
                threshold = t;
            }
        }

        console.log(`[OCR Preprocessing] Otsu Threshold Calculado: ${threshold}`);

        // === PASO 3: Aplicar Binarización con el umbral calculado ===
        for (let i = 0; i < d.length; i += 4) {
            const val = grays[i / 4] > threshold ? 255 : 0;
            d[i] = val;
            d[i + 1] = val;
            d[i + 2] = val;
            // Alpha (d[i+3]) se mantiene igual
        }

        ctx.putImageData(imgData, 0, 0);
    }

    async function startSmartOCR(canvas) {
        const updateStatus = (msg) => {
            const statusText = document.querySelector('#ocrProcessing p');
            if (statusText) statusText.textContent = msg;
        };

        const handleResult = (text, method) => {
            console.log(`[OCR Success] Method: ${method}, Text: ${text}`);

            // Limpieza más agresiva: Solo letras, números y guiones
            const cleanedSerial = text.replace(/[^a-zA-Z0-9-]/g, '').toUpperCase().trim();

            if (!cleanedSerial || cleanedSerial.length < 3) {
                alert(`No se pudo leer un serial válido con ${method}.\nTexto detectado: "${text.trim()}"\n\nIntenta mejorar la iluminación.`);
                closeOcrModal();
                return;
            }

            // ÉXITO
            closeOcrModal();
            registerModal.classList.remove('hidden');
            regIdInput.value = pendingQrCode || "";
            regSerieInput.value = cleanedSerial;
            populateLocations();
            setTimeout(() => regSerieInput.focus(), 100);

            // Notificar visualmente qué método se usó
            // (Opcional: Podríamos agregar un toast, pero por ahora el log basta)
            console.log("OCR Final Cleaned:", cleanedSerial);
        };

        // 1. INTENTO NATIVO (Shape Detection API)
        if ('TextDetector' in window) {
            try {
                updateStatus("⚡ Usando Motor Nativo (Rápido)...");
                console.log("[OCR] Attempting Native TextDetector...");
                const textDetector = new TextDetector();
                const texts = await textDetector.detect(canvas);

                if (texts.length > 0) {
                    const rawText = texts.map(t => t.rawValue).join(' ');
                    // Verificar si parece un resultado decente (al menos 3 chars)
                    if (rawText.replace(/[^a-zA-Z0-9]/g, '').length >= 3) {
                        ocrBar.style.width = "100%";
                        handleResult(rawText, "NATIVE_API");
                        return;
                    }
                }
                console.warn("[OCR] Native API returned empty or poor results. Fallback to Tesseract.");
            } catch (e) {
                console.error("[OCR] Native API Error:", e);
                // Fallback continúa abajo
            }
        } else {
            console.log("[OCR] Native TextDetector not supported in this browser.");
        }

        // 2. FALLBACK TESSERACT
        updateStatus("🐢 Usando Tesseract (Motor Secundario)...");
        runTesseract(canvas, updateStatus, handleResult);
    }

    function runTesseract(canvas, updateStatusRef, callbackSuccess) {
        Tesseract.recognize(
            canvas.toDataURL('image/png'),
            'eng',
            {
                logger: m => {
                    console.log("[Tesseract]", m);
                    if (m.status === 'recognizing text') {
                        ocrBar.style.width = `${Math.floor(m.progress * 100)}%`;
                        updateStatusRef(`⏳ Analizando... ${Math.floor(m.progress * 100)}%`);
                    } else if (m.status === 'loading tesseract core') {
                        updateStatusRef("⏳ Cargando Tesseract Core...");
                    } else if (m.status === 'loading language traineddata') {
                        updateStatusRef("⏳ Descargando datos (requiere internet la 1ra vez)...");
                    }
                },
                tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789- '
            }
        ).then(({ data: { text } }) => {
            callbackSuccess(text, "TESSERACT");
        }).catch(err => {
            console.error("CRITICAL OCR ERROR:", err);
            alert(`Error CRÍTICO de OCR:\n${err.message || err}`);
            closeOcrModal();
        });
    }

    // Modificar processData para actualizar Locations al cargar archivo
    const originalProcessData = processData;
    processData = function (data) {
        originalProcessData(data);
        setTimeout(refreshLocations, 100); // Async helper
    };


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
