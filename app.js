document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const fileLabelText = document.querySelector('.file-label .text');
    const processBtn = document.getElementById('processBtn');
    const loadingDiv = document.getElementById('loading');
    const resultsArea = document.getElementById('resultsArea');

    // Feedback visual al seleccionar archivo
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

        // Usar FileReader para leer el archivo localmente
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });

                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" }); // defval to avoid undefined

                processData(jsonData);

                loadingDiv.classList.add('hidden');
                resultsArea.classList.remove('hidden');

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

    /**
     * Busca una key en el objeto row de forma "fuzzy" (ignorando mayúsculas, tildes, espacios)
     */
    function findKey(row, keywords) {
        if (!row) return null;
        const keys = Object.keys(row);
        // Normalizar string: quitar tildes, mayúsculas, espacios trim
        const normalize = (s) => s.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

        return keys.find(k => {
            const normKey = normalize(k);
            return keywords.some(kw => normKey.includes(normalize(kw)));
        });
    }

    function formatDate(val) {
        if (val instanceof Date) return val.toLocaleDateString('es-ES');

        if (typeof val === 'string') {
            // Un formato largo como "Wed Oct 01 2025..."
            const d = new Date(val);
            if (!isNaN(d.getTime())) {
                return d.toLocaleDateString('es-ES');
            }
        }
        return val;
    }

    function processData(rawData) {
        if (rawData.length === 0) return;

        // Detectar columnas reales
        const firstRow = rawData[0];
        const keyUnidad = findKey(firstRow, ['unidad', 'unit', 'ubicacion']) || 'UNIDAD';
        const keyRango = findKey(firstRow, ['rango', 'range']) || 'RANGO';
        const keyFecha = findKey(firstRow, ['fecha', 'date', 'calib']) || 'FECHA DE CALIBRACION';

        console.log("Columnas detectadas:", { keyUnidad, keyRango, keyFecha });

        // 1. Agrupación por UNIDAD y RANGO (SOLO)
        // Clave compuesta: Unidad|Rango
        const counts = {};

        rawData.forEach(row => {
            const unidad = row[keyUnidad] || 'Sin Unidad';
            const rango = row[keyRango] || 'Sin Rango';

            const compositeKey = `${unidad}|||${rango}`;

            if (!counts[compositeKey]) {
                counts[compositeKey] = {
                    unidad: unidad,
                    rango: rango,
                    cantidad: 0
                };
            }
            counts[compositeKey].cantidad++;
        });

        // Convertir a lista
        const summaryList = Object.values(counts);

        // Ordenar resumen: Unidad -> Rango
        summaryList.sort((a, b) => {
            if (a.unidad < b.unidad) return -1;
            if (a.unidad > b.unidad) return 1;
            if (a.rango < b.rango) return -1;
            if (a.rango > b.rango) return 1;
            return 0;
        });


        // 2. Ordenar datos principales (Detalle)
        const sortedData = [...rawData].sort((a, b) => {
            const uA = (a[keyUnidad] || '').toString();
            const uB = (b[keyUnidad] || '').toString();
            if (uA < uB) return -1;
            if (uA > uB) return 1;

            const rA = (a[keyRango] || '').toString();
            const rB = (b[keyRango] || '').toString();
            if (rA < rB) return -1;
            if (rA > rB) return 1;

            const fA = a[keyFecha] instanceof Date ? a[keyFecha] : new Date(0);
            const fB = b[keyFecha] instanceof Date ? b[keyFecha] : new Date(0);
            return fB - fA;
        });

        // Pasar keyFecha para formatear esa columna específicamente
        renderResults(summaryList, sortedData, keyFecha);
    }

    function renderResults(summaryList, sortedData, keyFecha) {
        renderSummaryTable('tableResumen', summaryList);
        renderMainTable('tableMain', sortedData, keyFecha);
    }

    function renderSummaryTable(tableId, list) {
        const tbody = document.querySelector(`#${tableId} tbody`);
        tbody.innerHTML = '';

        list.forEach(item => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${item.unidad}</td>
                <td>${item.rango}</td>
                <td>${item.cantidad}</td>
            `;
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

                // Si es la columna de fecha detectada, o parece fecha
                if (h === keyFecha) {
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

    // --- Lógica para Escaneo de QR y Verificación ---
    const startScanBtn = document.getElementById('startScanBtn');
    const stopScanBtn = document.getElementById('stopScanBtn');
    const qrInput = document.getElementById('qrInput');
    const compareInput = document.getElementById('compareInput');
    const compareBtn = document.getElementById('compareBtn');
    const scanResult = document.getElementById('scanResult');
    const readerDiv = document.getElementById('reader');

    let html5QrcodeScanner = null;

    startScanBtn.addEventListener('click', () => {
        readerDiv.classList.remove('hidden');
        startScanBtn.classList.add('hidden');
        stopScanBtn.classList.remove('hidden');

        // Configuración más robusta para QRs densos
        html5QrcodeScanner = new Html5Qrcode("reader");

        const qrBoxSize = Math.min(window.innerWidth, window.innerHeight) * 0.7;

        const config = {
            fps: 30, // Mayor FPS para escanear más rápido
            qrbox: { width: qrBoxSize, height: qrBoxSize },
            aspectRatio: 1.0,
            experimentalFeatures: {
                useBarCodeDetectorIfSupported: true // Usa detector nativo del celular (más rápido)
            }
        };

        // Configuración de cámara simplificada para evitar errores de la librería
        // La librería exige solo una llave (facingMode o deviceId) en este objeto.
        const cameraConfig = {
            facingMode: "environment"
        };

        // Manejo de errores más detallado
        html5QrcodeScanner.start(cameraConfig, config, onScanSuccess)
            .catch(err => {
                console.error("Error iniciando cámara", err);
                let msg = "No se pudo iniciar la cámara.";

                if (err.name === 'NotAllowedError' || err.toString().includes('Permission')) {
                    msg = "Permiso denegado. Debes 'Permitir' el acceso a la cámara en el navegador.";
                } else if (err.toString().includes('HTTPS')) {
                    msg = "El navegador bloqueó la cámara por seguridad. Asegúrate de usar HTTPS o localhost.";
                } else if (err.toString().includes('found 2 keys')) {
                    // Fallback a config simple si la avanzada falla en este dispositivo
                    console.warn("Falló config avanzada, intentando básica...");
                    html5QrcodeScanner.start({ facingMode: "environment" }, config, onScanSuccess);
                    return;
                }

                alert(msg + "\nDetalle: " + err);
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
            }).catch(err => {
                console.error("Error deteniendo cámara", err);
            });
        }
    }

    function onScanSuccess(decodedText, decodedResult) {
        console.log(`Código escaneado: ${decodedText}`);

        let finalValue = decodedText;

        // Lógica de extracción personalizada para URLs
        // "3 letras antes del % y los numeros restantes del enlace a la parte final"
        if (decodedText.includes('%')) {
            try {
                // 1. Buscar la posición del % (usamos el último por si hay varios, o el primero?)
                // Asumiremos el primero que cumpla la estructura lógica, o el último índice.
                // Generalmente estos códigos tienen un separador claro. Probemos buscando "%".
                const parts = decodedText.split('%');
                if (parts.length > 1) {
                    const prefixPart = parts[0];
                    // Tomar las 3 letras antes del %
                    let threeChars = "";
                    if (prefixPart.length >= 3) {
                        threeChars = prefixPart.substring(prefixPart.length - 3);
                    } else {
                        threeChars = prefixPart;
                    }

                    // 2. Buscar números al final del string completo
                    // Regex para capturar dígitos al final de la línea
                    const matches = decodedText.match(/(\d+)$/);
                    let endNumbers = "";
                    if (matches && matches[1]) {
                        endNumbers = matches[1];
                    }

                    if (threeChars || endNumbers) {
                        finalValue = `${threeChars.toUpperCase()}${endNumbers}`;
                        console.log(`Dato extraído: ${finalValue} (de ${decodedText})`);
                    }
                }
            } catch (e) {
                console.error("Error parseando URL", e);
            }
        }

        qrInput.value = finalValue;
        stopScanning();

        // Feedback visual
        scanResult.classList.remove('hidden');
        scanResult.innerHTML = `
            <div style="text-align:center;">
                <strong>¡Dato Detectado!</strong><br>
                <span style="font-size:1.2em; color:#4ade80;">${finalValue}</span><br>
                <span style="font-size:0.8em; color:#94a3b8;">(Original: ${decodedText})</span>
            </div>
        `;

        setTimeout(() => {
            scanResult.classList.add('hidden');
        }, 4000);
    }

    compareBtn.addEventListener('click', () => {
        const val1 = qrInput.value.trim();
        const val2 = compareInput.value.trim();

        if (!val1) {
            alert("Por favor escanea o escribe un código primero.");
            return;
        }

        // TODO: Aquí implementaremos la lógica real de comparación con el Excel
        // Por ahora, solo comparamos los dos inputs si existen
        if (val2) {
            if (val1.toLowerCase() === val2.toLowerCase()) {
                scanResult.textContent = "✅ ¡Coincidencia Exacta!";
                scanResult.style.color = "#4ade80"; // green
                scanResult.style.borderColor = "#4ade80";
            } else {
                scanResult.textContent = "❌ No coinciden.";
                scanResult.style.color = "#f87171"; // red
                scanResult.style.borderColor = "#f87171";
            }
        } else {
            scanResult.textContent = `Dato validado: ${val1} (Sin comparación externa)`;
            scanResult.style.color = "#d1fae5";
        }
        scanResult.classList.remove('hidden');
    });

});
