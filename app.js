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
    const scanResult = document.getElementById('scanResult');
    const readerDiv = document.getElementById('reader');
    
    // Elementos de la nueva tabla
    const scannedTableBody = document.querySelector('#scannedTable tbody');
    const exportBtn = document.getElementById('exportBtn');
    const clearScansBtn = document.getElementById('clearScansBtn');

    let html5QrcodeScanner = null;
    
    // Base de datos externa (del Excel cargado) y Escaneos de sesión
    let externalDatabaseSet = new Set(); // Para búsqueda rápida de coincidencias
    let sessionScans = []; // { equipo: string, serie: string, status: string, timestamp: Date }

    // Interceptar la carga de datos para llenar nuestra "Base de Datos" (Columna 1)
    const originalProcessData = processData;
    processData = function(rawData) {
        // Llamar a la función original para mantener la funcionalidad de estadísticas
        originalProcessData(rawData);

        // Lógica adicional: Extraer columna 1 para validación
        if (rawData.length > 0) {
            externalDatabaseSet.clear();
            const firstRow = rawData[0];
            const keys = Object.keys(firstRow);
            if (keys.length > 0) {
                const firstKey = keys[0]; // Asumimos Columna 1 = index 0
                // O usamos 'Campo Equipo' si existe
                const keyEquipo = findKey(firstRow, ['equipo', 'device', 'id']) || firstKey;
                
                console.log(`Usando columna '${keyEquipo}' como referencia de 'Campo Equipo'`);

                rawData.forEach(row => {
                    const val = row[keyEquipo];
                    if (val) {
                        // Normalizamos para comparar
                        externalDatabaseSet.add(val.toString().trim().toUpperCase());
                    }
                });
                alert(`Base de datos cargada. ${externalDatabaseSet.size} equipos registrados para validación.`);
            }
        }
    };

    startScanBtn.addEventListener('click', () => {
        readerDiv.classList.remove('hidden');
        startScanBtn.classList.add('hidden');
        stopScanBtn.classList.remove('hidden');

        html5QrcodeScanner = new Html5Qrcode("reader");
        const qrBoxSize = Math.min(window.innerWidth, window.innerHeight) * 0.7;
        const config = {
            fps: 30,
            qrbox: { width: qrBoxSize, height: qrBoxSize },
            aspectRatio: 1.0,
            experimentalFeatures: { useBarCodeDetectorIfSupported: true }
        };

        html5QrcodeScanner.start({ facingMode: "environment" }, config, onScanSuccess)
            .catch(err => {
                console.error("Error camara", err);
                alert("No se pudo iniciar la cámara: " + err);
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
        
        let processedValue = decodedText.trim();
        let isSerialNumber = false;

        // 1. Detectar si es Número de Serie (Solo dígitos)
        if (/^\d+$/.test(processedValue)) {
            isSerialNumber = true;
        } else {
            // 2. Lógica de Guion (Letras - Números)
            // Ejemplo: ABC1234 -> ABC-1234
            // Regex: Captura grupo de letras al inicio, luego números
            const parts = processedValue.match(/^([A-Za-z]+)(\d+)$/);
            if (parts) {
                processedValue = `${parts[1].toUpperCase()}-${parts[2]}`;
            } else {
                // Fallback: Si ya tiene guion o formato raro, normalizar mayúsculas
                processedValue = processedValue.toUpperCase();
            }
        }

        qrInput.value = processedValue;
        
        // 3. Agregar a la lista de sesión
        addScanToSession(processedValue, isSerialNumber);

        // Feedback visual temporal
        scanResult.classList.remove('hidden');
        scanResult.innerHTML = `✅ Leído: ${processedValue}`;
        setTimeout(() => scanResult.classList.add('hidden'), 2000);
    }

    function addScanToSession(value, isSerialNumber) {
        let entry = {
            equipo: '',
            serie: '',
            estado: 'Nuevo', // Nuevo, Registrado (Duplicado)
            raw: value
        };

        if (isSerialNumber) {
            // Si es serie, buscamos el último item ingresado que NO tenga serie, o creamos uno nuevo vacío?
            // "si es el numero de serie el que se escanea ponerlo en la columna 2"
            // Asumiremos que se refiere al item ACTUAL o el último escaneado.
            // Para simplificar, si el último scan fue un Equipo y le falta Serie, se la pegamos.
            const lastItem = sessionScans.length > 0 ? sessionScans[sessionScans.length - 1] : null;
            
            if (lastItem && !lastItem.serie && lastItem.equipo) {
                lastItem.serie = value;
                updateTable();
                return; 
            } else {
                // Si no hay anterior, o el anterior ya tiene serie, ¿creamos fila nueva solo con serie?
                // El requerimiento es ambiguo, pero asumiremos fila nueva.
                entry.serie = value;
                entry.equipo = '(Pendiente)';
            }
        } else {
            // Es un Campo Equipo
            entry.equipo = value;
            
            // 4. Comparar con Excel (Columna 1)
            // "Si esta ya registrado poner ese calor en rojo en el nuevo excel"
            // Buscamos si existe en externalDatabaseSet
            // OJO: La comparación debería ser contra el Excel cargado.
            if (externalDatabaseSet.has(value)) {
                entry.estado = 'Registrado'; // Se marcará en rojo visualmente
            }
        }

        sessionScans.push(entry);
        updateTable();
    }

    function updateTable() {
        scannedTableBody.innerHTML = '';
        
        // Mostrar del más reciente al más antiguo o viceversa?
        // Generalmente orden de llegada (append)
        sessionScans.forEach((item, index) => {
            const tr = document.createElement('tr');
            
            // Si está "Registrado" (encontrado en excel), poner en rojo
            if (item.estado === 'Registrado') {
                tr.style.color = '#ef4444'; // Red-500
                tr.style.fontWeight = 'bold';
            }

            tr.innerHTML = `
                <td>${item.equipo}</td>
                <td>${item.serie || '-'}</td>
                <td>${item.estado}</td>
            `;
            scannedTableBody.appendChild(tr);
        });
    }

    // --- Exportar Excel ---
    exportBtn.addEventListener('click', () => {
        if (sessionScans.length === 0) {
            alert("No hay datos para exportar.");
            return;
        }

        const ws = XLSX.utils.json_to_sheet(sessionScans);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Escaneados");
        XLSX.writeFile(wb, "Reporte_Escaneos.xlsx");
    });

    clearScansBtn.addEventListener('click', () => {
        if(confirm("¿Borrar lista actual?")) {
            sessionScans = [];
            updateTable();
        }
    });

});
