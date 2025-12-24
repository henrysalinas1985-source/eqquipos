document.addEventListener('DOMContentLoaded', () => {
    // === INDEXEDDB PARA IMÁGENES ===
    let imageDB = null;
    const DB_NAME = 'EquiposImageDB';
    const STORE_NAME = 'images';

    function initImageDB() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, 1);
            request.onerror = () => reject(request.error);
            request.onsuccess = () => {
                imageDB = request.result;
                resolve(imageDB);
            };
            request.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    db.createObjectStore(STORE_NAME, { keyPath: 'id' });
                }
            };
        });
    }

    function saveImageToDB(id, dataUrl) {
        return new Promise((resolve, reject) => {
            if (!imageDB) return reject('DB not initialized');
            const tx = imageDB.transaction(STORE_NAME, 'readwrite');
            const store = tx.objectStore(STORE_NAME);
            store.put({ id, dataUrl });
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    }

    function getImageFromDB(id) {
        return new Promise((resolve, reject) => {
            if (!imageDB) return resolve(null);
            const tx = imageDB.transaction(STORE_NAME, 'readonly');
            const store = tx.objectStore(STORE_NAME);
            const request = store.get(id);
            request.onsuccess = () => resolve(request.result?.dataUrl || null);
            request.onerror = () => resolve(null);
        });
    }

    function getAllImagesFromDB() {
        return new Promise((resolve, reject) => {
            if (!imageDB) return resolve([]);
            const tx = imageDB.transaction(STORE_NAME, 'readonly');
            const store = tx.objectStore(STORE_NAME);
            const request = store.getAll();
            request.onsuccess = () => resolve(request.result || []);
            request.onerror = () => resolve([]);
        });
    }

    // Inicializar DB al cargar
    initImageDB().then(() => console.log('ImageDB lista')).catch(console.error);

    // --- BACKUP DE IMÁGENES ---
    const exportImagesBtn = document.getElementById('exportImagesBtn');
    const importImagesInput = document.getElementById('importImagesInput');
    const backupStatus = document.getElementById('backupStatus');

    function showBackupStatus(msg, isError = false) {
        backupStatus.textContent = msg;
        backupStatus.style.color = isError ? '#ff6464' : '#00ff88';
        backupStatus.classList.remove('hidden');
        setTimeout(() => backupStatus.classList.add('hidden'), 4000);
    }

    exportImagesBtn.addEventListener('click', async () => {
        try {
            exportImagesBtn.disabled = true;
            exportImagesBtn.textContent = '⏳ Exportando...';
            
            const images = await getAllImagesFromDB();
            
            if (images.length === 0) {
                showBackupStatus('No hay imágenes para exportar', true);
                return;
            }

            const zip = new JSZip();
            
            images.forEach(img => {
                // Convertir dataUrl a blob
                const base64 = img.dataUrl.split(',')[1];
                zip.file(img.id, base64, { base64: true });
            });

            const content = await zip.generateAsync({ type: 'blob' });
            
            // Descargar
            const link = document.createElement('a');
            link.href = URL.createObjectURL(content);
            link.download = `backup_imagenes_${new Date().toISOString().slice(0,10)}.zip`;
            link.click();
            
            showBackupStatus(`✅ ${images.length} imágenes exportadas`);
            
        } catch (err) {
            console.error(err);
            showBackupStatus('Error al exportar', true);
        } finally {
            exportImagesBtn.disabled = false;
            exportImagesBtn.textContent = '⬇️ Exportar Imágenes';
        }
    });

    importImagesInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        try {
            showBackupStatus('⏳ Importando...');
            
            const zip = await JSZip.loadAsync(file);
            let count = 0;

            for (const [filename, zipEntry] of Object.entries(zip.files)) {
                if (!zipEntry.dir) {
                    const base64 = await zipEntry.async('base64');
                    const dataUrl = `data:image/jpeg;base64,${base64}`;
                    await saveImageToDB(filename, dataUrl);
                    count++;
                }
            }

            showBackupStatus(`✅ ${count} imágenes importadas`);
            
        } catch (err) {
            console.error(err);
            showBackupStatus('Error al importar', true);
        }
        
        importImagesInput.value = '';
    });

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
    const equipoNombre = document.getElementById('equipoNombre');
    const editLocationSelect = document.getElementById('editLocationSelect');
    const dateInput = document.getElementById('dateInput');
    const editObservaciones = document.getElementById('editObservaciones');
    const observacionesContainer = document.getElementById('observacionesContainer');
    const addObsBtn = document.getElementById('addObsBtn');
    const updateBtn = document.getElementById('updateBtn');
    const editImageContainer = document.getElementById('editImageContainer');
    const editImagePreview = document.getElementById('editImagePreview');
    const editImageName = document.getElementById('editImageName');
    const noImageMsg = document.getElementById('noImageMsg');
    const editImageInput = document.getElementById('editImageInput');
    const loadImageBtn = document.getElementById('loadImageBtn');
    let currentMatchIndex = -1;
    let currentImageRef = null;

    // Elementos de Registro
    const registerSerieBtn = document.getElementById('registerSerieBtn');
    const registerModal = document.getElementById('registerModal');
    const regSerieInput = document.getElementById('regSerieInput');
    const regLocationSelect = document.getElementById('regLocationSelect');
    const regObservaciones = document.getElementById('regObservaciones');
    const confirmRegBtn = document.getElementById('confirmRegBtn');
    const cancelRegBtn = document.getElementById('cancelRegBtn');
    const regFeedback = document.getElementById('regFeedback');

    // Elementos de cámara para registro
    const cameraContainer = document.getElementById('cameraContainer');
    const cameraVideo = document.getElementById('cameraVideo');
    const capturedImage = document.getElementById('capturedImage');
    const startCameraBtn = document.getElementById('startCameraBtn');
    const capturePhotoBtn = document.getElementById('capturePhotoBtn');
    const retakePhotoBtn = document.getElementById('retakePhotoBtn');
    const deletePhotoBtn = document.getElementById('deletPhotoBtn');
    let cameraStream = null;
    let capturedImageData = null;

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

    function renderTable(filterValue = '') {
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

        const serieKey = getColumnKey('serie');
        const normalize = s => String(s || '').trim().toUpperCase();
        const filterNorm = normalize(filterValue);

        globalDataRaw.forEach((row, index) => {
            // Filtrar por serie si hay filtro
            if (filterValue && serieKey) {
                const serieVal = normalize(row[serieKey]);
                if (!serieVal.includes(filterNorm)) return;
            }

            const tr = document.createElement('tr');
            tr.style.cursor = 'pointer';
            tr.addEventListener('click', () => openEditFromTable(index));
            
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

    // Abrir edición desde la tabla
    function openEditFromTable(index) {
        currentMatchIndex = index;
        const row = globalDataRaw[index];

        const idKey = globalHeaders[0];
        const serieKey = getColumnKey('serie');
        const equipoKey = getColumnKey('equipo');
        const obsKey = getColumnKey('observacion');
        const locKey = getColumnKey('ubicacion') || getColumnKey('tecnica');
        const imgKey = getColumnKey('imagen') || getColumnKey('foto');

        scanResult.textContent = `📝 Editando fila ${index + 2}`;
        scanResult.className = 'feedback success';
        scanResult.classList.remove('hidden');

        matchInfo.innerHTML = `<strong>ID:</strong> ${row[idKey] || 'N/A'} | <strong>Serie:</strong> ${serieKey ? row[serieKey] : 'N/A'}`;
        
        equipoNombre.textContent = equipoKey ? (row[equipoKey] || 'Sin nombre') : 'N/A';
        
        // Cargar todas las observaciones
        loadObservacionesForRow(row);
        
        // Cargar ubicación actual
        if (locKey && row[locKey]) {
            editLocationSelect.value = row[locKey];
        } else {
            editLocationSelect.value = '';
        }
        
        // Mostrar referencia de imagen si existe
        currentImageRef = null;
        editImageInput.value = '';
        if (imgKey && row[imgKey]) {
            currentImageRef = row[imgKey];
            editImageName.textContent = row[imgKey];
            editImageContainer.classList.remove('hidden');
            noImageMsg.classList.add('hidden');
            loadImageBtn.textContent = '📂 Cambiar Imagen';
            
            // Cargar imagen desde IndexedDB
            getImageFromDB(row[imgKey]).then(dataUrl => {
                if (dataUrl) {
                    editImagePreview.src = dataUrl;
                    editImagePreview.style.display = 'block';
                } else {
                    editImagePreview.style.display = 'none';
                    editImageName.textContent = row[imgKey] + ' (no encontrada en este dispositivo)';
                }
            });
        } else {
            editImageContainer.classList.add('hidden');
            editImagePreview.style.display = 'block';
            noImageMsg.classList.remove('hidden');
            loadImageBtn.textContent = '📂 Cargar Imagen';
        }
        
        const dateKey = getColumnKey('calibracion') || getColumnKey('fecha');
        if (dateKey) {
            dateInput.value = formatDateForInput(row[dateKey]);
            dateInput.dataset.targetKey = dateKey;
        }

        editPanel.classList.remove('hidden');
        
        // Scroll al panel de edición
        editPanel.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    // Evento filtro de tabla
    document.getElementById('filterSerieInput').addEventListener('input', function() {
        renderTable(this.value);
    });

    // --- CARGAR IMAGEN EN EDICIÓN ---
    loadImageBtn.addEventListener('click', () => {
        editImageInput.click();
    });

    // --- MÚLTIPLES OBSERVACIONES ---
    let obsFieldCount = 1;

    function getObsColumns() {
        // Buscar todas las columnas que contengan "observacion"
        const normalize = str => str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        return globalHeaders.filter(h => normalize(h).includes('observacion'));
    }

    function resetObservacionesUI() {
        // Limpiar campos adicionales
        observacionesContainer.innerHTML = `
            <textarea id="editObservaciones" placeholder="Observación principal..." rows="2" 
                style="width:100%; padding:14px; border:1px solid rgba(255,255,255,0.1); border-radius:10px; background:rgba(0,0,0,0.3); color:#fff; font-size:1rem; resize:vertical; outline:none; margin-bottom:8px;"></textarea>
        `;
        obsFieldCount = 1;
    }

    function loadObservacionesForRow(row) {
        resetObservacionesUI();
        const obsCols = getObsColumns();
        
        // Cargar primera observación
        const mainTextarea = document.getElementById('editObservaciones');
        if (obsCols.length > 0 && row[obsCols[0]]) {
            mainTextarea.value = row[obsCols[0]] || '';
        } else {
            mainTextarea.value = '';
        }

        // Cargar observaciones adicionales existentes
        for (let i = 1; i < obsCols.length; i++) {
            if (row[obsCols[i]]) {
                addObservacionField(obsCols[i], row[obsCols[i]]);
            }
        }
    }

    function addObservacionField(colName = null, value = '') {
        obsFieldCount++;
        const fieldId = `editObs_${obsFieldCount}`;
        const label = colName || `Observaciones ${obsFieldCount}`;
        
        const div = document.createElement('div');
        div.style.marginBottom = '8px';
        div.innerHTML = `
            <small style="color:#888; font-size:0.75rem;">${label}</small>
            <textarea id="${fieldId}" data-colname="${colName || ''}" placeholder="Nueva observación..." rows="2" 
                style="width:100%; padding:12px; border:1px solid rgba(0,217,255,0.3); border-radius:10px; background:rgba(0,0,0,0.3); color:#fff; font-size:0.95rem; resize:vertical; outline:none;">${value}</textarea>
        `;
        observacionesContainer.appendChild(div);
        return fieldId;
    }

    addObsBtn.addEventListener('click', () => {
        if (currentMatchIndex === -1) {
            alert('Primero selecciona un equipo');
            return;
        }

        // Contar columnas de observaciones existentes
        const obsCols = getObsColumns();
        let maxNum = 1;
        
        // Buscar el número más alto existente
        obsCols.forEach(col => {
            const match = col.match(/(\d+)$/);
            if (match) {
                const num = parseInt(match[1]);
                if (num > maxNum) maxNum = num;
            }
        });
        
        const newColName = `Observaciones ${maxNum + 1}`;
        
        // Agregar columna a headers si no existe
        if (!globalHeaders.includes(newColName)) {
            globalHeaders.push(newColName);
            // Agregar columna vacía a todos los registros
            globalDataRaw.forEach(row => {
                if (!(newColName in row)) row[newColName] = '';
            });
        }

        addObservacionField(newColName, '');
        
        // Scroll al nuevo campo
        observacionesContainer.lastChild.scrollIntoView({ behavior: 'smooth' });
    });

    editImageInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file && currentMatchIndex !== -1) {
            const reader = new FileReader();
            reader.onload = (ev) => {
                const dataUrl = ev.target.result;
                editImagePreview.src = dataUrl;
                editImagePreview.style.display = 'block';
                editImageName.textContent = file.name;
                editImageContainer.classList.remove('hidden');
                noImageMsg.classList.add('hidden');
                
                // Guardar en IndexedDB y actualizar referencia
                const serieKey = getColumnKey('serie');
                const imgKey = getColumnKey('imagen') || getColumnKey('foto');
                const row = globalDataRaw[currentMatchIndex];
                const serieVal = serieKey ? row[serieKey] : `equipo_${currentMatchIndex}`;
                const imgFilename = `${String(serieVal).replace(/[^a-zA-Z0-9]/g, '_')}.jpg`;
                
                if (imgKey) {
                    globalDataRaw[currentMatchIndex][imgKey] = imgFilename;
                }
                
                saveImageToDB(imgFilename, dataUrl)
                    .then(() => {
                        console.log('Imagen cargada y guardada:', imgFilename);
                        currentImageRef = imgFilename;
                        editImageName.textContent = imgFilename;
                    })
                    .catch(console.error);
            };
            reader.readAsDataURL(file);
        }
    });

    function populateLocations() {
        // Buscar columna que contenga "ubicacion" o "tecnica" o "location"
        const locKey = getColumnKey('ubicacion') || getColumnKey('tecnica') || getColumnKey('location');
        
        console.log("Headers disponibles:", globalHeaders);
        console.log("Columna ubicación encontrada:", locKey);
        
        if (!locKey) {
            console.warn("No se encontró columna de ubicación. Headers:", globalHeaders);
            regLocationSelect.innerHTML = '<option value="">-- No se encontró columna ubicación --</option>';
            editLocationSelect.innerHTML = '<option value="">-- No se encontró columna ubicación --</option>';
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

        const options = '<option value="">-- Seleccionar Ubicación --</option>' + 
            Array.from(locSet).sort().map(loc => `<option value="${loc}">${loc}</option>`).join('');
        
        regLocationSelect.innerHTML = options;
        editLocationSelect.innerHTML = options;
    }

    // --- REGISTRO SERIE ---
    registerSerieBtn.addEventListener('click', () => {
        if (registerSerieBtn.classList.contains('disabled')) return;
        if (globalDataRaw.length === 0) {
            alert("Primero carga un archivo Excel.");
            return;
        }
        resetCameraUI();
        regSerieInput.value = '';
        regLocationSelect.value = '';
        regObservaciones.value = '';
        regFeedback.classList.add('hidden');
        registerModal.classList.remove('hidden');
        regSerieInput.focus();
    });

    cancelRegBtn.addEventListener('click', () => {
        stopCamera();
        registerModal.classList.add('hidden');
    });

    // --- CÁMARA PARA FOTO ---
    function resetCameraUI() {
        stopCamera();
        cameraContainer.classList.add('hidden');
        capturedImage.classList.add('hidden');
        startCameraBtn.classList.remove('hidden');
        capturePhotoBtn.classList.add('hidden');
        retakePhotoBtn.classList.add('hidden');
        deletePhotoBtn.classList.add('hidden');
        capturedImageData = null;
    }

    function stopCamera() {
        if (cameraStream) {
            cameraStream.getTracks().forEach(track => track.stop());
            cameraStream = null;
        }
    }

    startCameraBtn.addEventListener('click', async () => {
        startCameraBtn.textContent = '⏳ Cargando...';
        startCameraBtn.disabled = true;
        
        try {
            cameraStream = await navigator.mediaDevices.getUserMedia({
                video: { 
                    facingMode: "environment",
                    width: { ideal: 1280 },
                    height: { ideal: 720 }
                }
            });
            
            cameraVideo.srcObject = cameraStream;
            cameraContainer.classList.remove('hidden');
            startCameraBtn.classList.add('hidden');
            capturePhotoBtn.classList.remove('hidden');
            
        } catch (e) {
            alert("No se pudo acceder a la cámara: " + e.message);
        }
        
        startCameraBtn.textContent = '📷 Abrir Cámara';
        startCameraBtn.disabled = false;
    });

    capturePhotoBtn.addEventListener('click', () => {
        const canvas = document.createElement('canvas');
        canvas.width = cameraVideo.videoWidth;
        canvas.height = cameraVideo.videoHeight;
        canvas.getContext('2d').drawImage(cameraVideo, 0, 0);
        
        capturedImageData = canvas.toDataURL('image/jpeg', 0.7);
        capturedImage.src = capturedImageData;
        
        stopCamera();
        cameraContainer.classList.add('hidden');
        capturedImage.classList.remove('hidden');
        capturePhotoBtn.classList.add('hidden');
        retakePhotoBtn.classList.remove('hidden');
        deletePhotoBtn.classList.remove('hidden');
    });

    retakePhotoBtn.addEventListener('click', () => {
        capturedImage.classList.add('hidden');
        retakePhotoBtn.classList.add('hidden');
        deletePhotoBtn.classList.add('hidden');
        capturedImageData = null;
        startCameraBtn.click();
    });

    deletePhotoBtn.addEventListener('click', () => {
        capturedImage.classList.add('hidden');
        retakePhotoBtn.classList.add('hidden');
        deletePhotoBtn.classList.add('hidden');
        startCameraBtn.classList.remove('hidden');
        capturedImageData = null;
    });

    function downloadImage(dataUrl, filename) {
        const link = document.createElement('a');
        link.href = dataUrl;
        link.download = filename;
        link.click();
    }

    // Filtro en tiempo real mientras escribe la serie
    let foundExistingIndex = -1; // Para guardar el índice si existe
    
    regSerieInput.addEventListener('input', () => {
        const searchVal = regSerieInput.value.trim().toUpperCase();
        foundExistingIndex = -1;
        
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
            foundExistingIndex = exactMatch;
            const row = globalDataRaw[exactMatch];
            const locKey = getColumnKey('ubicacion') || getColumnKey('tecnica');
            const equipoKey = getColumnKey('equipo');
            const obsKey = getColumnKey('observacion');
            
            // Cargar datos existentes
            if (locKey && row[locKey]) {
                // Seleccionar la ubicación existente
                regLocationSelect.value = row[locKey];
            }
            if (obsKey) {
                regObservaciones.value = row[obsKey] || '';
            }
            
            regFeedback.innerHTML = `⚠️ Serie existe (fila ${exactMatch + 2})<br>Equipo: <strong>${equipoKey ? row[equipoKey] : 'N/A'}</strong><br>Puedes actualizar las observaciones`;
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
        const obsVal = regObservaciones.value.trim();

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
        const obsKey = getColumnKey('observacion');

        console.log("Columnas encontradas:", { serieKey, locKey, obsKey });
        console.log("Headers:", globalHeaders);

        if (!serieKey) {
            alert("No se encontró columna 'Serie' en el Excel. Columnas disponibles: " + globalHeaders.join(', '));
            return;
        }

        // Verificar si existe
        const normalize = s => String(s || '').trim().toUpperCase();
        const existsIndex = globalDataRaw.findIndex(row => normalize(row[serieKey]) === serieVal);

        if (existsIndex !== -1) {
            // EXISTE - Actualizar observaciones e imagen
            if (obsKey) {
                globalDataRaw[existsIndex][obsKey] = obsVal;
            }
            
            // Si hay imagen nueva, descargarla y actualizar referencia
            const imgKey = getColumnKey('imagen') || getColumnKey('foto');
            if (capturedImageData) {
                const imgFilename = `${serieVal.replace(/[^a-zA-Z0-9]/g, '_')}.jpg`;
                if (imgKey) globalDataRaw[existsIndex][imgKey] = imgFilename;
                
                // Guardar en IndexedDB
                saveImageToDB(imgFilename, capturedImageData)
                    .then(() => console.log('Imagen actualizada en DB:', imgFilename))
                    .catch(console.error);
                
                downloadImage(capturedImageData, imgFilename);
            }
            
            renderTable();
            regFeedback.textContent = `✅ Actualizado para serie "${serieVal}"` + (capturedImageData ? ' (imagen descargada)' : '');
            regFeedback.className = 'feedback success';
            regFeedback.classList.remove('hidden');
            
            setTimeout(() => {
                resetCameraUI();
                registerModal.classList.add('hidden');
            }, 1500);
            return;
        }

        // NO EXISTE - Crear nueva fila
        const newRow = {};
        globalHeaders.forEach(h => newRow[h] = "");
        newRow[serieKey] = serieVal;
        if (locKey) newRow[locKey] = locVal;
        if (obsKey) newRow[obsKey] = obsVal;

        const dateKey = getColumnKey('calibracion') || getColumnKey('fecha');
        if (dateKey) newRow[dateKey] = new Date();

        // Guardar referencia de imagen si se capturó
        const imgKey = getColumnKey('imagen') || getColumnKey('foto');
        if (capturedImageData) {
            const imgFilename = `${serieVal.replace(/[^a-zA-Z0-9]/g, '_')}.jpg`;
            if (imgKey) newRow[imgKey] = imgFilename;
            
            // Guardar en IndexedDB
            saveImageToDB(imgFilename, capturedImageData)
                .then(() => console.log('Imagen guardada en DB:', imgFilename))
                .catch(console.error);
            
            // También descargar
            downloadImage(capturedImageData, imgFilename);
        }

        globalDataRaw.push(newRow);
        renderTable();

        regFeedback.textContent = `✅ Serie "${serieVal}" registrada` + (capturedImageData ? ' (imagen descargada)' : '');
        regFeedback.className = 'feedback success';
        regFeedback.classList.remove('hidden');

        setTimeout(() => {
            resetCameraUI();
            registerModal.classList.add('hidden');
        }, 1500);
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

            const idKey = globalHeaders[0];
            const serieKey = getColumnKey('serie');
            const equipoKey = getColumnKey('equipo');
            const obsKey = getColumnKey('observacion');
            const locKey = getColumnKey('ubicacion') || getColumnKey('tecnica');
            const imgKey = getColumnKey('imagen') || getColumnKey('foto');

            matchInfo.innerHTML = `<strong>ID:</strong> ${row[idKey] || 'N/A'} | <strong>Serie:</strong> ${serieKey ? row[serieKey] : 'N/A'}`;
            
            // Mostrar nombre del equipo
            equipoNombre.textContent = equipoKey ? (row[equipoKey] || 'Sin nombre') : 'N/A';
            
            // Cargar todas las observaciones
            loadObservacionesForRow(row);
            
            // Cargar ubicación actual
            if (locKey && row[locKey]) {
                editLocationSelect.value = row[locKey];
            } else {
                editLocationSelect.value = '';
            }
            
            // Mostrar referencia de imagen si existe
            currentImageRef = null;
            editImageInput.value = '';
            if (imgKey && row[imgKey]) {
                currentImageRef = row[imgKey];
                editImageName.textContent = row[imgKey];
                editImageContainer.classList.remove('hidden');
                noImageMsg.classList.add('hidden');
                loadImageBtn.textContent = '📂 Cambiar Imagen';
                
                // Cargar imagen desde IndexedDB
                getImageFromDB(row[imgKey]).then(dataUrl => {
                    if (dataUrl) {
                        editImagePreview.src = dataUrl;
                        editImagePreview.style.display = 'block';
                    } else {
                        editImagePreview.style.display = 'none';
                        editImageName.textContent = row[imgKey] + ' (no encontrada en este dispositivo)';
                    }
                });
            } else {
                editImageContainer.classList.add('hidden');
                editImagePreview.style.display = 'block';
                noImageMsg.classList.remove('hidden');
                loadImageBtn.textContent = '📂 Cargar Imagen';
            }
            
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

            resetCameraUI();
            regSerieInput.value = scannedValue;
            regLocationSelect.value = '';
            regObservaciones.value = '';
            regFeedback.classList.add('hidden');
            registerModal.classList.remove('hidden');
        }
    }

    // --- ACTUALIZAR FECHA, UBICACIÓN Y OBSERVACIONES ---
    updateBtn.addEventListener('click', () => {
        if (currentMatchIndex === -1) return;

        const newDate = dateInput.value;
        const newLoc = editLocationSelect.value;

        const targetKey = dateInput.dataset.targetKey;
        const locKey = getColumnKey('ubicacion') || getColumnKey('tecnica');

        if (targetKey && newDate) {
            const [y, m, d] = newDate.split('-').map(Number);
            globalDataRaw[currentMatchIndex][targetKey] = new Date(y, m - 1, d);
        }

        if (locKey && newLoc) {
            globalDataRaw[currentMatchIndex][locKey] = newLoc;
        }

        // Guardar todas las observaciones
        const obsCols = getObsColumns();
        
        // Primera observación
        const mainObs = document.getElementById('editObservaciones');
        if (mainObs && obsCols.length > 0) {
            globalDataRaw[currentMatchIndex][obsCols[0]] = mainObs.value.trim();
        }

        // Observaciones adicionales
        const additionalTextareas = observacionesContainer.querySelectorAll('textarea[id^="editObs_"]');
        additionalTextareas.forEach(textarea => {
            const colName = textarea.dataset.colname;
            if (colName && globalHeaders.includes(colName)) {
                globalDataRaw[currentMatchIndex][colName] = textarea.value.trim();
            }
        });

        alert(`✅ Actualizado (fila ${currentMatchIndex + 2})`);
        editPanel.classList.add('hidden');
        scanResult.classList.add('hidden');
        
        // Mantener el filtro actual
        const filterInput = document.getElementById('filterSerieInput');
        renderTable(filterInput ? filterInput.value : '');
    });
});
