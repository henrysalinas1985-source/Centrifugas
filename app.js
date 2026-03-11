document.addEventListener('DOMContentLoaded', () => {
    // === ESTADO ===
    let db = null;
    let allSheetsData = {};
    let currentClinic = '';
    let calibrationDates = {};
    let instrumentsBank = [];
    let savedTemplates = [];
    let selectedSerieForEdit = null;

    // === SCHEMA EXACTO DEL EXCEL 2025-MP CEN ===
    // 8.1 Test de Inspección y Funcionalidad — 14 ítems, filas 24-37
    const SCHEMA_81 = [
        { code: '8.1.1', label: 'Chasis', row: 24 },
        { code: '8.1.2', label: 'Montajes y Apoyos', row: 25 },
        { code: '8.1.3', label: 'Enchufe de Red y Base de Enchufe', row: 26 },
        { code: '8.1.4', label: 'Cable de Red', row: 27 },
        { code: '8.1.5', label: 'Amarres contra tirones', row: 28 },
        { code: '8.1.6', label: 'Interruptores y Fusibles', row: 29 },
        { code: '8.1.7', label: 'Controles y Teclas', row: 30 },
        { code: '8.1.8', label: 'Motor, Rotor y Bomba', row: 31 },
        { code: '8.1.9', label: 'Indicadores y Displays', row: 32 },
        { code: '8.1.10', label: 'Alarmas y Dispositivos para cortar la corriente', row: 33 },
        { code: '8.1.11', label: 'Señales audibles', row: 34 },
        { code: '8.1.12', label: 'Etiquetado', row: 35 },
        { code: '8.1.13', label: 'Accesorios', row: 36 },
        { code: '8.1.14', label: 'Frenos', row: 37 },
    ];

    // 8.2 Test Cuantitativos — Seteado/Medido/P-F
    const SCHEMA_82 = [
        {
            code: '8.2.1', label: 'Exactitud de la Temperatura (Valor Aceptable +/- 3°C)',
            unit: '°C', readings: 1, row: 40,
        },
        {
            code: '8.2.2', label: 'Exactitud del Tiempo (Valor Aceptable +/- 10 %)',
            unit: 'segundos', readings: 1, row: 43,
        },
        {
            code: '8.2.3', label: 'Exactitud de la Velocidad (Valor Aceptable +/- 10 %)',
            unit: 'rpm', readings: 2, row: 45,
            readingLabels: ['2000 RPM', '3500 RPM'],
        },
    ];

    // 8.3 Mantenimiento Preventivo — 6 ítems
    const SCHEMA_83 = [
        { code: '8.3.1', label: 'Limpieza Exterior', row: 51 },
        { code: '8.3.2', label: 'Lubricación', row: 52 },
        { code: '8.3.3', label: 'Sustitución de escobillas, frenos, juntas de estanqueidad y de cierre', row: 53 },
        { code: '8.3.4', label: 'Grupo de frío', row: 54 },
        { code: '8.3.5', label: 'Placas electrónicas y conectores', row: 55 },
        { code: '8.3.6', label: 'Test de Seguridad Eléctrica', row: 56 },
    ];

    // 8.4 Test de Aceptación — 1 ítem
    const SCHEMA_84 = [
        { code: '8.4.1', label: 'Seguridad de puertas', row: 59 },
    ];

    // Estado de Valoración — 3 ítems
    const EVALUATION_SCHEMA = [
        { label: 'Inspección superada, el equipo es apto para el uso', row: 15 },
        { label: 'El equipo ha necesitado reparación', row: 16 },
        { label: 'El equipo no está reparado. No se puede usar', row: 17 },
    ];

    const DB_NAME = 'CalibracionesDB_Centrifugas_v1';
    const DB_VERSION = 1;

    // === UTILIDADES ===
    function escapeHtml(str) {
        const div = document.createElement('div');
        div.appendChild(document.createTextNode(str));
        return div.innerHTML;
    }

    function safeParseFloat(val) {
        if (val === '' || val === null || val === undefined) return null;
        const n = parseFloat(val);
        return isNaN(n) ? null : n;
    }

    // DOM
    const fileInput = document.getElementById('fileInput');
    const fileLabel = document.getElementById('fileLabel');
    const mainContent = document.getElementById('mainContent');
    const sheetSelector = document.getElementById('sheetSelector');
    const serieFilter = document.getElementById('serieFilter');
    const equiposTableBody = document.getElementById('equiposTableBody');
    const editModal = document.getElementById('editModal');
    const calibDateInput = document.getElementById('calibDateInput');
    const ordenMInput = document.getElementById('ordenMInput');
    const technicianInput = document.getElementById('technicianInput');
    const buildingInput = document.getElementById('buildingInput');
    const sectorInput = document.getElementById('sectorInput');
    const locationInput = document.getElementById('locationInput');
    const commentsInput = document.getElementById('commentsInput');
    const equipmentNameInput = document.getElementById('equipmentNameInput');
    const modalSerieInput = document.getElementById('modalSerieInput');
    const modelInput = document.getElementById('modelInput');
    const brandInput = document.getElementById('brandInput');
    const addInstrumentBtn = document.getElementById('addInstrumentBtn');
    const instrumentsContainer = document.getElementById('instrumentsContainer');
    const certFileInput = document.getElementById('certFileInput');
    const certStatus = document.getElementById('certStatus');
    const templateSelector = document.getElementById('templateSelector');
    const templateNameInput = document.getElementById('templateNameInput');
    const saveNewTemplateBtn = document.getElementById('saveNewTemplateBtn');
    const saveTemplateRow = document.getElementById('saveTemplateRow');
    const saveCalibBtn = document.getElementById('saveCalibBtn');
    const totalEquiposEl = document.getElementById('totalEquipos').querySelector('.val');
    const cercaVencerEl = document.getElementById('cercaVencer').querySelector('.val');
    const vencidosEl = document.getElementById('vencidos').querySelector('.val');

    // === INIT ===
    async function init() {
        try {
            await initDB();
            await loadSavedData();
            setupEventListeners();
            loadTemplates();
        } catch (err) {
            console.error('Init error:', err);
            alert('Error al iniciar: ' + err.message);
        }
    }

    // === INDEXEDDB ===
    function initDB() {
        return new Promise((resolve, reject) => {
            const req = indexedDB.open(DB_NAME, DB_VERSION);
            req.onupgradeneeded = e => {
                const d = e.target.result;
                if (!d.objectStoreNames.contains('calibrations')) d.createObjectStore('calibrations', { keyPath: 'serie' });
                if (!d.objectStoreNames.contains('appData')) d.createObjectStore('appData', { keyPath: 'id' });
                if (!d.objectStoreNames.contains('templates')) d.createObjectStore('templates', { keyPath: 'id', autoIncrement: true });
            };
            req.onsuccess = e => { db = e.target.result; resolve(); };
            req.onerror = e => reject(e.target.error);
        });
    }

    async function storeCalibration(data) {
        const tx = db.transaction('calibrations', 'readwrite');
        const store = tx.objectStore('calibrations');
        if (!data.certificate) {
            const existing = await new Promise(r => { const q = store.get(data.serie); q.onsuccess = () => r(q.result); });
            if (existing && existing.certificate) {
                data.certificate = existing.certificate;
                data.certName = existing.certName;
            }
        }
        store.put(data);
    }

    function getAllCalibrations() {
        return new Promise(resolve => {
            if (!db) { resolve({}); return; }
            const map = {}, tx = db.transaction('calibrations', 'readonly');
            tx.objectStore('calibrations').openCursor().onsuccess = e => {
                const cur = e.target.result;
                if (cur) { map[cur.key] = cur.value; cur.continue(); }
                else { calibrationDates = map; updateInstrumentsBank(); resolve(map); }
            };
        });
    }

    function updateInstrumentsBank() {
        const uniq = new Map();
        Object.values(calibrationDates).forEach(c => {
            (c.instruments || []).forEach(i => {
                if (i.name && !uniq.has(i.name.toUpperCase())) uniq.set(i.name.toUpperCase(), i);
            });
        });
        instrumentsBank = Array.from(uniq.values());
        const dl = document.getElementById('instrumentsHistory');
        if (dl) { dl.innerHTML = ''; instrumentsBank.forEach(i => { const o = document.createElement('option'); o.value = i.name; dl.appendChild(o); }); }
    }

    // === EXCEL LOADING ===
    fileInput.addEventListener('change', e => {
        const file = e.target.files[0]; if (!file) return;
        const reader = new FileReader();
        reader.onload = ev => processWorkbook(XLSX.read(new Uint8Array(ev.target.result), { type: 'array' }), file.name);
        reader.readAsArrayBuffer(file);
    });

    async function processWorkbook(wb, filename) {
        allSheetsData = {};
        sheetSelector.innerHTML = '';
        wb.SheetNames.forEach(name => {
            allSheetsData[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' }).filter(r => Object.values(r).some(v => v !== ''));
            const opt = document.createElement('option'); opt.value = opt.textContent = name;
            sheetSelector.appendChild(opt);
        });
        currentClinic = wb.SheetNames[0];
        fileLabel.textContent = `✅ ${filename}`;
        db.transaction('appData', 'readwrite').objectStore('appData').put({ id: 'lastExcel', filename, allSheetsData, sheetNames: wb.SheetNames, currentClinic });
        mainContent.classList.remove('hidden');
        document.getElementById('configActions').classList.remove('hidden');
        renderTable();
    }

    async function loadSavedData() {
        const tx = db.transaction('appData', 'readonly');
        const last = await new Promise(r => { const q = tx.objectStore('appData').get('lastExcel'); q.onsuccess = () => r(q.result); });
        if (!last) return;
        allSheetsData = last.allSheetsData; currentClinic = last.currentClinic;
        sheetSelector.innerHTML = '';
        last.sheetNames.forEach(n => { const o = document.createElement('option'); o.value = o.textContent = n; if (n === currentClinic) o.selected = true; sheetSelector.appendChild(o); });
        fileLabel.textContent = `✅ ${last.filename} (Recuperado)`;
        mainContent.classList.remove('hidden');
        document.getElementById('configActions').classList.remove('hidden');
        renderTable();
    }

    document.getElementById('clearDataBtn').addEventListener('click', () => {
        if (confirm('¿Borrar datos cargados?')) { db.transaction('appData', 'readwrite').objectStore('appData').delete('lastExcel'); location.reload(); }
    });

    // === BACKUP ===
    function blobToBase64(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    }

    function base64ToBlob(dataUrl) {
        const [header, data] = dataUrl.split(',');
        const mime = header.match(/:(.*?);/)[1];
        const bytes = atob(data);
        const arr = new Uint8Array(bytes.length);
        for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
        return new Blob([arr], { type: mime });
    }

    document.getElementById('exportBackupBtn').addEventListener('click', async () => {
        try {
            const backup = { version: 1, exportDate: new Date().toISOString(), calibrations: {}, templates: [] };
            const calTx = db.transaction('calibrations', 'readonly');
            const allCals = await new Promise(r => { const q = calTx.objectStore('calibrations').getAll(); q.onsuccess = () => r(q.result); });
            for (const cal of allCals) {
                const entry = { ...cal };
                if (entry.certificate instanceof Blob) { entry._certBase64 = await blobToBase64(entry.certificate); delete entry.certificate; }
                backup.calibrations[cal.serie] = entry;
            }
            const tmplTx = db.transaction('templates', 'readonly');
            const allTmpls = await new Promise(r => { const q = tmplTx.objectStore('templates').getAll(); q.onsuccess = () => r(q.result); });
            for (const tmpl of allTmpls) {
                const entry = { ...tmpl };
                if (entry.blob instanceof Blob) { entry._blobBase64 = await blobToBase64(entry.blob); delete entry.blob; }
                backup.templates.push(entry);
            }
            const json = JSON.stringify(backup, null, 2);
            const blob = new Blob([json], { type: 'application/json' });
            const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
            a.download = `backup_centrifugas_${new Date().toISOString().slice(0, 10)}.json`;
            a.click(); URL.revokeObjectURL(a.href);
            alert('✅ Backup exportado correctamente.');
        } catch (err) { console.error(err); alert('Error al exportar: ' + err.message); }
    });

    document.getElementById('importBackupBtn').addEventListener('click', () => document.getElementById('importBackupFile').click());

    document.getElementById('importBackupFile').addEventListener('change', async e => {
        const file = e.target.files[0]; if (!file) return;
        try {
            const backup = JSON.parse(await file.text());
            if (!backup.version || !backup.calibrations) throw new Error('Formato inválido');
            const count = { cals: 0, tmpls: 0 };
            const calStore = db.transaction('calibrations', 'readwrite').objectStore('calibrations');
            for (const [serie, cal] of Object.entries(backup.calibrations)) {
                const entry = { ...cal, serie };
                if (entry._certBase64) { entry.certificate = base64ToBlob(entry._certBase64); delete entry._certBase64; }
                calStore.put(entry); count.cals++;
            }
            if (backup.templates?.length > 0) {
                const tmplStore = db.transaction('templates', 'readwrite').objectStore('templates');
                for (const tmpl of backup.templates) {
                    const entry = { ...tmpl }; delete entry.id;
                    if (entry._blobBase64) { entry.blob = base64ToBlob(entry._blobBase64); delete entry._blobBase64; }
                    tmplStore.add(entry); count.tmpls++;
                }
            }
            alert(`✅ Importado: ${count.cals} calibraciones, ${count.tmpls} plantillas.`);
            location.reload();
        } catch (err) { console.error(err); alert('Error al importar: ' + err.message); }
        e.target.value = '';
    });

    // === TABLA ===
    async function renderTable() {
        if (!currentClinic || !allSheetsData[currentClinic]) return;
        await getAllCalibrations();
        const rows = allSheetsData[currentClinic];
        const search = serieFilter.value.trim().toUpperCase();
        equiposTableBody.innerHTML = '';
        let stats = { total: 0, warn: 0, danger: 0 };

        rows.forEach(row => {
            const keys = Object.keys(row);
            const serieKey = keys.find(k => k.toLowerCase().includes('serie') || k.toLowerCase().includes('n°') || k.toLowerCase().includes('sensor'));
            const nombreKey = keys.find(k => k.toLowerCase().includes('equipo') || k.toLowerCase().includes('nombre') || k.toLowerCase().includes('ubicacion') || k.toLowerCase().includes('ubicación'));
            const serie = String(row[serieKey] || '').toUpperCase().trim();
            if (!serie || serie === '') return;
            if (search && !serie.includes(search)) return;

            stats.total++;
            const cal = calibrationDates[serie] || null;
            const status = getStatus(cal?.date);
            if (status.class === 'status-warning') stats.warn++;
            if (status.class === 'status-danger') stats.danger++;

            const displayName = cal?.editedName || (nombreKey ? row[nombreKey] : 'N/A');
            const displaySerie = cal?.editedSerie || serie;
            const safeSerie = escapeHtml(serie);

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${escapeHtml(String(displayName))}</td>
                <td>${escapeHtml(String(displaySerie))}</td>
                <td>${cal ? formatDate(cal.date) : '<span style="color:#aaa">Sin registrar</span>'}</td>
                <td>${escapeHtml(cal?.technician || '-')}</td>
                <td>${cal?.certificate ? `<button class="btn btn-small" data-action="viewCert" data-serie="${safeSerie}">📄</button>` : '-'}</td>
                <td><span class="status-badge ${status.class}">${escapeHtml(status.text)}</span></td>
                <td><button class="btn btn-secondary btn-small" data-action="openEdit" data-serie="${safeSerie}">📅 Registrar</button></td>
            `;
            equiposTableBody.appendChild(tr);
        });

        totalEquiposEl.textContent = stats.total;
        cercaVencerEl.textContent = stats.warn;
        vencidosEl.textContent = stats.danger;
    }

    function getStatus(dateStr) {
        if (!dateStr) return { text: 'Pendiente', class: '' };
        const next = new Date(dateStr); next.setFullYear(next.getFullYear() + 1);
        const diff = Math.ceil((next - new Date()) / 86400000);
        if (diff < 0) return { text: 'Vencido', class: 'status-danger' };
        if (diff <= 30) return { text: `Vence ${diff}d`, class: 'status-warning' };
        return { text: 'Vigente', class: 'status-ok' };
    }

    // === INSTRUMENTAL ===
    function createInstrumentRow(data = {}) {
        const div = document.createElement('div'); div.className = 'instrument-item';
        const dateVal = formatDateForInput(data.date);
        div.innerHTML = `
            <button type="button" class="remove-instrument">×</button>
            <div class="field-group full-width"><label>Nombre del Instrumental</label><input type="text" class="inst-name" list="instrumentsHistory" value="${escapeHtml(data.name || '')}"></div>
            <div class="field-group"><label>Marca</label><input type="text" class="inst-brand" value="${escapeHtml(data.brand || '')}"></div>
            <div class="field-group"><label>Modelo</label><input type="text" class="inst-model" value="${escapeHtml(data.model || '')}"></div>
            <div class="field-group"><label>N° Serie</label><input type="text" class="inst-serie" value="${escapeHtml(data.serie || '')}"></div>
            <div class="field-group"><label>Últ. Calibración</label><input type="date" class="inst-date" value="${dateVal}"></div>
        `;
        div.querySelector('.remove-instrument').onclick = () => div.remove();
        instrumentsContainer.appendChild(div);
    }
    addInstrumentBtn.onclick = () => createInstrumentRow();

    function getInstrumentsData() {
        return Array.from(instrumentsContainer.querySelectorAll('.instrument-item')).map(div => ({
            name: div.querySelector('.inst-name').value,
            brand: div.querySelector('.inst-brand').value,
            model: div.querySelector('.inst-model').value,
            serie: div.querySelector('.inst-serie').value,
            date: div.querySelector('.inst-date').value,
        }));
    }

    // === INSPECCIÓN UI ===
    function renderInspectionPoints(saved = {}) {
        const container = document.getElementById('inspectionPointsContainer');
        container.innerHTML = '';

        // ── 8.1 Test de Inspección y Funcionalidad (14 ítems) ──
        addSectionHeader(container, '8.1 Test de Inspección y Funcionalidad');
        addSubLabel(container, 'Marcar P (Pasó) o F (Falló) según corresponda');
        SCHEMA_81.forEach(item => {
            addChoiceRow(container, item, saved, '81');
        });

        // ── 8.2 Test Cuantitativos ──
        addSectionHeader(container, '8.2 Test Cuantitativos');
        addSubLabel(container, 'Ingresar valores Seteado y Medido, marcar P o F');

        SCHEMA_82.forEach(test => {
            const savedTest = saved[`_quant_${test.code}`] || {};
            const testDiv = document.createElement('div');
            testDiv.className = 'quant-test-block';
            testDiv.dataset.code = test.code;
            testDiv.style.cssText = 'border:1px solid #3a3a5c;border-radius:8px;padding:12px;margin-bottom:12px;background:#1a1a2e;';

            let readingsHtml = '';
            for (let i = 0; i < test.readings; i++) {
                const sVal = escapeHtml(savedTest[`set${i}`] || '');
                const mVal = escapeHtml(savedTest[`med${i}`] || '');
                const pf = savedTest[`pf${i}`] || 'na';
                const rLabel = test.readingLabels ? test.readingLabels[i] : '';
                const suffix = rLabel ? ` (${rLabel})` : (test.readings > 1 ? ` ${i + 1}` : '');
                readingsHtml += `
                    <div style="display:grid;grid-template-columns:1fr 1fr auto;gap:6px;margin-bottom:4px;background:#2a2a4a;padding:6px;border-radius:4px;align-items:center;">
                        <div class="field-group"><label>Seteado${suffix}</label><input type="number" step="any" class="q-set" data-idx="${i}" value="${sVal}"></div>
                        <div class="field-group"><label>Medido${suffix}</label><input type="number" step="any" class="q-med" data-idx="${i}" value="${mVal}"></div>
                        <div class="inspection-options" data-idx="${i}" style="display:flex;gap:4px;min-width:80px;">
                            <div class="inspection-opt ${pf === 'P' ? 'selected' : ''}" data-val="P" style="padding:4px 8px;cursor:pointer;border-radius:4px;">P</div>
                            <div class="inspection-opt ${pf === 'F' ? 'selected' : ''}" data-val="F" style="padding:4px 8px;cursor:pointer;border-radius:4px;background:${pf === 'F' ? '#e53e3e' : ''}">F</div>
                            <div class="inspection-opt ${pf === 'na' ? 'selected' : ''}" data-val="na" style="padding:4px 8px;cursor:pointer;border-radius:4px;">NA</div>
                        </div>
                    </div>`;
            }

            testDiv.innerHTML = `
                <div style="margin-bottom:8px;">
                    <strong style="color:#7c83fd;">${test.code}</strong> ${test.label}
                    <span style="color:#aaa;font-size:0.8em;margin-left:8px;">Unidad: ${test.unit}</span>
                </div>
                ${readingsHtml}
            `;

            // Wire P/F/NA clicks
            testDiv.querySelectorAll('.inspection-options').forEach(group => {
                group.querySelectorAll('.inspection-opt').forEach(opt => {
                    opt.onclick = () => {
                        group.querySelectorAll('.inspection-opt').forEach(o => {
                            o.classList.remove('selected');
                            if (o.dataset.val === 'F') o.style.background = '';
                        });
                        opt.classList.add('selected');
                        if (opt.dataset.val === 'F') opt.style.background = '#e53e3e';
                    };
                });
            });

            container.appendChild(testDiv);
        });

        // ── 8.3 Mantenimiento Preventivo (6 ítems) ──
        addSectionHeader(container, '8.3 Mantenimiento Preventivo');
        addSubLabel(container, 'Marcar P (Pasó) o F (Falló) según corresponda');
        SCHEMA_83.forEach(item => {
            addChoiceRow(container, item, saved, '83');
        });

        // ── 8.4 Test de Aceptación (1 ítem) ──
        addSectionHeader(container, '8.4 Test de Aceptación');
        addSubLabel(container, 'Marcar P (Pasó) o F (Falló) según corresponda');
        SCHEMA_84.forEach(item => {
            addChoiceRow(container, item, saved, '84');
        });
    }

    function addChoiceRow(container, item, saved, section) {
        const saved_val = saved[`${section}_${item.label}`] || 'na';
        const rowEl = document.createElement('div'); rowEl.className = 'inspection-row';
        rowEl.innerHTML = `
            <div class="inspection-label"><strong>${item.code}</strong> ${item.label}</div>
            <div class="inspection-options" data-label="${section}_${item.label}" data-type="choice" data-row="${item.row}">
                <div class="inspection-opt ${saved_val === 'P' ? 'selected' : ''}" data-val="P">P</div>
                <div class="inspection-opt ${saved_val === 'F' ? 'selected' : ''}" data-val="F" style="background:${saved_val === 'F' ? '#e53e3e' : ''}">F</div>
                <div class="inspection-opt ${saved_val === 'na' ? 'selected' : ''}" data-val="na">NA</div>
            </div>`;
        wireChoiceOpts(rowEl);
        container.appendChild(rowEl);
    }

    function addSectionHeader(container, text) {
        const h = document.createElement('div'); h.className = 'inspection-category'; h.textContent = text;
        container.appendChild(h);
    }
    function addSubLabel(container, text) {
        const p = document.createElement('p'); p.style.cssText = 'font-size:0.78em;color:#aaa;margin:2px 0 8px 0;'; p.textContent = text;
        container.appendChild(p);
    }

    function wireChoiceOpts(rowEl) {
        rowEl.querySelectorAll('.inspection-opt').forEach(opt => {
            opt.onclick = () => {
                rowEl.querySelectorAll('.inspection-opt').forEach(o => {
                    o.classList.remove('selected');
                    if (o.dataset.val === 'F') o.style.background = '';
                });
                opt.classList.add('selected');
                if (opt.dataset.val === 'F') opt.style.background = '#e53e3e';
            };
        });
    }

    function getInspectionsData() {
        const data = {};
        // Recoger opciones P/F/NA de 8.1, 8.3, 8.4
        document.querySelectorAll('.inspection-options[data-type="choice"]').forEach(g => {
            const label = g.dataset.label;
            const sel = g.querySelector('.inspection-opt.selected');
            data[label] = sel ? sel.dataset.val : 'na';
        });

        // Recoger datos cuantitativos de 8.2
        document.querySelectorAll('.quant-test-block').forEach(block => {
            const code = block.dataset.code;
            const testData = {};
            block.querySelectorAll('.q-set').forEach(inp => { testData[`set${inp.dataset.idx}`] = inp.value; });
            block.querySelectorAll('.q-med').forEach(inp => { testData[`med${inp.dataset.idx}`] = inp.value; });
            block.querySelectorAll('.inspection-options').forEach(g => {
                const idx = g.dataset.idx;
                const sel = g.querySelector('.inspection-opt.selected');
                testData[`pf${idx}`] = sel ? sel.dataset.val : 'na';
            });
            data[`_quant_${code}`] = testData;
        });

        return data;
    }

    function getEvaluationsData() {
        const data = {};
        document.querySelectorAll('#evaluationStatusContainer .inspection-options').forEach(g => {
            const sel = g.querySelector('.inspection-opt.selected');
            data[g.dataset.label] = sel ? sel.dataset.val : '';
        });
        return data;
    }

    // === TEMPLATES ===
    async function loadTemplates() {
        if (!db) return;
        const tx = db.transaction('templates', 'readonly');
        savedTemplates = await new Promise(r => { const q = tx.objectStore('templates').getAll(); q.onsuccess = () => r(q.result); });
        templateSelector.innerHTML = '<option value="">-- Seleccionar Plantilla --</option>';
        savedTemplates.forEach(t => { const o = document.createElement('option'); o.value = t.id; o.textContent = t.name; templateSelector.appendChild(o); });
    }

    // === EVENT LISTENERS ===
    function setupEventListeners() {
        document.getElementById('dropZone').addEventListener('click', () => fileInput.click());
        sheetSelector.addEventListener('change', e => { currentClinic = e.target.value; renderTable(); });
        serieFilter.addEventListener('input', renderTable);

        certFileInput.addEventListener('change', async e => {
            const file = e.target.files[0]; if (!file) return;
            const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
            saveTemplateRow.classList.toggle('hidden', !isExcel);
            if (isExcel) {
                try {
                    const extracted = await extractInstrumentsFromExcel(file);
                    if (extracted && extracted.length > 0) {
                        const cur = getInstrumentsData();
                        if (cur.length > 0) {
                            if (confirm(`Se detectaron ${extracted.length} instrumentos. ¿Añadir a la lista?`)) {
                                extracted.forEach(inst => createInstrumentRow(inst));
                            } else {
                                instrumentsContainer.innerHTML = '';
                                extracted.forEach(inst => createInstrumentRow(inst));
                            }
                        } else {
                            extracted.forEach(inst => createInstrumentRow(inst));
                        }
                    }
                } catch (err) { console.error("Error extracción:", err); }
            }
        });

        saveNewTemplateBtn.addEventListener('click', async () => {
            const file = certFileInput.files[0], name = templateNameInput.value.trim();
            if (!file || !name) { alert('Falta archivo o nombre'); return; }
            const tx = db.transaction('templates', 'readwrite');
            tx.objectStore('templates').add({ name, blob: file });
            tx.oncomplete = () => { alert('Plantilla guardada'); templateNameInput.value = ''; saveTemplateRow.classList.add('hidden'); loadTemplates(); };
        });

        document.getElementById('closeModalBtn').onclick = () => editModal.classList.add('hidden');

        equiposTableBody.addEventListener('click', e => {
            const btn = e.target.closest('[data-action]');
            if (!btn) return;
            const serie = btn.dataset.serie;
            if (btn.dataset.action === 'viewCert') {
                const c = calibrationDates[serie];
                if (c?.certificate) window.open(URL.createObjectURL(c.certificate), '_blank');
            } else if (btn.dataset.action === 'openEdit') {
                openEditModal(serie);
            }
        });

        function openEditModal(serie) {
            selectedSerieForEdit = serie;
            const existing = calibrationDates[serie] || {};
            const eqRow = (allSheetsData[currentClinic] || [])
                .find(r => String(r[Object.keys(r).find(k => k.toLowerCase().includes('serie') || k.toLowerCase().includes('n°'))] || '').toUpperCase() === serie) || {};

            calibDateInput.value = existing.date || '';
            ordenMInput.value = existing.ordenM || '';
            technicianInput.value = existing.technician || '';
            buildingInput.value = existing.building || eqRow.edificio || '';
            sectorInput.value = existing.sector || eqRow.sector || '';
            locationInput.value = existing.location || eqRow.ubicacion || eqRow['ubicación'] || '';
            equipmentNameInput.value = existing.editedName || eqRow.equipo || '';
            modalSerieInput.value = existing.editedSerie || serie;
            modelInput.value = existing.model || eqRow.modelo || '';
            brandInput.value = existing.brand || eqRow.marca || '';
            commentsInput.value = existing.comments || '';

            instrumentsContainer.innerHTML = '';
            if (existing.instruments && existing.instruments.length > 0) {
                existing.instruments.forEach(i => createInstrumentRow(i));
            } else if (existing.certificate && existing.certName && (existing.certName.toLowerCase().endsWith('.xlsx') || existing.certName.toLowerCase().endsWith('.xls'))) {
                extractInstrumentsFromExcel(existing.certificate).then(extracted => {
                    if (extracted && extracted.length > 0) {
                        instrumentsContainer.innerHTML = '';
                        extracted.forEach(inst => createInstrumentRow(inst));
                    }
                }).catch(err => console.error(err));
            }

            // Evaluación
            const evalContainer = document.getElementById('evaluationStatusContainer');
            evalContainer.innerHTML = '';
            EVALUATION_SCHEMA.forEach(item => {
                const cur = (existing.evaluations || {})[item.label] || '';
                const rowEl = document.createElement('div'); rowEl.className = 'inspection-row';
                rowEl.innerHTML = `
                    <div class="inspection-label">${item.label}</div>
                    <div class="inspection-options" data-label="${item.label}" data-type="evaluation">
                        <div class="inspection-opt ${cur === 'P' ? 'selected' : ''}" data-val="P">P</div>
                        <div class="inspection-opt ${cur === 'NA' ? 'selected' : ''}" data-val="NA">NA</div>
                    </div>`;
                rowEl.querySelectorAll('.inspection-opt').forEach(o => o.onclick = () => {
                    rowEl.querySelectorAll('.inspection-opt').forEach(x => x.classList.remove('selected'));
                    o.classList.add('selected');
                });
                evalContainer.appendChild(rowEl);
            });

            certStatus.textContent = existing.certName ? `Certificado: ${existing.certName}` : 'Sin certificado';
            renderInspectionPoints(existing.inspections || {});
            editModal.classList.remove('hidden');
        }

        document.getElementById('resetCalibBtn').onclick = async () => {
            if (!selectedSerieForEdit) return;
            if (!confirm(`¿Resetear "${selectedSerieForEdit}"? No se puede deshacer.`)) return;
            const tx = db.transaction('calibrations', 'readwrite');
            tx.objectStore('calibrations').delete(selectedSerieForEdit);
            tx.oncomplete = () => { editModal.classList.add('hidden'); renderTable(); alert('Equipo reseteado.'); };
        };

        saveCalibBtn.onclick = async () => {
            if (!selectedSerieForEdit) return;
            if (!calibDateInput.value) { alert('Fecha requerida'); return; }
            const inspections = getInspectionsData();
            const evaluations = getEvaluationsData();
            const instruments = getInstrumentsData();
            const tmpl = savedTemplates.find(t => String(t.id) === String(templateSelector.value));
            let blob = certFileInput.files[0] || (tmpl ? tmpl.blob : null);

            try {
                let finalCert = blob;
                if (blob && (blob.name?.endsWith('.xlsx') || blob.name?.endsWith('.xls'))) {
                    finalCert = await updateExcelCertificate(blob, {
                        editedName: equipmentNameInput.value,
                        editedSerie: modalSerieInput.value,
                        model: modelInput.value, brand: brandInput.value,
                        building: buildingInput.value, sector: sectorInput.value,
                        location: locationInput.value, date: calibDateInput.value,
                        ordenM: ordenMInput.value, technician: technicianInput.value,
                        instruments, inspections, evaluations,
                    });
                }

                await storeCalibration({
                    serie: selectedSerieForEdit, date: calibDateInput.value,
                    technician: technicianInput.value, ordenM: ordenMInput.value,
                    building: buildingInput.value, sector: sectorInput.value,
                    location: locationInput.value, brand: brandInput.value,
                    model: modelInput.value, comments: commentsInput.value,
                    editedName: equipmentNameInput.value, editedSerie: modalSerieInput.value,
                    instruments, inspections, evaluations,
                    certificate: finalCert, certName: finalCert?.name,
                });
                editModal.classList.add('hidden');
                renderTable();
                alert('✅ Calibración guardada exitosamente.');
            } catch (err) { console.error(err); alert('Error: ' + err.message); }
        };
    }

    // === EXCEL CERTIFICATE UPDATE ===
    async function updateExcelCertificate(blob, d) {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await blob.arrayBuffer());
        const ws = wb.getWorksheet('Certificado') || wb.worksheets[0];
        if (!ws) throw new Error('Hoja "Certificado" no encontrada');

        // ── Cabecera ──
        ws.getCell('A5').value = `Equipo: ${d.editedName}`;
        ws.getCell('E5').value = `Modelo: ${d.model}`;
        ws.getCell('A8').value = `N° serie: ${d.editedSerie}`;
        ws.getCell('E8').value = `Marca: ${d.brand}`;
        ws.getCell('H5').value = d.building;
        ws.getCell('H6').value = d.sector;
        ws.getCell('H7').value = d.location;
        ws.getCell('H8').value = d.date ? formatDate(d.date) : '';
        ws.getCell('H9').value = d.ordenM;
        ws.getCell('H10').value = d.technician;

        // ── Instrumental (filas 12+) ──
        (d.instruments || []).forEach((inst, i) => {
            if (i >= 4) return;
            const r = 12 + i;
            ws.getCell(`A${r}`).value = inst.name || '';
            ws.getCell(`B${r}`).value = inst.brand || '';
            ws.getCell(`D${r}`).value = inst.model || '';
            ws.getCell(`E${r}`).value = inst.serie || '';
            ws.getCell(`F${r}`).value = inst.date ? formatDate(inst.date) : '';
        });

        // ── Evaluación (filas 15-17) ──
        EVALUATION_SCHEMA.forEach(item => {
            const val = (d.evaluations || {})[item.label] || '';
            ws.getCell(`H${item.row}`).value = val === 'P' ? 'x' : (val || 'NA');
        });

        // ── 8.1 Inspección (14 ítems) ──
        SCHEMA_81.forEach(item => {
            const val = (d.inspections || {})[`81_${item.label}`] || '';
            if (val === 'P') ws.getCell(`H${item.row}`).value = 'P';
            else if (val === 'F') ws.getCell(`H${item.row}`).value = 'F';
            else ws.getCell(`H${item.row}`).value = 'N/A';
        });

        // ── 8.2 Test Cuantitativos ──
        SCHEMA_82.forEach(test => {
            const td = (d.inspections || {})[`_quant_${test.code}`] || {};
            for (let i = 0; i < test.readings; i++) {
                const rowOffset = test.row + i;
                const setVal = safeParseFloat(td[`set${i}`]);
                const medVal = safeParseFloat(td[`med${i}`]);
                const pfVal = td[`pf${i}`] || 'N/A';
                if (setVal !== null) ws.getCell(`G${rowOffset}`).value = setVal;
                if (medVal !== null) ws.getCell(`H${rowOffset}`).value = medVal;
                ws.getCell(`I${rowOffset}`).value = pfVal === 'P' ? 'P' : (pfVal === 'F' ? 'F' : 'N/A');
            }
        });

        // ── 8.3 Mantenimiento Preventivo ──
        SCHEMA_83.forEach(item => {
            const val = (d.inspections || {})[`83_${item.label}`] || '';
            if (val === 'P') ws.getCell(`H${item.row}`).value = 'P';
            else if (val === 'F') ws.getCell(`H${item.row}`).value = 'F';
            else ws.getCell(`H${item.row}`).value = 'N/A';
        });

        // ── 8.4 Test de Aceptación ──
        SCHEMA_84.forEach(item => {
            const val = (d.inspections || {})[`84_${item.label}`] || '';
            if (val === 'P') ws.getCell(`H${item.row}`).value = 'P';
            else if (val === 'F') ws.getCell(`H${item.row}`).value = 'F';
            else ws.getCell(`H${item.row}`).value = 'N/A';
        });

        const out = await wb.xlsx.writeBuffer();
        return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }

    function formatDate(s) {
        if (!s) return '';
        const d = new Date(s + 'T12:00:00');
        if (isNaN(d.getTime())) return s;
        return d.toLocaleDateString('es-ES');
    }

    function formatDateForInput(s) {
        if (!s) return '';
        if (s instanceof Date) return s.toISOString().split('T')[0];
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
        const d = new Date(s);
        if (!isNaN(d.getTime())) return d.toISOString().split('T')[0];
        return '';
    }

    async function extractInstrumentsFromExcel(blob) {
        try {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(await blob.arrayBuffer());
            const worksheet = workbook.getWorksheet('Certificado') || workbook.worksheets[0];
            if (!worksheet) return [];
            const instruments = [];
            for (let i = 0; i < 4; i++) {
                const rowIdx = 12 + i;
                const name = worksheet.getCell(`A${rowIdx}`).value;
                if (!name || (typeof name === 'string' && name.trim() === '')) break;
                const lowerName = String(name).toLowerCase().trim();
                if (lowerName.includes('estado de') || lowerName.includes('comentarios')) break;
                if (lowerName.includes('instrumental') || lowerName.includes('patrón')) continue;
                const instName = String(name).trim();
                const brand = String(worksheet.getCell(`B${rowIdx}`).value || '').trim();
                const model = String(worksheet.getCell(`D${rowIdx}`).value || '').trim();
                const serie = String(worksheet.getCell(`E${rowIdx}`).value || '').trim();
                const dateCell = worksheet.getCell(`F${rowIdx}`).value;
                let dateStr = '';
                if (dateCell instanceof Date) dateStr = dateCell.toISOString().split('T')[0];
                else dateStr = String(dateCell || '').trim();
                if (instName.length > 0 && !instName.startsWith('N/A')) {
                    instruments.push({ name: instName, brand, model, serie, date: dateStr });
                }
            }
            return instruments;
        } catch (err) { console.error(err); return []; }
    }

    init();
});
