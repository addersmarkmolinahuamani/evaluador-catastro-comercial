/**
 * EVALUADOR DE META DE GESTIÓN: CATASTRO COMERCIAL
 * Aplicativo web para el cruce y validación de las 3 bases de datos remitidas por EPS:
 * 1. Conexiones Totales (Padrón general)
 * 2. Fichas Catastrales (Levantamiento de campo)
 * 3. Tabla Atributiva GIS (Plataforma gráfica georreferenciada)
 */

document.addEventListener('DOMContentLoaded', () => {
  // Application State
  const state = {
    evalYear: 2025,
    epsName: 'EPS SEDAPAL S.A.',
    targetGoal: 100,
    dateRangeMode: 'year',
    customStart: '2025-01-01',
    customEnd: '2025-12-31',
    
    // Multi-file bases parsed from Excel (.xlsx, .xls, .xlsb, .csv)
    bases: {
      1: { files: [], data: [], columns: [] },
      2: { files: [], data: [], columns: [] },
      3: { files: [], data: [], columns: [] }
    },

    // Column mappings
    mappings: {
      base1: { suministro: '', servicio: '', usuario: '', direccion: '', localidad: '', fecha: '' },
      base2: { localidad: '', suministro: '', servicio: '', fecha: '', estado: '' },
      base3: { localidad: '', suministro: '', servicio: '', fecha: '', codigo: '' }
    },

    // Dedicated Date Range Filter for Base 1
    base1DateFilter: {
      enabled: false,
      start: '',
      end: ''
    },

    // Dedicated Locality Filter for Base 1
    base1LocalityFilter: 'ALL',

    // Active Results Stage (1, 2, 3, or 4)
    activeStage: 4,

    // Stage 1, 2, 3 data & filter states
    stages: {
      s1: { data: [], matrix: [], search: '', localityFilter: 'ALL', userTypeFilter: 'ALL', statusFilter: 'ALL', page: 1, pageSize: 25 },
      s2: { data: [], matrix: [], search: '', localityFilter: 'ALL', userTypeFilter: 'ALL', statusFilter: 'ALL', page: 1, pageSize: 25 },
      s3: { data: [], matrix: [], search: '', localityFilter: 'ALL', userTypeFilter: 'ALL', statusFilter: 'ALL', page: 1, pageSize: 25 }
    },

    // Processed evaluation output
    evaluatedRecords: [],
    stats: null,
    userTypeFilter: 'ALL', // 'ALL' | 'ANTIGUO' | 'NUEVO'

    // Table view state
    table: {
      search: '',
      statusFilter: 'ALL',
      serviceFilter: 'ALL',
      localityFilter: 'ALL',
      userTypeFilter: 'ALL', // 'ALL' | 'ANTIGUO' | 'NUEVO'
      page: 1,
      pageSize: 30,
      sortBy: 'suministro',
      sortAsc: true
    },

    // Chart.js instances
    charts: {
      donut: null,
      service: null,
      locality: null,
      timeline: null
    }
  };

  // DOM Elements
  const evalYearInput = document.getElementById('eval-year');
  const epsNameInput = document.getElementById('eps-name');
  const targetGoalInput = document.getElementById('target-goal');
  const dateRangeSelect = document.getElementById('date-range-mode');
  const customRangeContainer = document.getElementById('custom-range-container');
  const startDateInput = document.getElementById('start-date');
  const endDateInput = document.getElementById('end-date');

  const btnProcess = document.getElementById('btn-process-cross');
  const btnLoadDemo = document.getElementById('btn-load-demo');
  const btnResetAll = document.getElementById('btn-reset-all');
  const btnToggleTheme = document.getElementById('btn-toggle-theme');

  const mappingContainer = document.getElementById('mapping-container');
  const btnToggleMapping = document.getElementById('btn-toggle-mapping');
  const mappingBody = document.getElementById('mapping-body');
  const mappingToggleText = document.getElementById('mapping-toggle-text');
  const mappingToggleIcon = document.getElementById('mapping-toggle-icon');

  const sectionResults = document.getElementById('section-results');
  const btnExportConsolidated = document.getElementById('btn-export-consolidated');
  const btnExportObservations = document.getElementById('btn-export-observations');

  const tableBody = document.getElementById('evaluation-table-body');
  const searchInput = document.getElementById('table-search-input');
  const btnClearSearch = document.getElementById('btn-clear-search');
  const filterServiceSelect = document.getElementById('filter-service-select');
  const filterLocalitySelect = document.getElementById('filter-locality-select');
  const pageSizeSelect = document.getElementById('table-page-size');
  const statusChips = document.querySelectorAll('.filter-chip');
  const localitySummaryBody = document.getElementById('locality-summary-body');

  const btnPagePrev = document.getElementById('btn-page-prev');
  const btnPageNext = document.getElementById('btn-page-next');
  const currentPageSpan = document.getElementById('current-page');
  const totalPagesSpan = document.getElementById('total-pages');
  const tableShowingCount = document.getElementById('table-showing-count');
  const tableTotalCount = document.getElementById('table-total-count');

  // Modal elements
  const detailModal = document.getElementById('detail-modal');
  const modalContent = document.getElementById('modal-content');
  const btnCloseModal = document.getElementById('btn-close-modal');
  const btnModalDismiss = document.getElementById('btn-modal-dismiss');

  // Base 1 Dedicated Date Filter Elements
  const btnToggleB1Date = document.getElementById('btn-toggle-b1-date');
  const badgeB1DateStatus = document.getElementById('badge-b1-date-status');
  const containerB1DateInputs = document.getElementById('container-b1-date-inputs');
  const b1DateStartInput = document.getElementById('b1-date-start');
  const b1DateEndInput = document.getElementById('b1-date-end');
  const rulePillBase1 = document.getElementById('rule-pill-base1');
  const ruleTextBase1 = document.getElementById('rule-text-base1');
  const ruleIconBase1 = document.getElementById('rule-icon-base1');

  // Initialize Lucide Icons
  if (window.lucide) {
    window.lucide.createIcons();
  }

  // Event Listeners for Configuration Inputs
  evalYearInput.addEventListener('change', (e) => {
    state.evalYear = parseInt(e.target.value) || 2025;
    if (state.dateRangeMode === 'year') {
      state.customStart = `${state.evalYear}-01-01`;
      state.customEnd = `${state.evalYear}-12-31`;
      startDateInput.value = state.customStart;
      endDateInput.value = state.customEnd;
    }
  });

  epsNameInput.addEventListener('input', (e) => {
    state.epsName = e.target.value.trim() || 'EPS';
  });

  if (targetGoalInput) {
    targetGoalInput.addEventListener('input', (e) => {
      state.targetGoal = parseInt(e.target.value) || 0;
    });
  }

  dateRangeSelect.addEventListener('change', (e) => {
    state.dateRangeMode = e.target.value;
    if (state.dateRangeMode === 'custom') {
      customRangeContainer.classList.remove('hidden');
    } else {
      customRangeContainer.classList.add('hidden');
      state.customStart = `${state.evalYear}-01-01`;
      state.customEnd = `${state.evalYear}-12-31`;
    }
  });

  startDateInput.addEventListener('change', (e) => {
    state.customStart = e.target.value;
  });

  endDateInput.addEventListener('change', (e) => {
    state.customEnd = e.target.value;
  });

  // Toggle Theme
  btnToggleTheme.addEventListener('click', () => {
    const isDark = document.body.classList.toggle('theme-dark');
    document.body.classList.toggle('theme-light', !isDark);
    const themeIcon = document.getElementById('theme-icon');
    if (themeIcon) {
      themeIcon.setAttribute('data-lucide', isDark ? 'sun' : 'moon');
      if (window.lucide) window.lucide.createIcons();
    }
    // Update chart colors if charts exist
    if (state.charts.donut) updateChartsTheme();
  });

  // Setup Dropzones for the 3 bases
  [1, 2, 3].forEach(baseId => {
    setupDropzone(baseId);
  });

  function setupDropzone(baseId) {
    const dropzone = document.getElementById(`dropzone-base${baseId}`);
    const area = document.getElementById(`area-base${baseId}`);
    const input = document.getElementById(`file-base${baseId}`);

    ['dragenter', 'dragover'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.add('dragover');
        area.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropzone.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.remove('dragover');
        area.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');
      area.classList.remove('dragover');
      if (e.dataTransfer && e.dataTransfer.files.length > 0) {
        handleFiles(baseId, e.dataTransfer.files);
      }
    });

    area.addEventListener('click', (e) => {
      input.click();
    });

    input.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleFiles(baseId, e.target.files);
        input.value = ''; // Reset input to allow selecting the same file again if needed
      }
    });

    // Event delegation on dropzone for "Agregar más", "Limpiar", "Quitar archivo"
    dropzone.addEventListener('click', (e) => {
      const addBtn = e.target.closest('.btn-add-more, .summary-drop-hint');
      if (addBtn) {
        const bId = parseInt(addBtn.dataset.base || baseId);
        if (bId === baseId) {
          e.stopPropagation();
          input.click();
          return;
        }
      }

      const clearBtn = e.target.closest('.btn-clear-base');
      if (clearBtn && parseInt(clearBtn.dataset.base) === baseId) {
        e.stopPropagation();
        clearBaseFiles(baseId);
        return;
      }

      const removeBtn = e.target.closest('.btn-remove-single-file');
      if (removeBtn && parseInt(removeBtn.dataset.base) === baseId) {
        e.stopPropagation();
        const fileId = removeBtn.dataset.fileId;
        removeFileFromBase(baseId, fileId);
      }
    });
  }

  // Clear all files for a base
  function clearBaseFiles(baseId) {
    const base = state.bases[baseId];
    base.files = [];
    rebuildBaseData(baseId);
    updateFilesSummaryUI(baseId);
    autoDetectMapping(baseId, []);
    updateMappingView();
    checkProcessReady();
    showToast(`Archivos de Base ${baseId} eliminados`, 'info');
  }

  // Remove a single file from a base
  function removeFileFromBase(baseId, fileId) {
    const base = state.bases[baseId];
    base.files = base.files.filter(f => f.id !== fileId);
    rebuildBaseData(baseId);
    updateFilesSummaryUI(baseId);
    autoDetectMapping(baseId, base.columns);
    updateMappingView();
    checkProcessReady();
    showToast(`Archivo eliminado de Base ${baseId}`, 'info');
  }

  // Handle uploaded Excel file(s)
  async function handleFiles(baseId, fileList) {
    const files = Array.from(fileList);
    if (!files || files.length === 0) return;

    showToast(`Cargando y procesando ${files.length} archivo(s)...`, 'info');

    for (const file of files) {
      try {
        const data = await readFileAsArrayBuffer(file);
        const workbook = XLSX.read(new Uint8Array(data), { type: 'array', cellDates: true });
        
        // Pick first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert sheet to json
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });
        
        if (!jsonData || jsonData.length === 0) {
          showToast(`El archivo ${file.name} no contiene filas con datos`, 'warning');
          continue;
        }

        const columns = Object.keys(jsonData[0]);

        // Clean locality hint from filename (e.g. "01_SAN_VICENTE.xlsb" -> "SAN VICENTE")
        const cleanFileLocality = file.name
          .replace(/\.[^/.]+$/, "") // Remove file extension
          .replace(/^\d+[\s_-]*/, "") // Remove leading numbers (01_, 02-, etc.)
          .replace(/[_-]+/g, " ") // Replace underscores and hyphens with spaces
          .trim().toUpperCase();

        // If the file does not have an explicit Localidad column, assign the filename as Localidad
        const hasLocCol = columns.some(c => /localidad|distrito|sector|zona|anexo/i.test(c));
        if (!hasLocCol && cleanFileLocality) {
          jsonData.forEach(row => {
            if (!row['Localidad']) row['Localidad'] = cleanFileLocality;
          });
          if (!columns.includes('Localidad')) columns.push('Localidad');
        }

        const fileObj = {
          id: 'f_' + Date.now() + '_' + Math.random().toString(36).substr(2, 6),
          name: file.name,
          size: file.size,
          rows: jsonData,
          rowCount: jsonData.length,
          columns: columns
        };

        state.bases[baseId].files.push(fileObj);
      } catch (err) {
        console.error(err);
        showToast(`Error al procesar ${file.name}: ${err.message}`, 'error');
      }
    }

    // Rebuild concatenated data & columns for this base
    rebuildBaseData(baseId);
    updateFilesSummaryUI(baseId);
    autoDetectMapping(baseId, state.bases[baseId].columns);
    updateMappingView();
    checkProcessReady();

    showToast(`Base ${baseId}: ${state.bases[baseId].files.length} archivo(s) | ${state.bases[baseId].data.length.toLocaleString()} filas acumuladas`, 'success');
  }

  function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target.result);
      reader.onerror = (e) => reject(new Error('Error de lectura de archivo'));
      reader.readAsArrayBuffer(file);
    });
  }

  // Rebuild concatenated data and merged column headers for a base
  function rebuildBaseData(baseId) {
    const base = state.bases[baseId];
    if (!base.files || base.files.length === 0) {
      base.data = [];
      base.columns = [];
      return;
    }

    base.data = base.files.flatMap(f => f.rows);

    const allCols = new Set();
    base.files.forEach(f => f.columns.forEach(c => allCols.add(c)));
    base.columns = Array.from(allCols);
  }

  // Update Dropzone and Files list UI
  function updateFilesSummaryUI(baseId) {
    const base = state.bases[baseId];
    const area = document.getElementById(`area-base${baseId}`);
    const summary = document.getElementById(`summary-base${baseId}`);
    const metaLabel = document.getElementById(`meta-base${baseId}`);
    const filesListEl = document.getElementById(`files-list-base${baseId}`);

    if (!base.files || base.files.length === 0) {
      area.classList.remove('hidden');
      summary.classList.add('hidden');
      if (filesListEl) filesListEl.innerHTML = '';
      metaLabel.textContent = `0 archivos | 0 filas`;
    } else {
      area.classList.add('hidden');
      summary.classList.remove('hidden');
      metaLabel.textContent = `${base.files.length} arch. | ${base.data.length.toLocaleString()} filas`;
      metaLabel.title = `${base.files.length} archivo(s) | ${base.data.length.toLocaleString()} filas totales`;

      if (filesListEl) {
        filesListEl.innerHTML = '';
        base.files.forEach(f => {
          const item = document.createElement('div');
          item.className = 'file-item-row';
          item.innerHTML = `
            <div class="file-item-left">
              <i data-lucide="file-spreadsheet" class="file-item-icon"></i>
              <div class="file-item-text">
                <span class="file-item-name" title="${f.name}">${f.name}</span>
                <span class="file-item-meta">${f.rowCount.toLocaleString()} filas | ${f.columns.length} col.</span>
              </div>
            </div>
            <button type="button" class="btn-icon-subtle btn-remove-single-file" data-base="${baseId}" data-file-id="${f.id}" title="Quitar este archivo">
              <i data-lucide="x"></i>
            </button>
          `;
          filesListEl.appendChild(item);
        });

        if (window.lucide) window.lucide.createIcons();
      }
    }
  }

  // Intelligent column auto-detection
  function autoDetectMapping(baseId, columns) {
    const colLower = columns.map(c => ({ original: c, clean: c.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim() }));

    function findMatch(patterns) {
      for (const p of patterns) {
        const match = colLower.find(c => p.test(c.clean));
        if (match) return match.original;
      }
      return '';
    }

    if (baseId === 1) {
      state.mappings.base1.suministro = findMatch([
        /^numero.*suministro$/,
        /suministro/,
        /^sumin/,
        /^nis/,
        /^cod.*suministro/
      ]) || columns[0] || '';

      state.mappings.base1.servicio = findMatch([
        /^tipo.*servicio/,
        /servicio/,
        /serv/
      ]);

      state.mappings.base1.usuario = findMatch([
        /usuario/,
        /cliente/,
        /titular/,
        /nombre/,
        /razon.*social/
      ]);

      state.mappings.base1.direccion = findMatch([
        /direccion/,
        /predio/,
        /calle/,
        /ubicacion/
      ]);

      state.mappings.base1.localidad = findMatch([
        /localidad/,
        /distrito/,
        /sector/,
        /zona/,
        /anexo/,
        /centro.*poblado/
      ]);

      state.mappings.base1.fecha = findMatch([
        /^fecha.*instalacion/,
        /^fecha.*conexion/,
        /^fecha.*registro/,
        /^fecha.*alta/,
        /^fecha.*facturacion/,
        /fecha/
      ]);
    } else if (baseId === 2) {
      state.mappings.base2.localidad = findMatch([
        /localidad/,
        /distrito/,
        /sector/,
        /zona/,
        /anexo/,
        /centro.*poblado/
      ]);

      state.mappings.base2.suministro = findMatch([
        /^numero.*suministro$/,
        /nro.*sumin/,
        /dist_sumi/,
        /suministro/,
        /^sumin/,
        /^nis/,
        /^cod.*suministro/
      ]) || columns[0] || '';

      state.mappings.base2.servicio = findMatch([
        /^tipo.*servicio/,
        /servicio/,
        /serv/
      ]);

      state.mappings.base2.fecha = findMatch([
        /fecha.*ficha/,
        /fecha.*catastr/,
        /fecha.*levantamiento/,
        /fecha.*censo/,
        /fecha.*de.*r/,
        /fecha.*registro/,
        /^fecha$/
      ]) || '';
    } else if (baseId === 3) {
      state.mappings.base3.localidad = findMatch([
        /localidad/,
        /distrito/,
        /sector/,
        /zona/,
        /anexo/,
        /centro.*poblado/
      ]);

      state.mappings.base3.suministro = findMatch([
        /^numero.*suministro$/,
        /nro.*sumin/,
        /dist_sumi/,
        /suministro/,
        /^sumin/,
        /^nis/,
        /^cod.*suministro/
      ]) || columns[0] || '';

      state.mappings.base3.servicio = findMatch([
        /^tipo.*servicio/,
        /servicio/,
        /serv/
      ]);

      state.mappings.base3.fecha = findMatch([
        /fecha.*registro.*sistema.*grafic/,
        /fecha.*sistema.*grafic/,
        /fecha.*gis/,
        /fecha.*ingreso.*gis/,
        /fecha.*digitalizac/,
        /fecha.*de.*r/,
        /fecha.*registro/,
        /^fecha$/
      ]) || '';
    }
  }

  // Populate Mapping Select Elements
  function updateMappingView() {
    const hasAnyFile = (state.bases[1]?.files?.length > 0) || (state.bases[2]?.files?.length > 0) || (state.bases[3]?.files?.length > 0);
    if (hasAnyFile) {
      mappingContainer.classList.remove('hidden');
    } else {
      mappingContainer.classList.add('hidden');
      return;
    }

    populateSelect('map-b1-suministro', state.bases[1]?.columns, state.mappings.base1.suministro);
    populateSelect('map-b1-localidad', state.bases[1]?.columns, state.mappings.base1.localidad, true);
    populateSelect('map-b1-servicio', state.bases[1]?.columns, state.mappings.base1.servicio, true);
    populateSelect('map-b1-fecha', state.bases[1]?.columns, state.mappings.base1.fecha, true);

    populateSelect('map-b2-suministro', state.bases[2]?.columns, state.mappings.base2.suministro);
    populateSelect('map-b2-localidad', state.bases[2]?.columns, state.mappings.base2.localidad, true);
    populateSelect('map-b2-servicio', state.bases[2]?.columns, state.mappings.base2.servicio, true);
    populateSelect('map-b2-fecha', state.bases[2]?.columns, state.mappings.base2.fecha);

    populateSelect('map-b3-suministro', state.bases[3]?.columns, state.mappings.base3.suministro);
    populateSelect('map-b3-localidad', state.bases[3]?.columns, state.mappings.base3.localidad, true);
    populateSelect('map-b3-servicio', state.bases[3]?.columns, state.mappings.base3.servicio, true);
    populateSelect('map-b3-fecha', state.bases[3]?.columns, state.mappings.base3.fecha);
  }

  function populateSelect(selectId, columns, selectedValue, optional = false) {
    const select = document.getElementById(selectId);
    if (!select) return;
    select.innerHTML = '';

    if (optional) {
      const optNone = document.createElement('option');
      optNone.value = '';
      optNone.textContent = '(Ninguno / Opcional)';
      select.appendChild(optNone);
    }

    if (!columns || columns.length === 0) {
      const optEmpty = document.createElement('option');
      optEmpty.value = '';
      optEmpty.textContent = 'Archivo aún no cargado';
      select.appendChild(optEmpty);
      select.disabled = true;
      return;
    }

    select.disabled = false;
    columns.forEach(col => {
      const opt = document.createElement('option');
      opt.value = col;
      opt.textContent = col;
      if (col === selectedValue) opt.selected = true;
      select.appendChild(opt);
    });
  }

  // Bind mapping selects to state
  bindMappingSelect('map-b1-suministro', (val) => state.mappings.base1.suministro = val);
  bindMappingSelect('map-b1-localidad', (val) => state.mappings.base1.localidad = val);
  bindMappingSelect('map-b1-servicio', (val) => state.mappings.base1.servicio = val);
  bindMappingSelect('map-b1-fecha', (val) => state.mappings.base1.fecha = val);

  bindMappingSelect('map-b2-suministro', (val) => state.mappings.base2.suministro = val);
  bindMappingSelect('map-b2-localidad', (val) => state.mappings.base2.localidad = val);
  bindMappingSelect('map-b2-servicio', (val) => state.mappings.base2.servicio = val);
  bindMappingSelect('map-b2-fecha', (val) => state.mappings.base2.fecha = val);

  bindMappingSelect('map-b3-suministro', (val) => state.mappings.base3.suministro = val);
  bindMappingSelect('map-b3-localidad', (val) => state.mappings.base3.localidad = val);
  bindMappingSelect('map-b3-servicio', (val) => state.mappings.base3.servicio = val);
  bindMappingSelect('map-b3-fecha', (val) => state.mappings.base3.fecha = val);

  function bindMappingSelect(selectId, updateFn) {
    const el = document.getElementById(selectId);
    if (el) {
      el.addEventListener('change', (e) => {
        updateFn(e.target.value);
        checkProcessReady();
      });
    }
  }

  // Toggle Mapping Card Accordion
  btnToggleMapping.addEventListener('click', () => {
    const isHidden = mappingBody.classList.toggle('hidden');
    mappingToggleText.textContent = isHidden ? 'Mostrar Mapeo' : 'Ocultar Mapeo';
    mappingToggleIcon.setAttribute('data-lucide', isHidden ? 'chevron-down' : 'chevron-up');
    if (window.lucide) window.lucide.createIcons();
  });

  // Check if ready to process
  function checkProcessReady() {
    const note = document.getElementById('process-status-note');
    const b1 = state.bases[1];
    const b2 = state.bases[2];
    const b3 = state.bases[3];

    const hasBases = (b1.files.length > 0) && (b2.files.length > 0) && (b3.files.length > 0);
    const hasKeys = state.mappings.base1.suministro &&
                    state.mappings.base2.suministro && state.mappings.base2.fecha &&
                    state.mappings.base3.suministro && state.mappings.base3.fecha;

    if (hasBases && hasKeys) {
      btnProcess.disabled = false;
      note.textContent = '¡Listo para procesar! Haz clic para ejecutar el cruce y validar la meta.';
      note.style.color = 'var(--success)';
    } else {
      btnProcess.disabled = true;
      let missing = [];
      if (!b1 || b1.files.length === 0) missing.push('Base 1 (Conexiones Totales)');
      if (!b2 || b2.files.length === 0) missing.push('Base 2 (Fichas Catastrales)');
      if (!b3 || b3.files.length === 0) missing.push('Base 3 (Tabla GIS)');
      note.textContent = `Pendiente por cargar: ${missing.join(', ')}`;
      note.style.color = 'var(--text-muted)';
    }
  }

  // Standardize Suministro
  function normalizeSuministro(val) {
    if (val === null || val === undefined) return '';
    return String(val).trim().toUpperCase();
  }

  // Parse Date Flexible
  function parseDateFlexible(val) {
    if (!val) return null;
    
    // If it's already a Date
    if (val instanceof Date && !isNaN(val)) {
      return val;
    }

    const str = String(val).trim();
    if (!str) return null;

    // Excel serial number test (e.g. 45123)
    if (!isNaN(str) && Number(str) > 30000 && Number(str) < 60000) {
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      return new Date(excelEpoch.getTime() + Number(str) * 86400000);
    }

    // Format DD/MM/YYYY or DD-MM-YYYY
    const dmy = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);
    if (dmy) {
      const d = parseInt(dmy[1], 10);
      const m = parseInt(dmy[2], 10) - 1;
      const y = parseInt(dmy[3], 10);
      return new Date(y, m, d);
    }

    // Format YYYY-MM-DD
    const ymd = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
    if (ymd) {
      const y = parseInt(ymd[1], 10);
      const m = parseInt(ymd[2], 10) - 1;
      const d = parseInt(ymd[3], 10);
      return new Date(y, m, d);
    }

    const parsed = new Date(str);
    return isNaN(parsed) ? null : parsed;
  }

  function isDateInEvaluation(dateObj) {
    if (!dateObj) return false;
    const year = dateObj.getFullYear();
    
    if (state.dateRangeMode === 'year') {
      return year === state.evalYear;
    } else {
      const start = new Date(state.customStart);
      const end = new Date(state.customEnd);
      end.setHours(23, 59, 59, 999);
      return dateObj >= start && dateObj <= end;
    }
  }

  // Distinción Regulatoria: Usuarios Antiguos vs Usuarios Nuevos
  // Regla Regulatoria: 
  // - Usuarios Antiguos: Conexiones instaladas ANTES del año regulatorio en evaluación.
  // - Usuarios Nuevos: Conexiones instaladas DURANTE el año regulatorio en evaluación.
  function classifyUserType(dateObj) {
    if (!dateObj) {
      return {
        type: 'ANTIGUO',
        label: 'Usuario Antiguo (Sin fecha)',
        isAntiguo: true,
        isNuevo: false
      };
    }
    
    let isBefore = false;
    let isInPeriod = false;
    
    if (state.dateRangeMode === 'year') {
      const y = dateObj.getFullYear();
      isBefore = y < state.evalYear;
      isInPeriod = y === state.evalYear;
    } else {
      const start = new Date(state.customStart);
      start.setHours(0, 0, 0, 0);
      const end = new Date(state.customEnd);
      end.setHours(23, 59, 59, 999);
      isBefore = dateObj < start;
      isInPeriod = dateObj >= start && dateObj <= end;
    }
    
    if (isBefore) {
      return {
        type: 'ANTIGUO',
        label: 'Usuario Antiguo',
        isAntiguo: true,
        isNuevo: false
      };
    } else if (isInPeriod) {
      return {
        type: 'NUEVO',
        label: 'Usuario Nuevo',
        isAntiguo: false,
        isNuevo: true
      };
    } else {
      return {
        type: 'POSTERIOR',
        label: 'Posterior al Periodo',
        isAntiguo: false,
        isNuevo: false
      };
    }
  }

  function formatDateDisplay(dateObj) {
    if (!dateObj) return '-';
    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const year = dateObj.getFullYear();
    return `${day}/${month}/${year}`;
  }

  // Cross-referencing Execution
  btnProcess.addEventListener('click', () => {
    executeCrossEvaluation();
  });

  // Service Classification & Normalization (Criterio SUNASS)
  // Regla SUNASS: Existen 2 servicios evaluados: 'Agua potable' y 'Alcantarillado'.
  // 'Agua potable' comprende: 'Agua potable' y 'Agua potable y alcantarillado'.
  // 'Alcantarillado' comprende: 'Alcantarillado'/'Saneamiento' y 'Agua potable y alcantarillado'.
  function getRegulatoryServices(rawServ) {
    if (!rawServ) {
      return ['Agua potable', 'Alcantarillado'];
    }
    const str = String(rawServ).trim();
    const s = str.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    
    const hasAgua = s.includes('AGUA') || s.includes('POTABLE') || s === 'AP';
    const hasAlc = s.includes('ALCANTARILLADO') || s.includes('SANEAMIENTO') || 
                   s.includes('DESAGUE') || s === 'ALC' || s === 'ALCANT';
                   
    if (hasAgua && hasAlc) {
      return ['Agua potable', 'Alcantarillado'];
    } else if (hasAlc) {
      return ['Alcantarillado'];
    } else if (hasAgua) {
      return ['Agua potable'];
    } else {
      // Default: assume both if present in municipal padrón
      return ['Agua potable', 'Alcantarillado'];
    }
  }

  // Check if a service string from Base 2 or Base 3 covers the target regulatory service
  function serviceCovers(rawServ, targetRegulatoryService) {
    if (!rawServ) return true; // If service is not specified in Ficha/GIS, assume it applies to the evaluated connection
    const s = String(rawServ).trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    
    const hasAgua = s.includes('AGUA') || s.includes('POTABLE') || s === 'AP';
    const hasAlc = s.includes('ALCANTARILLADO') || s.includes('SANEAMIENTO') || 
                   s.includes('DESAGUE') || s === 'ALC' || s === 'ALCANT';

    if (targetRegulatoryService === 'Agua potable') {
      return hasAgua || (!hasAgua && !hasAlc);
    } else if (targetRegulatoryService === 'Alcantarillado') {
      return hasAlc || (!hasAgua && !hasAlc);
    }
    return true;
  }

  // Cross-referencing Execution
  btnProcess.addEventListener('click', () => {
    executeCrossEvaluation();
  });

  function executeCrossEvaluation() {
    showToast('Ejecutando cruce tripartito y desglose regulatorio de servicios...', 'info');

    const b1 = state.bases[1].data;
    const b2 = state.bases[2].data;
    const b3 = state.bases[3].data;

    const m1 = state.mappings.base1;
    const m2 = state.mappings.base2;
    const m3 = state.mappings.base3;

    // Index Base 1 (Conexiones Totales Padrón)
    const base1Map = new Map();
    let b1LocalityFilteredOut = 0;

    b1.forEach(row => {
      const sum = normalizeSuministro(row[m1.suministro]);
      if (sum) {
        // Filtro opcional por localidad en Base 1 si no es ALL
        if (state.base1LocalityFilter && state.base1LocalityFilter !== 'ALL' && m1.localidad) {
          const rawLoc = row[m1.localidad];
          const normLoc = rawLoc ? String(rawLoc).trim().toUpperCase() : '';
          if (normLoc !== state.base1LocalityFilter) {
            b1LocalityFilteredOut++;
            return;
          }
        }
        base1Map.set(sum, row);
      }
    });

    // Index Base 2 (Fichas Catastrales en Campo)
    const base2Map = new Map();
    b2.forEach(row => {
      const sum = normalizeSuministro(row[m2.suministro]);
      if (sum) {
        const rawDate = row[m2.fecha];
        const dateObj = parseDateFlexible(rawDate);
        const item = {
          raw: row,
          dateObj: dateObj,
          isValidDate: isDateInEvaluation(dateObj),
          dateStr: formatDateDisplay(dateObj),
          servicioRaw: m2.servicio ? (row[m2.servicio] || '') : ''
        };
        if (!base2Map.has(sum)) {
          base2Map.set(sum, []);
        }
        base2Map.get(sum).push(item);
      }
    });

    // Index Base 3 (Tabla Atributiva QGIS)
    const base3Map = new Map();
    b3.forEach(row => {
      const sum = normalizeSuministro(row[m3.suministro]);
      if (sum) {
        const rawDate = row[m3.fecha];
        const dateObj = parseDateFlexible(rawDate);
        const item = {
          raw: row,
          dateObj: dateObj,
          isValidDate: isDateInEvaluation(dateObj),
          dateStr: formatDateDisplay(dateObj),
          codigo: m3.codigo ? (row[m3.codigo] || '') : '',
          servicioRaw: m3.servicio ? (row[m3.servicio] || '') : ''
        };
        if (!base3Map.has(sum)) {
          base3Map.set(sum, []);
        }
        base3Map.get(sum).push(item);
      }
    });

    // Combined universe of all unique supplies
    const allSuppliesSet = new Set([
      ...base1Map.keys(),
      ...base2Map.keys(),
      ...base3Map.keys()
    ]);

    const evaluated = [];

    // Helper to get raw locality
    function extractLocality(r1, r2List, r3List) {
      if (r1) {
        const loc1 = (m1.localidad ? r1[m1.localidad] : null) || r1['Localidad'] || r1['LOCALIDAD'] || r1['Distrito'] || r1['Sector'];
        if (loc1) return String(loc1).trim().toUpperCase();
      }
      if (r2List && r2List.length > 0) {
        const loc2 = (m2.localidad ? r2List[0].raw[m2.localidad] : null) || r2List[0].raw['Localidad'] || r2List[0].raw['LOCALIDAD'] || r2List[0].raw['Sector'];
        if (loc2) return String(loc2).trim().toUpperCase();
      }
      if (r3List && r3List.length > 0) {
        const loc3 = (m3.localidad ? r3List[0].raw[m3.localidad] : null) || r3List[0].raw['Localidad'] || r3List[0].raw['LOCALIDAD'] || r3List[0].raw['Sector'];
        if (loc3) return String(loc3).trim().toUpperCase();
      }
      return 'SIN LOCALIDAD';
    }

    allSuppliesSet.forEach(suministro => {
      const r1 = base1Map.get(suministro);
      const r2List = base2Map.get(suministro) || [];
      const r3List = base3Map.get(suministro) || [];

      const inBase1 = !!r1;

      // Fecha de instalación de conexión en Base 1 para clasificar Antiguo vs Nuevo
      const rawB1Date = (r1 && m1.fecha) ? r1[m1.fecha] : null;
      const b1DateObj = rawB1Date ? parseDateFlexible(rawB1Date) : null;
      const userTypeInfo = r1 
        ? classifyUserType(b1DateObj) 
        : { type: 'NO_PADRON', label: 'No en Padrón Base', isAntiguo: false, isNuevo: false };

      // Extract raw service string
      let rawServicio = '';
      if (r1 && m1.servicio && r1[m1.servicio]) {
        rawServicio = r1[m1.servicio];
      } else if (r2List.length > 0 && r2List[0].servicioRaw) {
        rawServicio = r2List[0].servicioRaw;
      } else if (r3List.length > 0 && r3List[0].servicioRaw) {
        rawServicio = r3List[0].servicioRaw;
      }

      const servicioReal = rawServicio ? String(rawServicio).trim() : 'Agua potable y alcantarillado';
      const targetServices = getRegulatoryServices(servicioReal);

      const usuario = r1 ? (r1[m1.usuario] || '') : (r2List.length > 0 ? (r2List[0].raw['Usuario'] || r2List[0].raw['Titular'] || '') : '');
      const direccion = r1 ? (r1[m1.direccion] || '') : (r2List.length > 0 ? (r2List[0].raw['Direccion'] || r2List[0].raw['Predio'] || '') : '');
      const localidad = extractLocality(r1, r2List, r3List);

      // Desglose regulatorio SUNASS: Crear una fila por cada servicio aplicable ('Agua potable' y/o 'Alcantarillado')
      targetServices.forEach(servicioCatastro => {
        // Matching in Base 2 (Fichas) covering this service
        const matchingFichas = r2List.filter(f => serviceCovers(f.servicioRaw, servicioCatastro));
        const inBase2 = matchingFichas.length > 0;
        const validFicha = matchingFichas.find(f => f.isValidDate);
        const fichaInYear = !!validFicha;
        const bestFicha = validFicha || matchingFichas[0] || null;

        // Matching in Base 3 (QGIS) covering this service
        const matchingGis = r3List.filter(g => serviceCovers(g.servicioRaw, servicioCatastro));
        const inBase3 = matchingGis.length > 0;
        const validGis = matchingGis.find(g => g.isValidDate);
        const gisInYear = !!validGis;
        const bestGis = validGis || matchingGis[0] || null;

        // Classify validation status for this specific service
        let status = '';
        let statusLabel = '';
        let observation = '';

        if (inBase1 && fichaInYear && gisInYear) {
          status = 'CUMPLE';
          statusLabel = 'CATASTRADO VÁLIDO (CUMPLE META)';
          observation = `Servicio de ${servicioCatastro} cumple satisfactoriamente: Ficha catastral y registro en GIS dentro del año evaluado.`;
        } else if (inBase1 && fichaInYear && !inBase3) {
          status = 'SOLO_FICHA';
          statusLabel = 'OBSERVADO: SOLO FICHA CATASTRAL';
          observation = `Servicio de ${servicioCatastro} cuenta con ficha catastral en el año, pero NO figura en la plataforma QGIS.`;
        } else if (inBase1 && !inBase2 && gisInYear) {
          status = 'SOLO_GIS';
          statusLabel = 'OBSERVADO: SOLO GIS';
          observation = `Servicio de ${servicioCatastro} figura en QGIS en el año, pero NO cuenta con ficha catastral de campo.`;
        } else if (inBase1 && (inBase2 || inBase3) && (!fichaInYear || !gisInYear)) {
          status = 'FUERA_ANO';
          statusLabel = 'OBSERVADO: FECHA FUERA DE PERIODO';
          const detalles = [];
          if (inBase2 && !fichaInYear) detalles.push(`Ficha (${bestFicha?.dateStr}) fuera de periodo`);
          if (inBase3 && !gisInYear) detalles.push(`QGIS (${bestGis?.dateStr}) fuera de periodo`);
          if (!inBase2) detalles.push('Falta ficha catastral');
          if (!inBase3) detalles.push('Falta registro en QGIS');
          observation = `Servicio de ${servicioCatastro}: ${detalles.join('; ')}.`;
        } else if (inBase1 && !inBase2 && !inBase3) {
          status = 'NO_CATASTRADO';
          statusLabel = 'NO CATASTRADO';
          observation = `Conexión registrada en el padrón base pero sin levantamiento catastral para ${servicioCatastro} en el periodo.`;
        } else if (!inBase1) {
          status = 'NO_PADRON';
          statusLabel = 'NO FIGURA EN PADRÓN BASE';
          observation = `Suministro registrado en Fichas/GIS pero no existe en el padrón de conexiones totales de la EPS.`;
        }

        evaluated.push({
          id: `${suministro}__${servicioCatastro}`,
          suministro,
          usuario,
          direccion,
          localidad,
          tipoUsuario: userTypeInfo.type, // 'ANTIGUO' | 'NUEVO' | 'POSTERIOR' | 'NO_PADRON'
          tipoUsuarioLabel: userTypeInfo.label,
          isAntiguo: userTypeInfo.isAntiguo,
          isNuevo: userTypeInfo.isNuevo,
          fechaInstalacion: b1DateObj ? formatDateDisplay(b1DateObj) : (rawB1Date ? String(rawB1Date) : '-'),
          fechaInstalacionObj: b1DateObj,
          servicioReal,
          servicioCatastro, // Strictly 'Agua potable' or 'Alcantarillado'
          tieneAgua: servicioCatastro === 'Agua potable',
          tieneAlcantarillado: servicioCatastro === 'Alcantarillado',
          inBase1,
          b1Date: (r1 && m1.fecha && r1[m1.fecha]) ? (formatDateDisplay(parseDateFlexible(r1[m1.fecha])) || String(r1[m1.fecha])) : '-',
          inBase2,
          fichaDate: bestFicha ? bestFicha.dateStr : '-',
          fichaDateObj: bestFicha ? bestFicha.dateObj : null,
          fichaInYear,
          inBase3,
          gisDate: bestGis ? bestGis.dateStr : '-',
          gisDateObj: bestGis ? bestGis.dateObj : null,
          gisInYear,
          codigoCatastral: bestGis ? bestGis.codigo : '',
          status,
          statusLabel,
          observation,
          rawBase1: r1,
          rawBase2: bestFicha ? bestFicha.raw : null,
          rawBase3: bestGis ? bestGis.raw : null
        });
      });
    });

    state.evaluatedRecords = evaluated;
    computeStats();
    populateLocalityFilter();
    renderResults();
    renderCharts();
    filterAndPaginateTable();

    // Compute & render Stages 1, 2, 3 first so their data and sets are available
    computeAndRenderStage1();
    computeAndRenderStage2();
    computeAndRenderStage3();
    updateStageTabsStats();

    // Render Stage 4 Locality Summary Table (Master Consolidated Table with 25 columns)
    renderLocalitySummaryTable();
    switchStage(state.activeStage || 4);

    // Scroll smoothly to results
    sectionResults.classList.remove('hidden');
    sectionResults.scrollIntoView({ behavior: 'smooth', block: 'start' });

    // Confetti celebration if compliance is good!
    if (state.stats.compliancePct >= 80 && window.confetti) {
      window.confetti({
        particleCount: 100,
        spread: 70,
        origin: { y: 0.6 }
      });
    }

    showToast(`¡Cruce finalizado! Se evaluaron ${evaluated.length.toLocaleString()} servicios en ${allSuppliesSet.size.toLocaleString()} conexiones.`, 'success');
  }

  // Populate Locality Filter Dropdown
  function populateLocalityFilter() {
    if (!filterLocalitySelect) return;
    const localitiesMap = new Map();
    state.evaluatedRecords.forEach(r => {
      const loc = r.localidad || 'SIN LOCALIDAD';
      localitiesMap.set(loc, (localitiesMap.get(loc) || 0) + 1);
    });

    filterLocalitySelect.innerHTML = '';
    const optAll = document.createElement('option');
    optAll.value = 'ALL';
    optAll.textContent = `Todas las localidades (${localitiesMap.size})`;
    filterLocalitySelect.appendChild(optAll);

    const sortedLocs = Array.from(localitiesMap.keys()).sort();
    sortedLocs.forEach(loc => {
      const opt = document.createElement('option');
      opt.value = loc;
      opt.textContent = `${loc} (${localitiesMap.get(loc)} reg.)`;
      if (state.table.localityFilter === loc) opt.selected = true;
      filterLocalitySelect.appendChild(opt);
    });
  }

  // Render Locality Summary Table (Master Matriz 4: Cruce Final Consolidado con Totales, Observados y Validados)
  function renderLocalitySummaryTable() {
    const summaryBody = document.getElementById('locality-summary-body');
    const summaryFooter = document.getElementById('locality-summary-footer');
    if (!summaryBody) return;

    summaryBody.innerHTML = '';
    if (summaryFooter) summaryFooter.innerHTML = '';

    // Collect all unique localities from all stages
    const allLocs = new Set();
    state.stages.s1.data.forEach(r => allLocs.add(r.localidad));
    state.stages.s2.data.forEach(r => allLocs.add(r.localidad));
    state.stages.s3.data.forEach(r => allLocs.add(r.localidad));
    state.evaluatedRecords.forEach(r => allLocs.add(r.localidad));

    const sortedLocs = Array.from(allLocs).filter(l => l && l !== 'SIN LOCALIDAD').sort();
    if (allLocs.has('SIN LOCALIDAD')) sortedLocs.push('SIN LOCALIDAD');

    const locMap = {};
    function getLoc(loc) {
      if (!locMap[loc]) {
        locMap[loc] = {
          // Base 1 Totales
          b1_ap: new Set(), b1_al: new Set(), b1_tot: new Set(),
          // Base 2 Fichas (Totales, Observadas, Validadas)
          b2_tot_ap: new Set(), b2_tot_al: new Set(), b2_tot: new Set(),
          b2_obs_ap: new Set(), b2_obs_al: new Set(), b2_obs: new Set(),
          b2_val_ap: new Set(), b2_val_al: new Set(), b2_val: new Set(),
          // Base 3 QGIS (Totales, Observados, Validados)
          b3_tot_ap: new Set(), b3_tot_al: new Set(), b3_tot: new Set(),
          b3_obs_ap: new Set(), b3_obs_al: new Set(), b3_obs: new Set(),
          b3_val_ap: new Set(), b3_val_al: new Set(), b3_val: new Set(),
          // Cumplen Meta
          cumplen: new Set()
        };
      }
      return locMap[loc];
    }

    // Build a map of supply to userType for segmentation filtering
    const supplyUserTypeMap = new Map();
    state.evaluatedRecords.forEach(r => supplyUserTypeMap.set(r.suministro, r.tipoUsuario));
    state.stages.s1.data.forEach(r => {
      if (!supplyUserTypeMap.has(r.suministro)) supplyUserTypeMap.set(r.suministro, r.tipoUsuario);
    });

    const isMatchUserType = (sum) => {
      if (!state.userTypeFilter || state.userTypeFilter === 'ALL') return true;
      const ut = supplyUserTypeMap.get(sum) || 'ANTIGUO';
      return ut === state.userTypeFilter;
    };

    // Populate Base 1 Conexiones Totales
    state.stages.s1.data.forEach(r => {
      if (!isMatchUserType(r.suministro)) return;
      const m = getLoc(r.localidad);
      m.b1_tot.add(r.suministro);
      if (r.hasAP) m.b1_ap.add(r.suministro);
      if (r.hasAL) m.b1_al.add(r.suministro);
    });

    // Populate Base 2 Fichas Catastrales
    state.stages.s2.data.forEach(r => {
      if (!isMatchUserType(r.suministro)) return;
      const m = getLoc(r.localidad);
      m.b2_tot.add(r.suministro);
      if (r.hasAP) m.b2_tot_ap.add(r.suministro);
      if (r.hasAL) m.b2_tot_al.add(r.suministro);

      if (r.isParaAnalizar) {
        if (r.isValid) {
          m.b2_val.add(r.suministro);
          if (r.hasAP) m.b2_val_ap.add(r.suministro);
          if (r.hasAL) m.b2_val_al.add(r.suministro);
        } else {
          m.b2_obs.add(r.suministro);
          if (r.hasAP) m.b2_obs_ap.add(r.suministro);
          if (r.hasAL) m.b2_obs_al.add(r.suministro);
        }
      }
    });

    // Populate Base 3 QGIS
    state.stages.s3.data.forEach(r => {
      if (!isMatchUserType(r.suministro)) return;
      const m = getLoc(r.localidad);
      m.b3_tot.add(r.suministro);
      if (r.hasAP) m.b3_tot_ap.add(r.suministro);
      if (r.hasAL) m.b3_tot_al.add(r.suministro);

      if (r.isParaAnalizar) {
        if (r.isValid) {
          m.b3_val.add(r.suministro);
          if (r.hasAP) m.b3_val_ap.add(r.suministro);
          if (r.hasAL) m.b3_val_al.add(r.suministro);
        } else {
          m.b3_obs.add(r.suministro);
          if (r.hasAP) m.b3_obs_ap.add(r.suministro);
          if (r.hasAL) m.b3_obs_al.add(r.suministro);
        }
      }
    });

    // Populate Cumplen Meta from evaluatedRecords
    state.evaluatedRecords.forEach(r => {
      if (!isMatchUserType(r.suministro)) return;
      if (r.status === 'CUMPLE') {
        const m = getLoc(r.localidad);
        m.cumplen.add(r.suministro);
      }
    });

    // Grand Totals accumulator
    const g = {
      b1_ap: 0, b1_al: 0, b1_tot: 0,
      b2_tot_ap: 0, b2_tot_al: 0, b2_tot: 0,
      b2_obs_ap: 0, b2_obs_al: 0, b2_obs: 0,
      b2_val_ap: 0, b2_val_al: 0, b2_val: 0,
      b3_tot_ap: 0, b3_tot_al: 0, b3_tot: 0,
      b3_obs_ap: 0, b3_obs_al: 0, b3_obs: 0,
      b3_val_ap: 0, b3_val_al: 0, b3_val: 0,
      cumplen: 0
    };

    sortedLocs.forEach(loc => {
      const d = locMap[loc];
      const b1_ap = d.b1_ap.size;
      const b1_al = d.b1_al.size;
      const b1_tot = d.b1_tot.size;

      const b2_tot_ap = d.b2_tot_ap.size;
      const b2_tot_al = d.b2_tot_al.size;
      const b2_tot = d.b2_tot.size;

      const b2_obs_ap = d.b2_obs_ap.size;
      const b2_obs_al = d.b2_obs_al.size;
      const b2_obs = d.b2_obs.size;

      const b2_val_ap = d.b2_val_ap.size;
      const b2_val_al = d.b2_val_al.size;
      const b2_val = d.b2_val.size;

      const b3_tot_ap = d.b3_tot_ap.size;
      const b3_tot_al = d.b3_tot_al.size;
      const b3_tot = d.b3_tot.size;

      const b3_obs_ap = d.b3_obs_ap.size;
      const b3_obs_al = d.b3_obs_al.size;
      const b3_obs = d.b3_obs.size;

      const b3_val_ap = d.b3_val_ap.size;
      const b3_val_al = d.b3_val_al.size;
      const b3_val = d.b3_val.size;

      const cumple = d.cumplen.size;
      const pct = b1_tot > 0 ? ((cumple / b1_tot) * 100).toFixed(1) : 0;

      // Accumulate
      g.b1_ap += b1_ap; g.b1_al += b1_al; g.b1_tot += b1_tot;
      g.b2_tot_ap += b2_tot_ap; g.b2_tot_al += b2_tot_al; g.b2_tot += b2_tot;
      g.b2_obs_ap += b2_obs_ap; g.b2_obs_al += b2_obs_al; g.b2_obs += b2_obs;
      g.b2_val_ap += b2_val_ap; g.b2_val_al += b2_val_al; g.b2_val += b2_val;
      g.b3_tot_ap += b3_tot_ap; g.b3_tot_al += b3_tot_al; g.b3_tot += b3_tot;
      g.b3_obs_ap += b3_obs_ap; g.b3_obs_al += b3_obs_al; g.b3_obs += b3_obs;
      g.b3_val_ap += b3_val_ap; g.b3_val_al += b3_val_al; g.b3_val += b3_val;
      g.cumplen += cumple;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong style="color:var(--text-main); font-size:0.875rem;">${loc}</strong></td>

        <!-- Base 1 Conexiones Totales -->
        <td class="text-center font-mono">${b1_ap.toLocaleString()}</td>
        <td class="text-center font-mono">${b1_al.toLocaleString()}</td>
        <td class="text-center font-mono font-bold" style="background:rgba(2,132,199,0.06);">${b1_tot.toLocaleString()}</td>

        <!-- Base 2 Fichas Totales -->
        <td class="text-center font-mono">${b2_tot_ap.toLocaleString()}</td>
        <td class="text-center font-mono">${b2_tot_al.toLocaleString()}</td>
        <td class="text-center font-mono font-bold">${b2_tot.toLocaleString()}</td>
        <!-- Base 2 Fichas Observadas -->
        <td class="text-center font-mono text-amber">${b2_obs_ap.toLocaleString()}</td>
        <td class="text-center font-mono text-amber">${b2_obs_al.toLocaleString()}</td>
        <td class="text-center font-mono font-bold text-amber" style="background:rgba(245,158,11,0.06);">${b2_obs.toLocaleString()}</td>
        <!-- Base 2 Fichas Validadas -->
        <td class="text-center font-mono text-emerald">${b2_val_ap.toLocaleString()}</td>
        <td class="text-center font-mono text-emerald">${b2_val_al.toLocaleString()}</td>
        <td class="text-center font-mono font-bold text-emerald" style="background:rgba(16,185,129,0.06);">${b2_val.toLocaleString()}</td>

        <!-- Base 3 QGIS Totales -->
        <td class="text-center font-mono">${b3_tot_ap.toLocaleString()}</td>
        <td class="text-center font-mono">${b3_tot_al.toLocaleString()}</td>
        <td class="text-center font-mono font-bold">${b3_tot.toLocaleString()}</td>
        <!-- Base 3 QGIS Observados -->
        <td class="text-center font-mono text-amber">${b3_obs_ap.toLocaleString()}</td>
        <td class="text-center font-mono text-amber">${b3_obs_al.toLocaleString()}</td>
        <td class="text-center font-mono font-bold text-amber" style="background:rgba(245,158,11,0.06);">${b3_obs.toLocaleString()}</td>
        <!-- Base 3 QGIS Validados -->
        <td class="text-center font-mono text-emerald">${b3_val_ap.toLocaleString()}</td>
        <td class="text-center font-mono text-emerald">${b3_val_al.toLocaleString()}</td>
        <td class="text-center font-mono font-bold text-emerald" style="background:rgba(16,185,129,0.06);">${b3_val.toLocaleString()}</td>

        <!-- Cumplen Meta -->
        <td class="text-center text-emerald font-bold" style="font-size:0.95rem;">${cumple.toLocaleString()}</td>

        <!-- % Avance -->
        <td>
          <div class="locality-progress-wrap">
            <div class="progress-bar-bg" style="flex:1; margin-bottom:0; height:8px;">
              <div class="progress-bar-fill bg-cyan" style="width: ${Math.min(pct, 100)}%"></div>
            </div>
            <span class="locality-pct-text ${pct >= 100 ? 'text-emerald font-bold' : ''}">${pct}%</span>
          </div>
        </td>

        <!-- Acción -->
        <td class="text-center">
          <button class="btn btn-outline btn-filter-loc" data-locality="${loc}" title="Filtrar tabla detallada por esta localidad">
            <i data-lucide="filter"></i> Ver
          </button>
        </td>
      `;
      summaryBody.appendChild(tr);
    });

    // Grand Total Row in Footer
    if (summaryFooter) {
      const grandPct = g.b1_tot > 0 ? ((g.cumplen / g.b1_tot) * 100).toFixed(1) : 0;
      summaryFooter.innerHTML = `
        <tr class="tr-total-footer">
          <td><strong>Total General</strong></td>

          <!-- Bloque 1 Total -->
          <td class="text-center font-mono font-bold">${g.b1_ap.toLocaleString()}</td>
          <td class="text-center font-mono font-bold">${g.b1_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold" style="color:var(--primary); font-size:0.95rem;">${g.b1_tot.toLocaleString()}</td>

          <!-- Bloque 2 Fichas Totales -->
          <td class="text-center font-mono font-bold">${g.b2_tot_ap.toLocaleString()}</td>
          <td class="text-center font-mono font-bold">${g.b2_tot_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold" style="color:var(--text-main);">${g.b2_tot.toLocaleString()}</td>
          <!-- Bloque 2 Fichas Observadas -->
          <td class="text-center font-mono font-bold text-amber">${g.b2_obs_ap.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-amber">${g.b2_obs_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-amber" style="font-size:0.95rem;">${g.b2_obs.toLocaleString()}</td>
          <!-- Bloque 2 Fichas Validadas -->
          <td class="text-center font-mono font-bold text-emerald">${g.b2_val_ap.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-emerald">${g.b2_val_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-emerald" style="font-size:0.95rem;">${g.b2_val.toLocaleString()}</td>

          <!-- Bloque 3 QGIS Totales -->
          <td class="text-center font-mono font-bold">${g.b3_tot_ap.toLocaleString()}</td>
          <td class="text-center font-mono font-bold">${g.b3_tot_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold" style="color:var(--text-main);">${g.b3_tot.toLocaleString()}</td>
          <!-- Bloque 3 QGIS Observados -->
          <td class="text-center font-mono font-bold text-amber">${g.b3_obs_ap.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-amber">${g.b3_obs_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-amber" style="font-size:0.95rem;">${g.b3_obs.toLocaleString()}</td>
          <!-- Bloque 3 QGIS Validados -->
          <td class="text-center font-mono font-bold text-emerald">${g.b3_val_ap.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-emerald">${g.b3_val_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-emerald" style="font-size:0.95rem;">${g.b3_val.toLocaleString()}</td>

          <!-- Cumplimiento Total -->
          <td class="text-center text-emerald font-bold" style="font-size:1rem;">${g.cumplen.toLocaleString()}</td>
          <td class="text-center font-bold" style="font-size:0.95rem;">${grandPct}%</td>
          <td class="text-center">-</td>
        </tr>
      `;
    }

    if (window.lucide) window.lucide.createIcons();

    // Attach click listeners to "Ver" button
    summaryBody.querySelectorAll('.btn-filter-loc').forEach(btn => {
      btn.addEventListener('click', () => {
        const targetLoc = btn.dataset.locality;
        state.table.localityFilter = targetLoc;
        filterLocalitySelect.value = targetLoc;
        state.table.page = 1;
        filterAndPaginateTable();
        document.querySelector('.data-explorer-card').scrollIntoView({ behavior: 'smooth' });
        showToast(`Filtrando por localidad: ${targetLoc}`, 'info');
      });
    });
  }

  // Compute Statistics
  function computeStats() {
    const totalRecords = state.evaluatedRecords.length;
    const totalPadron = state.evaluatedRecords.filter(r => r.inBase1).length;
    const validCount = state.evaluatedRecords.filter(r => r.status === 'CUMPLE').length;
    const soloFichaCount = state.evaluatedRecords.filter(r => r.status === 'SOLO_FICHA').length;
    const soloGisCount = state.evaluatedRecords.filter(r => r.status === 'SOLO_GIS').length;
    const fueraAnoCount = state.evaluatedRecords.filter(r => r.status === 'FUERA_ANO').length;
    const noCatastradoCount = state.evaluatedRecords.filter(r => r.status === 'NO_CATASTRADO').length;
    const noPadronCount = state.evaluatedRecords.filter(r => r.status === 'NO_PADRON').length;

    const observedTotal = soloFichaCount + soloGisCount + fueraAnoCount;

    // La meta se determina en función de la Base 1 (Conexiones Totales)
    const target = totalPadron;
    const compliancePct = target > 0 ? ((validCount / target) * 100) : 0;
    const padronCoveragePct = totalPadron > 0 ? ((validCount / totalPadron) * 100) : 0;

    // Regulatory Service Breakdown (Criterio SUNASS: 2 Servicios Evaluados)
    // 1. Agua Potable: Comprende "Agua potable" + "Agua y Alcantarillado"
    const totalAgua = state.evaluatedRecords.filter(r => r.inBase1 && r.tieneAgua).length;
    const validAgua = state.evaluatedRecords.filter(r => r.status === 'CUMPLE' && r.tieneAgua).length;
    const obsAgua = state.evaluatedRecords.filter(r => (r.status === 'SOLO_FICHA' || r.status === 'SOLO_GIS' || r.status === 'FUERA_ANO') && r.tieneAgua).length;
    const pctAgua = totalAgua > 0 ? Number(((validAgua / totalAgua) * 100).toFixed(1)) : 0;

    // 2. Alcantarillado / Saneamiento: Comprende "Alcantarillado" + "Agua y Alcantarillado"
    const totalAlc = state.evaluatedRecords.filter(r => r.inBase1 && r.tieneAlcantarillado).length;
    const validAlc = state.evaluatedRecords.filter(r => r.status === 'CUMPLE' && r.tieneAlcantarillado).length;
    const obsAlc = state.evaluatedRecords.filter(r => (r.status === 'SOLO_FICHA' || r.status === 'SOLO_GIS' || r.status === 'FUERA_ANO') && r.tieneAlcantarillado).length;
    const pctAlc = totalAlc > 0 ? Number(((validAlc / totalAlc) * 100).toFixed(1)) : 0;

    // Segmentos puros
    const totalAmbos = state.evaluatedRecords.filter(r => r.inBase1 && r.tieneAgua && r.tieneAlcantarillado).length;
    const validAmbos = state.evaluatedRecords.filter(r => r.status === 'CUMPLE' && r.tieneAgua && r.tieneAlcantarillado).length;
    const totalSoloAgua = state.evaluatedRecords.filter(r => r.inBase1 && r.tieneAgua && !r.tieneAlcantarillado).length;
    const validSoloAgua = state.evaluatedRecords.filter(r => r.status === 'CUMPLE' && r.tieneAgua && !r.tieneAlcantarillado).length;
    const totalSoloAlc = state.evaluatedRecords.filter(r => r.inBase1 && !r.tieneAgua && r.tieneAlcantarillado).length;
    const validSoloAlc = state.evaluatedRecords.filter(r => r.status === 'CUMPLE' && !r.tieneAgua && r.tieneAlcantarillado).length;

    // 3. Distinción Regulatoria: Usuarios Antiguos vs Usuarios Nuevos
    const totalAntiguos = state.evaluatedRecords.filter(r => r.inBase1 && r.tipoUsuario === 'ANTIGUO').length;
    const validAntiguos = state.evaluatedRecords.filter(r => r.status === 'CUMPLE' && r.tipoUsuario === 'ANTIGUO').length;
    const obsAntiguos = state.evaluatedRecords.filter(r => (r.status === 'SOLO_FICHA' || r.status === 'SOLO_GIS' || r.status === 'FUERA_ANO') && r.tipoUsuario === 'ANTIGUO').length;
    const noCatAntiguos = state.evaluatedRecords.filter(r => r.status === 'NO_CATASTRADO' && r.tipoUsuario === 'ANTIGUO').length;
    const pctAntiguos = totalAntiguos > 0 ? Number(((validAntiguos / totalAntiguos) * 100).toFixed(1)) : 0;

    const totalNuevos = state.evaluatedRecords.filter(r => r.inBase1 && r.tipoUsuario === 'NUEVO').length;
    const validNuevos = state.evaluatedRecords.filter(r => r.status === 'CUMPLE' && r.tipoUsuario === 'NUEVO').length;
    const obsNuevos = state.evaluatedRecords.filter(r => (r.status === 'SOLO_FICHA' || r.status === 'SOLO_GIS' || r.status === 'FUERA_ANO') && r.tipoUsuario === 'NUEVO').length;
    const noCatNuevos = state.evaluatedRecords.filter(r => r.status === 'NO_CATASTRADO' && r.tipoUsuario === 'NUEVO').length;
    const pctNuevos = totalNuevos > 0 ? Number(((validNuevos / totalNuevos) * 100).toFixed(1)) : 0;

    state.stats = {
      totalRecords,
      totalPadron,
      validCount,
      soloFichaCount,
      soloGisCount,
      fueraAnoCount,
      observedTotal,
      noCatastradoCount,
      noPadronCount,
      target,
      compliancePct: Number(compliancePct.toFixed(1)),
      padronCoveragePct: Number(padronCoveragePct.toFixed(1)),
      // Regulatory Sunass breakdown
      totalAgua, validAgua, obsAgua, pctAgua,
      totalAlc, validAlc, obsAlc, pctAlc,
      totalAmbos, validAmbos,
      totalSoloAgua, validSoloAgua,
      totalSoloAlc, validSoloAlc,
      // Distinción Antiguos vs Nuevos
      totalAntiguos, validAntiguos, obsAntiguos, noCatAntiguos, pctAntiguos,
      totalNuevos, validNuevos, obsNuevos, noCatNuevos, pctNuevos
    };
  }

  // Render Result KPI Header and Cards
  function renderResults() {
    const s = state.stats;
    const setTxt = (id, val) => {
      const el = document.getElementById(id);
      if (el) el.textContent = val;
    };

    setTxt('res-year-label', state.evalYear);
    setTxt('res-eps-label', state.epsName);
    setTxt('res-total-evaluated', s.totalRecords.toLocaleString());

    // Update User Segmentation Labels and KPIs
    document.querySelectorAll('.lbl-eval-year').forEach(el => el.textContent = state.evalYear);
    setTxt('seg-count-all', s.totalRecords.toLocaleString());
    setTxt('seg-count-antiguos', s.totalAntiguos.toLocaleString());
    setTxt('seg-count-nuevos', s.totalNuevos.toLocaleString());

    setTxt('ut-antiguos-total', s.totalAntiguos.toLocaleString());
    setTxt('ut-antiguos-valid', s.validAntiguos.toLocaleString());
    setTxt('ut-antiguos-obs', s.obsAntiguos.toLocaleString());
    setTxt('ut-antiguos-pct', `${s.pctAntiguos}%`);

    setTxt('ut-nuevos-total', s.totalNuevos.toLocaleString());
    setTxt('ut-nuevos-valid', s.validNuevos.toLocaleString());
    setTxt('ut-nuevos-obs', s.obsNuevos.toLocaleString());
    setTxt('ut-nuevos-pct', `${s.pctNuevos}%`);

    setTxt('kpi-compliance-pct', `${s.compliancePct}%`);
    setTxt('kpi-goal-comparison', `Meta: ${s.target.toLocaleString()}`);
    const elBar = document.getElementById('kpi-progress-bar');
    if (elBar) elBar.style.width = `${Math.min(s.compliancePct, 100)}%`;
    setTxt('kpi-validated-count', s.validCount.toLocaleString());
    setTxt('kpi-target-total', s.target.toLocaleString());

    setTxt('kpi-total-connections', s.totalPadron.toLocaleString());

    setTxt('kpi-valid-count', s.validCount.toLocaleString());
    setTxt('kpi-valid-pct', `${s.padronCoveragePct}% del padrón`);

    setTxt('kpi-observed-count', s.observedTotal.toLocaleString());
    const obsPct = s.totalPadron > 0 ? ((s.observedTotal / s.totalPadron) * 100).toFixed(1) : 0;
    setTxt('kpi-observed-pct', `${obsPct}% del padrón`);

    setTxt('kpi-not-cadastral-count', s.noCatastradoCount.toLocaleString());
    const nocatPct = s.totalPadron > 0 ? ((s.noCatastradoCount / s.totalPadron) * 100).toFixed(1) : 0;
    setTxt('kpi-not-cadastral-pct', `${nocatPct}% del padrón`);

    // Filter Chips Counts
    setTxt('count-chip-all', s.totalRecords.toLocaleString());
    setTxt('count-chip-cumple', s.validCount.toLocaleString());
    setTxt('count-chip-soloficha', s.soloFichaCount.toLocaleString());
    setTxt('count-chip-sologis', s.soloGisCount.toLocaleString());
    setTxt('count-chip-fuera', s.fueraAnoCount.toLocaleString());
    setTxt('count-chip-nocat', s.noCatastradoCount.toLocaleString());

    // Goal Status Badge
    const goalBadge = document.getElementById('goal-status-badge');
    const goalText = document.getElementById('goal-status-text');
    if (goalBadge && goalText) {
      if (s.compliancePct >= 100) {
        goalBadge.className = 'status-indicator-badge bg-emerald-subtle text-emerald';
        goalText.textContent = '¡Meta Cumplida al 100%! Evaluado y Validado';
      } else if (s.compliancePct >= 75) {
        goalBadge.className = 'status-indicator-badge bg-blue-subtle text-blue';
        goalText.textContent = 'Buen Nivel de Avance en Meta';
      } else {
        goalBadge.className = 'status-indicator-badge bg-amber-subtle text-amber';
        goalText.textContent = 'Avance Parcial / Conexiones Pendientes';
      }
    }

    // Regulatory Service Breakdown UI (SUNASS)
    const elAguaTotal = document.getElementById('kpi-agua-total');
    if (elAguaTotal) {
      elAguaTotal.textContent = s.totalAgua.toLocaleString();
      setTxt('kpi-agua-valid', s.validAgua.toLocaleString());
      setTxt('kpi-agua-obs', s.obsAgua.toLocaleString());
      setTxt('kpi-agua-pct', `${s.pctAgua}%`);
      setTxt('kpi-agua-footer-pct', `${s.pctAgua}% (${s.validAgua.toLocaleString()} de ${s.totalAgua.toLocaleString()})`);
      const elAguaBar = document.getElementById('kpi-agua-bar');
      if (elAguaBar) elAguaBar.style.width = `${Math.min(s.pctAgua, 100)}%`;

      setTxt('kpi-alc-total', s.totalAlc.toLocaleString());
      setTxt('kpi-alc-valid', s.validAlc.toLocaleString());
      setTxt('kpi-alc-obs', s.obsAlc.toLocaleString());
      setTxt('kpi-alc-pct', `${s.pctAlc}%`);
      setTxt('kpi-alc-footer-pct', `${s.pctAlc}% (${s.validAlc.toLocaleString()} de ${s.totalAlc.toLocaleString()})`);
      const elAlcBar = document.getElementById('kpi-alc-bar');
      if (elAlcBar) elAlcBar.style.width = `${Math.min(s.pctAlc, 100)}%`;
    }

    if (window.lucide) window.lucide.createIcons();
  }

  // Render Charts with Chart.js (optional)
  function renderCharts() {
    if (!window.Chart) return;
    const donutCanvas = document.getElementById('chart-status-donut');
    if (!donutCanvas) return;
    const isDark = document.body.classList.contains('theme-dark');
    const textColor = isDark ? '#f8fafc' : '#0f172a';
    const gridColor = isDark ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.06)';

    const s = state.stats;

    // 1. Donut Chart: Validation Status
    const ctxDonut = document.getElementById('chart-status-donut').getContext('2d');
    if (state.charts.donut) state.charts.donut.destroy();

    const donutData = [
      s.validCount,
      s.soloFichaCount,
      s.soloGisCount,
      s.fueraAnoCount,
      s.noCatastradoCount
    ];

    const donutColors = [
      '#10b981', // Cumple
      '#f59e0b', // Solo Ficha
      '#8b5cf6', // Solo GIS
      '#ef4444', // Fuera de año
      '#94a3b8'  // No catastrado
    ];

    const donutLabels = [
      'Cumple Meta (Ficha + GIS)',
      'Solo Ficha Catastral',
      'Solo Registro GIS',
      'Fuera de Periodo Evaluado',
      'No Catastrado en Padrón'
    ];

    state.charts.donut = new Chart(ctxDonut, {
      type: 'doughnut',
      data: {
        labels: donutLabels,
        datasets: [{
          data: donutData,
          backgroundColor: donutColors,
          borderWidth: 2,
          borderColor: isDark ? '#111827' : '#ffffff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '68%',
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                const pct = total > 0 ? ((ctx.raw / total) * 100).toFixed(1) : 0;
                return ` ${ctx.label}: ${ctx.raw.toLocaleString()} (${pct}%)`;
              }
            }
          }
        }
      }
    });

    // Custom Legend for Donut
    const legendContainer = document.getElementById('donut-legend-container');
    legendContainer.innerHTML = '';
    donutLabels.forEach((label, idx) => {
      const chip = document.createElement('div');
      chip.className = 'legend-chip';
      chip.innerHTML = `<span class="legend-dot" style="background:${donutColors[idx]}"></span><span>${label}: <strong>${donutData[idx]}</strong></span>`;
      legendContainer.appendChild(chip);
    });

    // 2. Bar Chart: Regulatory Service Distribution (Regla SUNASS)
    // Agua Potable (comprende Agua + Mixto) vs Alcantarillado (comprende Alc. + Mixto) vs Mixto
    const ctxService = document.getElementById('chart-service-bar').getContext('2d');
    if (state.charts.service) state.charts.service.destroy();

    state.charts.service = new Chart(ctxService, {
      type: 'bar',
      data: {
        labels: [
          'Agua Potable\n(Agua + Mixto)', 
          'Alcantarillado\n(Alc. + Mixto)', 
          'Ambos Servicios\n(Mixto)'
        ],
        datasets: [
          {
            label: 'Total Padrón Base',
            data: [s.totalAgua, s.totalAlc, s.totalAmbos],
            backgroundColor: '#0284c7',
            borderRadius: 6
          },
          {
            label: 'Catastradas Válidas (Cumple)',
            data: [s.validAgua, s.validAlc, s.validAmbos],
            backgroundColor: '#10b981',
            borderRadius: 6
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            labels: { color: textColor, font: { family: 'Outfit', weight: 600 } }
          },
          tooltip: {
            callbacks: {
              afterTitle: (items) => {
                const idx = items[0].dataIndex;
                if (idx === 0) return 'Regla: Comprende "Agua potable" + "Agua y Alcantarillado"';
                if (idx === 1) return 'Regla: Comprende "Alcantarillado/Saneamiento" + "Agua y Alcantarillado"';
                return 'Conexiones con ambos servicios';
              }
            }
          }
        },
        scales: {
          x: {
            ticks: { color: textColor, font: { family: 'Outfit', size: 11, weight: 600 } },
            grid: { display: false }
          },
          y: {
            ticks: { color: textColor, precision: 0 },
            grid: { color: gridColor }
          }
        }
      }
    });

    // 3. Bar Chart: Locality Distribution (Total vs Validated)
    const ctxLocality = document.getElementById('chart-locality-bar').getContext('2d');
    if (state.charts.locality) state.charts.locality.destroy();

    const locMap = {};
    state.evaluatedRecords.forEach(r => {
      const loc = r.localidad || 'SIN LOCALIDAD';
      if (!locMap[loc]) locMap[loc] = { total: 0, valid: 0 };
      if (r.inBase1) locMap[loc].total++;
      if (r.status === 'CUMPLE') locMap[loc].valid++;
    });

    const locLabels = Object.keys(locMap).sort();
    const locTotals = locLabels.map(l => locMap[l].total);
    const locValids = locLabels.map(l => locMap[l].valid);

    state.charts.locality = new Chart(ctxLocality, {
      type: 'bar',
      data: {
        labels: locLabels,
        datasets: [
          {
            label: 'Total Padrón',
            data: locTotals,
            backgroundColor: '#8b5cf6',
            borderRadius: 6
          },
          {
            label: 'Catastradas Válidas',
            data: locValids,
            backgroundColor: '#10b981',
            borderRadius: 6
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            labels: { color: textColor, font: { family: 'Outfit', weight: 600 } }
          }
        },
        scales: {
          x: {
            ticks: { color: textColor },
            grid: { display: false }
          },
          y: {
            ticks: { color: textColor, precision: 0 },
            grid: { color: gridColor }
          }
        }
      }
    });

    // 4. Timeline Chart: Monthly Distribution in Evaluation Year
    const months = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Set', 'Oct', 'Nov', 'Dic'];
    const fichaMonthly = new Array(12).fill(0);
    const gisMonthly = new Array(12).fill(0);

    state.evaluatedRecords.forEach(r => {
      if (r.fichaDateObj && r.fichaDateObj.getFullYear() === state.evalYear) {
        fichaMonthly[r.fichaDateObj.getMonth()]++;
      }
      if (r.gisDateObj && r.gisDateObj.getFullYear() === state.evalYear) {
        gisMonthly[r.gisDateObj.getMonth()]++;
      }
    });

    const ctxTimeline = document.getElementById('chart-timeline-bar').getContext('2d');
    if (state.charts.timeline) state.charts.timeline.destroy();

    state.charts.timeline = new Chart(ctxTimeline, {
      type: 'bar',
      data: {
        labels: months,
        datasets: [
          {
            label: 'Fichas Catastrales en Campo',
            data: fichaMonthly,
            backgroundColor: 'rgba(16, 185, 129, 0.8)',
            borderColor: '#10b981',
            borderWidth: 1,
            borderRadius: 4
          },
          {
            label: 'Incorporadas en Plataforma GIS',
            data: gisMonthly,
            backgroundColor: 'rgba(59, 130, 246, 0.8)',
            borderColor: '#2563eb',
            borderWidth: 1,
            borderRadius: 4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            labels: { color: textColor, font: { family: 'Outfit', weight: 600 } }
          }
        },
        scales: {
          x: {
            ticks: { color: textColor },
            grid: { display: false }
          },
          y: {
            ticks: { color: textColor, precision: 0 },
            grid: { color: gridColor }
          }
        }
      }
    });
  }

  function updateChartsTheme() {
    renderCharts();
  }

  // Filter, Sort and Paginate Dynamic Table
  function filterAndPaginateTable() {
    let filtered = [...state.evaluatedRecords];

    // Status Filter
    if (state.table.statusFilter !== 'ALL') {
      filtered = filtered.filter(r => r.status === state.table.statusFilter);
    }

    // Locality Filter
    if (state.table.localityFilter !== 'ALL') {
      filtered = filtered.filter(r => r.localidad === state.table.localityFilter);
    }

    // User Type Filter (Usuarios Antiguos vs Usuarios Nuevos)
    if (state.table.userTypeFilter && state.table.userTypeFilter !== 'ALL') {
      filtered = filtered.filter(r => r.tipoUsuario === state.table.userTypeFilter);
    }

    // Service Filter (Criterio Regulatorio SUNASS)
    if (state.table.serviceFilter !== 'ALL') {
      const sf = state.table.serviceFilter;
      if (sf === 'Agua potable') {
        filtered = filtered.filter(r => r.servicioCatastro === 'Agua potable');
      } else if (sf === 'Alcantarillado') {
        filtered = filtered.filter(r => r.servicioCatastro === 'Alcantarillado');
      } else {
        filtered = filtered.filter(r => r.servicioCatastro === sf || r.servicioReal === sf);
      }
    }

    // Text Search
    if (state.table.search) {
      const q = state.table.search.toLowerCase();
      filtered = filtered.filter(r => 
        r.suministro.toLowerCase().includes(q) ||
        r.usuario.toLowerCase().includes(q) ||
        r.direccion.toLowerCase().includes(q) ||
        r.localidad.toLowerCase().includes(q) ||
        r.servicioReal.toLowerCase().includes(q) ||
        r.servicioCatastro.toLowerCase().includes(q) ||
        r.codigoCatastral.toLowerCase().includes(q) ||
        r.observation.toLowerCase().includes(q)
      );
    }

    // Sorting
    filtered.sort((a, b) => {
      let valA = a[state.table.sortBy] || '';
      let valB = b[state.table.sortBy] || '';
      if (typeof valA === 'string') valA = valA.toLowerCase();
      if (typeof valB === 'string') valB = valB.toLowerCase();
      if (valA < valB) return state.table.sortAsc ? -1 : 1;
      if (valA > valB) return state.table.sortAsc ? 1 : -1;
      return 0;
    });

    const totalFiltered = filtered.length;
    const totalPages = Math.max(1, Math.ceil(totalFiltered / state.table.pageSize));

    if (state.table.page > totalPages) state.table.page = totalPages;
    if (state.table.page < 1) state.table.page = 1;

    const startIdx = (state.table.page - 1) * state.table.pageSize;
    const pageRows = filtered.slice(startIdx, startIdx + state.table.pageSize);

    // Update table stats
    tableShowingCount.textContent = pageRows.length.toLocaleString();
    tableTotalCount.textContent = totalFiltered.toLocaleString();
    currentPageSpan.textContent = state.table.page;
    totalPagesSpan.textContent = totalPages;

    btnPagePrev.disabled = state.table.page <= 1;
    btnPageNext.disabled = state.table.page >= totalPages;

    renderTableRows(pageRows);
    renderPaginationButtons(totalPages);
  }

  function renderTableRows(rows) {
    tableBody.innerHTML = '';

    if (rows.length === 0) {
      tableBody.innerHTML = `
        <tr>
          <td colspan="10" class="text-center" style="padding: 36px; color: var(--text-muted);">
            <i data-lucide="search-x" style="width: 32px; height: 32px; margin-bottom: 8px; opacity: 0.5;"></i>
            <p>No se encontraron conexiones con los filtros seleccionados.</p>
          </td>
        </tr>
      `;
      if (window.lucide) window.lucide.createIcons();
      return;
    }

    rows.forEach(r => {
      const tr = document.createElement('tr');

      // Status pill class & label
      let pillClass = 'pill-nocat';
      let pillIcon = 'x-circle';
      if (r.status === 'CUMPLE') {
        pillClass = 'pill-cumple';
        pillIcon = 'check-circle-2';
      } else if (r.status === 'SOLO_FICHA' || r.status === 'SOLO_GIS') {
        pillClass = 'pill-soloficha';
        pillIcon = 'alert-triangle';
      } else if (r.status === 'FUERA_ANO') {
        pillClass = 'pill-fuera';
        pillIcon = 'clock-alert';
      }

      tr.innerHTML = `
        <td>
          <span class="cell-suministro">${r.suministro}</span>
        </td>
        <td>
          <div class="cell-user-title">${r.usuario || '<span class="text-muted">Sin nombre</span>'}</div>
          <div class="cell-user-sub">${r.direccion || '-'}</div>
        </td>
        <td>
          <span class="badge-tag" style="background:var(--bg-subtle); color:var(--text-main); font-size:0.75rem;">${r.localidad}</span>
        </td>
        <td class="text-center">
          ${r.tipoUsuario === 'NUEVO'
            ? `<span class="badge-tag badge-user-nuevo" title="Instalación en ${r.fechaInstalacion} (Año evaluado)"><i data-lucide="sparkles" style="width:11px;height:11px;vertical-align:middle;margin-right:2px;"></i> Nuevo</span>`
            : `<span class="badge-tag badge-user-antiguo" title="Instalación en ${r.fechaInstalacion} (Antes del año)"><i data-lucide="history" style="width:11px;height:11px;vertical-align:middle;margin-right:2px;"></i> Antiguo</span>`
          }
          <div style="font-size:0.675rem; color:var(--text-muted); margin-top:2px;">${r.fechaInstalacion || '-'}</div>
        </td>
        <td>
          <span class="badge-tag" style="background:var(--bg-subtle); color:var(--text-muted); font-size:0.75rem;" title="Servicio remitido en archivo base">${r.servicioReal}</span>
        </td>
        <td>
          ${r.servicioCatastro === 'Agua potable'
            ? '<span class="badge-tag" style="background:rgba(2,132,199,0.12); color:#0284c7; font-weight:700;"><i data-lucide="droplet" style="width:12px;height:12px;display:inline;"></i> Agua potable</span>'
            : '<span class="badge-tag" style="background:rgba(147,51,234,0.12); color:#9333ea; font-weight:700;"><i data-lucide="waves" style="width:12px;height:12px;display:inline;"></i> Alcantarillado</span>'}
        </td>
        <td class="text-center">
          ${r.inBase1 
            ? '<span class="check-badge check-ok"><i data-lucide="check"></i> Sí</span>' 
            : '<span class="check-badge check-alert"><i data-lucide="x"></i> No</span>'}
        </td>
        <td class="text-center">
          ${r.inBase2 
            ? (r.fichaInYear 
                ? `<span class="check-badge check-ok" title="En año evaluado"><i data-lucide="check-check"></i> ${r.fichaDate}</span>` 
                : `<span class="check-badge check-alert" title="Fuera de año evaluado"><i data-lucide="calendar-x"></i> ${r.fichaDate}</span>`)
            : '<span class="check-badge check-missing"><i data-lucide="minus"></i> No</span>'}
        </td>
        <td class="text-center">
          ${r.inBase3 
            ? (r.gisInYear 
                ? `<span class="check-badge check-ok" title="En año evaluado"><i data-lucide="check-check"></i> ${r.gisDate}</span>` 
                : `<span class="check-badge check-alert" title="Fuera de año evaluado"><i data-lucide="calendar-x"></i> ${r.gisDate}</span>`)
            : '<span class="check-badge check-missing"><i data-lucide="minus"></i> No</span>'}
        </td>
        <td class="text-center">
          <span class="status-pill ${pillClass}">
            <i data-lucide="${pillIcon}"></i> ${r.statusLabel}
          </span>
        </td>
        <td class="text-center">
          <button class="btn btn-outline btn-sm btn-inspect-audit" data-record-id="${r.id}" title="Ver auditoría de las 3 bases">
            <i data-lucide="eye"></i>
          </button>
        </td>
      `;

      tableBody.appendChild(tr);
    });

    if (window.lucide) window.lucide.createIcons();

    // Attach inspect buttons
    document.querySelectorAll('.btn-inspect-audit').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const recId = btn.dataset.recordId;
        openAuditModal(recId);
      });
    });
  }

  function renderPaginationButtons(totalPages) {
    const container = document.getElementById('pagination-numbers');
    container.innerHTML = '';
    const cur = state.table.page;

    let start = Math.max(1, cur - 2);
    let end = Math.min(totalPages, cur + 2);

    for (let i = start; i <= end; i++) {
      const btn = document.createElement('button');
      btn.className = `page-num-btn ${i === cur ? 'active' : ''}`;
      btn.textContent = i;
      btn.addEventListener('click', () => {
        state.table.page = i;
        filterAndPaginateTable();
      });
      container.appendChild(btn);
    }
  }

  // Table Search and Filter Event Handlers
  searchInput.addEventListener('input', (e) => {
    state.table.search = e.target.value.trim();
    state.table.page = 1;
    btnClearSearch.classList.toggle('hidden', !state.table.search);
    filterAndPaginateTable();
  });

  btnClearSearch.addEventListener('click', () => {
    searchInput.value = '';
    state.table.search = '';
    btnClearSearch.classList.add('hidden');
    state.table.page = 1;
    filterAndPaginateTable();
  });

  if (filterLocalitySelect) {
    filterLocalitySelect.addEventListener('change', (e) => {
      state.table.localityFilter = e.target.value;
      state.table.page = 1;
      filterAndPaginateTable();
    });
  }

  const filterUserTypeSelect = document.getElementById('filter-user-type-select');
  if (filterUserTypeSelect) {
    filterUserTypeSelect.addEventListener('change', (e) => {
      const val = e.target.value;
      state.table.userTypeFilter = val;
      state.userTypeFilter = val;
      document.querySelectorAll('.btn-user-segment').forEach(b => {
        b.classList.toggle('active', b.getAttribute('data-type') === val);
      });
      state.table.page = 1;
      filterAndPaginateTable();
      renderLocalitySummaryTable();
    });
  }

  document.querySelectorAll('.btn-user-segment').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.btn-user-segment').forEach(b => b.classList.remove('active'));
      const targetBtn = e.currentTarget;
      targetBtn.classList.add('active');
      const userType = targetBtn.getAttribute('data-type');
      state.userTypeFilter = userType;
      state.table.userTypeFilter = userType;
      if (filterUserTypeSelect) {
        filterUserTypeSelect.value = userType;
      }
      state.table.page = 1;
      filterAndPaginateTable();
      renderLocalitySummaryTable();
      showToast(`Visualizando: ${userType === 'ALL' ? 'Todos los usuarios' : (userType === 'ANTIGUO' ? 'Usuarios Antiguos' : 'Usuarios Nuevos')}`, 'info');
    });
  });

  filterServiceSelect.addEventListener('change', (e) => {
    state.table.serviceFilter = e.target.value;
    state.table.page = 1;
    filterAndPaginateTable();
  });

  pageSizeSelect.addEventListener('change', (e) => {
    state.table.pageSize = parseInt(e.target.value) || 30;
    state.table.page = 1;
    filterAndPaginateTable();
  });

  statusChips.forEach(chip => {
    chip.addEventListener('click', () => {
      statusChips.forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      state.table.statusFilter = chip.dataset.status;
      state.table.page = 1;
      filterAndPaginateTable();
    });
  });

  btnPagePrev.addEventListener('click', () => {
    if (state.table.page > 1) {
      state.table.page--;
      filterAndPaginateTable();
    }
  });

  btnPageNext.addEventListener('click', () => {
    state.table.page++;
    filterAndPaginateTable();
  });

  // Table Header Sorts
  document.querySelectorAll('.th-sortable').forEach(th => {
    th.addEventListener('click', () => {
      const field = th.dataset.sort;
      if (state.table.sortBy === field) {
        state.table.sortAsc = !state.table.sortAsc;
      } else {
        state.table.sortBy = field;
        state.table.sortAsc = true;
      }
      filterAndPaginateTable();
    });
  });

  // Supply Audit Modal
  function openAuditModal(recordId) {
    const rec = state.evaluatedRecords.find(r => r.id === recordId || r.suministro === recordId);
    if (!rec) return;

    modalContent.innerHTML = `
      <div style="margin-bottom: 20px;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
          <div>
            <h2 style="font-size:1.4rem; color:var(--primary); font-family:'Outfit'; margin-bottom:2px;">
              Suministro N° ${rec.suministro}
            </h2>
            <div style="display:flex; gap:8px; align-items:center;">
              <span class="badge-tag" style="background:var(--bg-subtle); color:var(--text-muted); font-size:0.75rem;">Real: ${rec.servicioReal}</span>
              <span class="badge-tag" style="background:${rec.servicioCatastro === 'Agua potable' ? 'rgba(2,132,199,0.12)' : 'rgba(147,51,234,0.12)'}; color:${rec.servicioCatastro === 'Agua potable' ? '#0284c7' : '#9333ea'}; font-weight:700; font-size:0.75rem;">Catastro: ${rec.servicioCatastro}</span>
            </div>
          </div>
          <span class="status-pill ${rec.status === 'CUMPLE' ? 'pill-cumple' : (rec.status === 'NO_CATASTRADO' ? 'pill-nocat' : 'pill-soloficha')}">
            ${rec.statusLabel}
          </span>
        </div>
        <p style="font-size:0.875rem; color:var(--text-muted);">
          <strong>Diagnóstico Regulatorio:</strong> ${rec.observation}
        </p>
      </div>

      <div class="audit-grid">
        
        <!-- Base 1 Details -->
        <div class="audit-source-card">
          <div class="audit-source-header">
            <span class="audit-source-title text-blue">
              <i data-lucide="database"></i> Base 1: Conexiones Totales
            </span>
            <span>${rec.inBase1 ? '<span class="check-badge check-ok"><i data-lucide="check"></i> Presente</span>' : '<span class="check-badge check-alert"><i data-lucide="x"></i> No Encontrado</span>'}</span>
          </div>
          <div class="audit-fields-list">
            <div><span class="audit-field-label">Usuario:</span> <span class="audit-field-val">${rec.usuario || '-'}</span></div>
            <div><span class="audit-field-label">Dirección:</span> <span class="audit-field-val">${rec.direccion || '-'}</span></div>
            <div><span class="audit-field-label">Localidad / Sector:</span> <span class="audit-field-val">${rec.localidad}</span></div>
            <div><span class="audit-field-label">Servicio Remitido:</span> <span class="audit-field-val">${rec.servicioReal}</span></div>
            <div><span class="audit-field-label">Evaluado Como:</span> <span class="audit-field-val">${rec.servicioCatastro}</span></div>
            <div><span class="audit-field-label">Fecha Instalación (Base 1):</span> <span class="audit-field-val">${rec.fechaInstalacion || rec.b1Date || '-'}</span></div>
            <div><span class="audit-field-label">Tipo de Usuario:</span> <span class="audit-field-val"><strong>${rec.tipoUsuarioLabel || (rec.tipoUsuario === 'NUEVO' ? 'Usuario Nuevo' : 'Usuario Antiguo')}</strong> (${rec.tipoUsuario === 'NUEVO' ? 'Instalado durante el año evaluado' : 'Instalado antes del año evaluado'})</span></div>
            <div><span class="audit-field-label">Estado Facturación:</span> <span class="audit-field-val">${rec.rawBase1 ? (rec.rawBase1['Estado de la facturacion'] || rec.rawBase1['Estado'] || 'ACTIVA') : '-'}</span></div>
          </div>
        </div>

        <!-- Base 2 Details -->
        <div class="audit-source-card">
          <div class="audit-source-header">
            <span class="audit-source-title text-emerald">
              <i data-lucide="clipboard-check"></i> Base 2: Fichas Catastrales en Campo
            </span>
            <span>${rec.inBase2 ? (rec.fichaInYear ? '<span class="check-badge check-ok"><i data-lucide="check-check"></i> En Año Evaluado</span>' : '<span class="check-badge check-alert"><i data-lucide="alert-circle"></i> Fuera de Año</span>') : '<span class="check-badge check-missing"><i data-lucide="x"></i> No Encontrado</span>'}</span>
          </div>
          <div class="audit-fields-list">
            <div><span class="audit-field-label">Fecha de Ficha:</span> <span class="audit-field-val">${rec.fichaDate}</span></div>
            <div><span class="audit-field-label">¿En Periodo ${state.evalYear}?:</span> <span class="audit-field-val">${rec.fichaInYear ? 'SÍ (Válido)' : 'NO (Observado)'}</span></div>
            <div><span class="audit-field-label">Servicio en Ficha:</span> <span class="audit-field-val">${rec.rawBase2 ? (rec.rawBase2['Tipo de servicio'] || rec.rawBase2['Servicio'] || 'No especificado') : '-'}</span></div>
            <div><span class="audit-field-label">Modalidad Catastro:</span> <span class="audit-field-val">${rec.rawBase2 ? (rec.rawBase2['Modalidad de catastro comercial'] || rec.rawBase2['Modalidad'] || 'ASIGNADO') : '-'}</span></div>
            <div><span class="audit-field-label">Condición:</span> <span class="audit-field-val">${rec.rawBase2 ? (rec.rawBase2['Condicion de la conexion'] || rec.rawBase2['Condicion catastral'] || 'ACTUALIZADO') : '-'}</span></div>
          </div>
        </div>

        <!-- Base 3 Details -->
        <div class="audit-source-card">
          <div class="audit-source-header">
            <span class="audit-source-title text-purple">
              <i data-lucide="map-pin"></i> Base 3: Tabla Atributiva QGIS
            </span>
            <span>${rec.inBase3 ? (rec.gisInYear ? '<span class="check-badge check-ok"><i data-lucide="check-check"></i> En Año Evaluado</span>' : '<span class="check-badge check-alert"><i data-lucide="alert-circle"></i> Fuera de Año</span>') : '<span class="check-badge check-missing"><i data-lucide="x"></i> No Encontrado</span>'}</span>
          </div>
          <div class="audit-fields-list">
            <div><span class="audit-field-label">Fecha Registro QGIS:</span> <span class="audit-field-val">${rec.gisDate}</span></div>
            <div><span class="audit-field-label">¿En Periodo ${state.evalYear}?:</span> <span class="audit-field-val">${rec.gisInYear ? 'SÍ (Válido)' : 'NO (Observado)'}</span></div>
            <div><span class="audit-field-label">Código Catastral:</span> <span class="audit-field-val">${rec.codigoCatastral || '-'}</span></div>
            <div><span class="audit-field-label">Servicio en QGIS:</span> <span class="audit-field-val">${rec.rawBase3 ? (rec.rawBase3['Tipo de servicio'] || rec.rawBase3['Servicio'] || 'No especificado') : '-'}</span></div>
            <div><span class="audit-field-label">Cartografía Digital:</span> <span class="audit-field-val">${rec.inBase3 ? 'GEORREFERENCIADO' : 'PENDIENTE'}</span></div>
          </div>
        </div>

      </div>
    `;

    detailModal.classList.remove('hidden');
    if (window.lucide) window.lucide.createIcons();
  }

  btnCloseModal.addEventListener('click', () => detailModal.classList.add('hidden'));
  btnModalDismiss.addEventListener('click', () => detailModal.classList.add('hidden'));
  detailModal.addEventListener('click', (e) => {
    if (e.target === detailModal) detailModal.classList.add('hidden');
  });

  // ==========================================
  // STAGES NAVIGATION & RENDERING (4 ETAPAS)
  // ==========================================

  function switchStage(stageNum) {
    state.activeStage = stageNum;
    [1, 2, 3, 4].forEach(n => {
      const tab = document.getElementById(`tab-stage-${n}`);
      const pane = document.getElementById(`stage-pane-${n}`);
      if (tab) tab.classList.toggle('active', n === stageNum);
      if (pane) pane.classList.toggle('hidden', n !== stageNum);
    });
    if (window.lucide) window.lucide.createIcons();
  }

  // Bind tab buttons
  [1, 2, 3, 4].forEach(n => {
    const tab = document.getElementById(`tab-stage-${n}`);
    if (tab) {
      tab.addEventListener('click', () => switchStage(n));
    }
  });

  function updateStageTabsStats() {
    const s1Stat = document.getElementById('stage1-tab-stat');
    const s2Stat = document.getElementById('stage2-tab-stat');
    const s3Stat = document.getElementById('stage3-tab-stat');
    const s4Stat = document.getElementById('stage4-tab-stat');

    if (s1Stat) s1Stat.textContent = `${(state.stages.s1.data?.length || 0).toLocaleString()} conexiones`;
    if (s2Stat) s2Stat.textContent = `${(state.stages.s2.data?.length || 0).toLocaleString()} fichas`;
    if (s3Stat) s3Stat.textContent = `${(state.stages.s3.data?.length || 0).toLocaleString()} en GIS`;
    if (s4Stat) s4Stat.textContent = `${(state.stats?.validCount || 0).toLocaleString()} cumplen meta`;
  }

  // ------------------------------------------
  // ETAPA 1: Conexiones Totales (Base 1)
  // ------------------------------------------
  function computeAndRenderStage1() {
    const b1 = state.bases[1].data;
    const m1 = state.mappings.base1;
    const items = [];
    const locMapTotal = {};
    const locMapAntiguos = {};
    const locMapNuevos = {};

    function createLocRecord() {
      return {
        rep_ap: new Set(), rep_al: new Set(), rep_ambos: new Set(),
        ap: new Set(), al: new Set(), total: new Set(),
        obs_ap: new Set(), obs_al: new Set(), obs: new Set(),
        val_ap: new Set(), val_al: new Set(), val: new Set()
      };
    }

    function addRecordToLocMap(targetMap, locName, sumVal, apFlag, alFlag, completeFlag) {
      if (!targetMap[locName]) {
        targetMap[locName] = createLocRecord();
      }
      const d = targetMap[locName];
      d.total.add(sumVal);
      if (apFlag) d.ap.add(sumVal);
      if (alFlag) d.al.add(sumVal);

      if (apFlag && alFlag) {
        d.rep_ambos.add(sumVal);
      } else if (apFlag) {
        d.rep_ap.add(sumVal);
      } else if (alFlag) {
        d.rep_al.add(sumVal);
      }

      if (completeFlag) {
        d.val.add(sumVal);
        if (apFlag) d.val_ap.add(sumVal);
        if (alFlag) d.val_al.add(sumVal);
      } else {
        d.obs.add(sumVal);
        if (apFlag) d.obs_ap.add(sumVal);
        if (alFlag) d.obs_al.add(sumVal);
      }
    }

    b1.forEach(row => {
      const sum = normalizeSuministro(row[m1.suministro]);
      if (!sum) return;

      // Check locality filter if active
      if (state.base1LocalityFilter && state.base1LocalityFilter !== 'ALL' && m1.localidad) {
        const rawL = row[m1.localidad];
        const normL = rawL ? String(rawL).trim().toUpperCase() : '';
        if (normL !== state.base1LocalityFilter) return;
      }

      const rawDate = m1.fecha ? row[m1.fecha] : null;
      const dateObj = rawDate ? parseDateFlexible(rawDate) : null;
      const userTypeInfo = classifyUserType(dateObj);

      const rawLoc = m1.localidad ? row[m1.localidad] : (row['Localidad'] || row['LOCALIDAD'] || row['Sector']);
      const localidad = rawLoc ? String(rawLoc).trim().toUpperCase() : 'SIN LOCALIDAD';
      const rawServ = m1.servicio ? row[m1.servicio] : 'Agua potable y alcantarillado';
      const serviciosCatastro = getRegulatoryServices(rawServ);
      const hasAP = serviciosCatastro.includes('Agua potable');
      const hasAL = serviciosCatastro.includes('Alcantarillado');
      const isComplete = !!sum && !!rawServ && localidad !== 'SIN LOCALIDAD';

      const item = {
        suministro: sum,
        usuario: m1.usuario ? (row[m1.usuario] || '') : (row['Usuario'] || row['Titular'] || ''),
        direccion: m1.direccion ? (row[m1.direccion] || '') : (row['Direccion'] || row['Predio'] || ''),
        localidad,
        servicioReal: String(rawServ).trim(),
        servicioCatastro: serviciosCatastro.join(' y '),
        hasAP,
        hasAL,
        fechaStr: dateObj ? formatDateDisplay(dateObj) : (rawDate ? String(rawDate) : '-'),
        fechaDateObj: dateObj,
        tipoUsuario: userTypeInfo.type,
        tipoUsuarioLabel: userTypeInfo.label,
        isAntiguo: userTypeInfo.isAntiguo,
        isNuevo: userTypeInfo.isNuevo,
        cumple: isComplete,
        obs: !isComplete ? 'Datos incompletos' : ''
      };
      items.push(item);

      addRecordToLocMap(locMapTotal, localidad, sum, hasAP, hasAL, isComplete);
      if (userTypeInfo.isAntiguo) {
        addRecordToLocMap(locMapAntiguos, localidad, sum, hasAP, hasAL, isComplete);
      } else if (userTypeInfo.isNuevo) {
        addRecordToLocMap(locMapNuevos, localidad, sum, hasAP, hasAL, isComplete);
      }
    });

    state.stages.s1.data = items;

    // Stage 1 KPIs
    const total = items.length;
    const valid = items.filter(i => i.cumple).length;
    const obs = total - valid;
    const pct = total > 0 ? ((valid / total) * 100).toFixed(1) : 0;
    const numLocs = Object.keys(locMapTotal).length;

    const elTotal = document.getElementById('s1-kpi-total');
    const elValid = document.getElementById('s1-kpi-valid');
    const elValidPct = document.getElementById('s1-kpi-valid-pct');
    const elObs = document.getElementById('s1-kpi-obs');
    const elObsPct = document.getElementById('s1-kpi-obs-pct');
    const elLocs = document.getElementById('s1-kpi-localities');

    if (elTotal) elTotal.textContent = total.toLocaleString();
    if (elValid) elValid.textContent = valid.toLocaleString();
    if (elValidPct) elValidPct.textContent = `${pct}%`;
    if (elObs) elObs.textContent = obs.toLocaleString();
    if (elObsPct) elObsPct.textContent = `${total > 0 ? ((obs / total) * 100).toFixed(1) : 0}%`;
    if (elLocs) elLocs.textContent = numLocs;

    // Helper to render any matrix table (Antiguos, Nuevos, Consolidado)
    function renderMatrixTableRows(targetLocMap, tbodyEl, tfootEl, localitiesList) {
      if (!tbodyEl) return;
      tbodyEl.innerHTML = '';
      let grandRepAP = 0;
      let grandRepAL = 0;
      let grandRepAmbos = 0;
      let grandAP = 0;
      let grandAL = 0;
      let grandTotal = 0;
      let grandObsAP = 0;
      let grandObsAL = 0;
      let grandObs = 0;
      let grandValAP = 0;
      let grandValAL = 0;
      let grandVal = 0;

      localitiesList.forEach(loc => {
        const d = targetLocMap[loc] || {
          rep_ap: new Set(), rep_al: new Set(), rep_ambos: new Set(),
          ap: new Set(), al: new Set(), total: new Set(),
          obs_ap: new Set(), obs_al: new Set(), obs: new Set(),
          val_ap: new Set(), val_al: new Set(), val: new Set()
        };

        const repApCount = d.rep_ap.size;
        const repAlCount = d.rep_al.size;
        const repAmbosCount = d.rep_ambos.size;

        const apCount = d.ap.size;
        const alCount = d.al.size;
        const totCount = d.total.size;

        const obsApCount = d.obs_ap.size;
        const obsAlCount = d.obs_al.size;
        const obsCount = d.obs.size;

        const valApCount = d.val_ap.size;
        const valAlCount = d.val_al.size;
        const valCount = d.val.size;

        grandRepAP += repApCount;
        grandRepAL += repAlCount;
        grandRepAmbos += repAmbosCount;

        grandAP += apCount;
        grandAL += alCount;
        grandTotal += totCount;

        grandObsAP += obsApCount;
        grandObsAL += obsAlCount;
        grandObs += obsCount;

        grandValAP += valApCount;
        grandValAL += valAlCount;
        grandVal += valCount;

        const tr = document.createElement('tr');
        tr.className = 'tr-img-row';
        tr.innerHTML = `
          <td class="cell-locality-img">${loc}</td>
          <!-- Tipo de servicio reportado -->
          <td class="text-center cell-val-num">${repApCount.toLocaleString()}</td>
          <td class="text-center cell-val-num">${repAlCount.toLocaleString()}</td>
          <td class="text-center cell-val-num col-divider-right">${repAmbosCount.toLocaleString()}</td>
          <!-- Conexiones totales a catastrar -->
          <td class="text-center cell-val-num">${apCount.toLocaleString()}</td>
          <td class="text-center cell-val-num">${alCount.toLocaleString()}</td>
          <td class="text-center cell-val-total col-divider-right">${totCount.toLocaleString()}</td>
          <!-- Conexiones observadas -->
          <td class="text-center cell-val-num text-amber">${obsApCount.toLocaleString()}</td>
          <td class="text-center cell-val-num text-amber">${obsAlCount.toLocaleString()}</td>
          <td class="text-center cell-val-total text-amber col-divider-right" style="background:rgba(245,158,11,0.06);">${obsCount.toLocaleString()}</td>
          <!-- Conexiones validadas -->
          <td class="text-center cell-val-num text-emerald">${valApCount.toLocaleString()}</td>
          <td class="text-center cell-val-num text-emerald">${valAlCount.toLocaleString()}</td>
          <td class="text-center cell-val-total text-emerald" style="background:rgba(16,185,129,0.06);">${valCount.toLocaleString()}</td>
        `;
        tbodyEl.appendChild(tr);
      });

      if (tfootEl) {
        tfootEl.innerHTML = `
          <tr class="tr-img-total">
            <td class="cell-locality-img">Total General</td>
            <!-- Tipo de servicio reportado -->
            <td class="text-center cell-val-num"><strong>${grandRepAP.toLocaleString()}</strong></td>
            <td class="text-center cell-val-num"><strong>${grandRepAL.toLocaleString()}</strong></td>
            <td class="text-center cell-val-num col-divider-right"><strong>${grandRepAmbos.toLocaleString()}</strong></td>
            <!-- Conexiones totales a catastrar -->
            <td class="text-center cell-val-num"><strong>${grandAP.toLocaleString()}</strong></td>
            <td class="text-center cell-val-num"><strong>${grandAL.toLocaleString()}</strong></td>
            <td class="text-center cell-img-grand-total col-divider-right">${grandTotal.toLocaleString()}</td>
            <!-- Conexiones observadas -->
            <td class="text-center cell-val-num text-amber"><strong>${grandObsAP.toLocaleString()}</strong></td>
            <td class="text-center cell-val-num text-amber"><strong>${grandObsAL.toLocaleString()}</strong></td>
            <td class="text-center cell-val-total text-amber col-divider-right" style="background:rgba(245,158,11,0.08); font-size:1.05rem;"><strong>${grandObs.toLocaleString()}</strong></td>
            <!-- Conexiones validadas -->
            <td class="text-center cell-val-num text-emerald"><strong>${grandValAP.toLocaleString()}</strong></td>
            <td class="text-center cell-val-num text-emerald"><strong>${grandValAL.toLocaleString()}</strong></td>
            <td class="text-center cell-val-total text-emerald" style="background:rgba(16,185,129,0.08); font-size:1.05rem;"><strong>${grandVal.toLocaleString()}</strong></td>
          </tr>
        `;
      }
    }

    const sortedLocs = Object.keys(locMapTotal).sort();

    // 1. Matriz de Conexiones - Usuarios Antiguos
    renderMatrixTableRows(
      locMapAntiguos,
      document.getElementById('s1-matrix-body-antiguos'),
      document.getElementById('s1-matrix-footer-antiguos'),
      sortedLocs
    );

    // 2. Matriz de Conexiones - Usuarios Nuevos
    renderMatrixTableRows(
      locMapNuevos,
      document.getElementById('s1-matrix-body-nuevos'),
      document.getElementById('s1-matrix-footer-nuevos'),
      sortedLocs
    );

    // 3. Matriz Consolidada Total (Padrón Base Completo)
    renderMatrixTableRows(
      locMapTotal,
      document.getElementById('s1-matrix-body'),
      document.getElementById('s1-matrix-footer'),
      sortedLocs
    );

    // Populate Stage 1 Locality Filter Dropdown
    const filterLocSel = document.getElementById('s1-filter-locality');
    if (filterLocSel) {
      filterLocSel.innerHTML = '<option value="ALL">Todas las localidades</option>';
      sortedLocs.forEach(loc => {
        const opt = document.createElement('option');
        opt.value = loc;
        opt.textContent = `${loc} (${locMapTotal[loc].total.size})`;
        filterLocSel.appendChild(opt);
      });
    }

    state.stages.s1.page = 1;
    filterAndRenderStage1Table();
  }

  function filterAndRenderStage1Table() {
    const s1 = state.stages.s1;
    let list = [...s1.data];

    if (s1.search) {
      const q = s1.search.toLowerCase();
      list = list.filter(i =>
        i.suministro.toLowerCase().includes(q) ||
        i.usuario.toLowerCase().includes(q) ||
        i.direccion.toLowerCase().includes(q) ||
        i.localidad.toLowerCase().includes(q)
      );
    }
    if (s1.localityFilter !== 'ALL') {
      list = list.filter(i => i.localidad === s1.localityFilter);
    }
    if (s1.userTypeFilter && s1.userTypeFilter !== 'ALL') {
      if (s1.userTypeFilter === 'ANTIGUO') {
        list = list.filter(i => i.isAntiguo);
      } else if (s1.userTypeFilter === 'NUEVO') {
        list = list.filter(i => i.isNuevo);
      }
    }
    if (s1.statusFilter === 'CUMPLE') {
      list = list.filter(i => i.cumple);
    } else if (s1.statusFilter === 'OBS') {
      list = list.filter(i => !i.cumple);
    }

    const total = list.length;
    const start = (s1.page - 1) * s1.pageSize;
    const pageRows = list.slice(start, start + s1.pageSize);

    const tbody = document.getElementById('s1-table-body');
    if (tbody) {
      tbody.innerHTML = '';
      if (pageRows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" class="text-center py-4 text-muted">No se encontraron conexiones para los filtros seleccionados</td></tr>';
      } else {
        pageRows.forEach(r => {
          const tr = document.createElement('tr');
          tr.innerHTML = `
            <td><span class="cell-suministro">${r.suministro}</span></td>
            <td><strong>${r.usuario || '-'}</strong></td>
            <td><small class="text-muted">${r.direccion || '-'}</small></td>
            <td><span class="badge-tag" style="background:var(--bg-subtle); color:var(--text-main); font-size:0.75rem;">${r.localidad}</span></td>
            <td><span class="badge" style="${r.isAntiguo ? 'background:#dbeafe; color:#1e40af; border:1px solid #bfdbfe;' : 'background:#d1fae5; color:#065f46; border:1px solid #a7f3d0;'} font-size:0.72rem; padding:2px 7px; border-radius:10px; font-weight:600;">${r.tipoUsuarioLabel}</span></td>
            <td><span class="badge-tag" style="background:var(--bg-subtle); color:var(--text-muted); font-size:0.75rem;">${r.servicioReal}</span></td>
            <td>${r.hasAP ? '<span class="text-blue font-semibold">Agua</span> ' : ''}${r.hasAL ? '<span class="text-purple font-semibold">Alcantarillado</span>' : ''}</td>
            <td>${r.fechaStr}</td>
            <td class="text-center">${r.cumple ? '<span class="status-pill pill-cumple"><i data-lucide="check-circle-2"></i> COMPLETO</span>' : '<span class="status-pill pill-fuera"><i data-lucide="alert-circle"></i> OBSERVADO</span>'}</td>
          `;
          tbody.appendChild(tr);
        });
      }
    }

    const pageInfo = document.getElementById('s1-page-info');
    const pageNum = document.getElementById('s1-page-num');
    const prevBtn = document.getElementById('s1-page-prev');
    const nextBtn = document.getElementById('s1-page-next');
    const totalPages = Math.ceil(total / s1.pageSize) || 1;

    if (pageInfo) pageInfo.textContent = `Mostrando ${pageRows.length} de ${total.toLocaleString()} conexiones (Pág. ${s1.page}/${totalPages})`;
    if (pageNum) pageNum.textContent = s1.page;
    if (prevBtn) prevBtn.disabled = s1.page <= 1;
    if (nextBtn) nextBtn.disabled = s1.page >= totalPages;

    if (window.lucide) window.lucide.createIcons();
  }

  // ------------------------------------------
  // ETAPA 2: Fichas Catastrales en Campo (Base 2)
  // ------------------------------------------
  function computeAndRenderStage2() {
    const b2 = state.bases[2].data;
    const m2 = state.mappings.base2;
    // Map with Base 1 items for exact user classification and validation status according strictly to Base 1:
    const b1Lookup = new Map();
    state.stages.s1.data.forEach(item => {
      b1Lookup.set(item.suministro, item);
    });

    const items = [];
    const locMapTotal = {};
    const locMapAntiguos = {};
    const locMapNuevos = {};

    let obsAnoCount = 0;
    let obsR1Count = 0;

    function createStage2LocRecord() {
      return {
        rem_ap: new Set(), rem_al: new Set(), rem: new Set(),
        ana_ap: new Set(), ana_al: new Set(), ana: new Set(),
        obs_ap: new Set(), obs_al: new Set(), obs: new Set(),
        val_ap: new Set(), val_al: new Set(), val: new Set()
      };
    }

    function addStage2Record(targetMap, locName, sumVal, apFlag, alFlag, isParaAnalizar, inYear) {
      if (!targetMap[locName]) {
        targetMap[locName] = createStage2LocRecord();
      }
      const lm = targetMap[locName];
      // 1. Total fichas catastrales remitidas (todas las enviadas por la EPS)
      lm.rem.add(sumVal);
      if (apFlag) lm.rem_ap.add(sumVal);
      if (alFlag) lm.rem_al.add(sumVal);

      // 2. Total fichas catastrales para analizar (suministros validados en RESULTADO 1)
      if (isParaAnalizar) {
        lm.ana.add(sumVal);
        if (apFlag) lm.ana_ap.add(sumVal);
        if (alFlag) lm.ana_al.add(sumVal);

        // 3. Fichas Observadas vs Validadas (según año regulatorio)
        if (inYear) {
          lm.val.add(sumVal);
          if (apFlag) lm.val_ap.add(sumVal);
          if (alFlag) lm.val_al.add(sumVal);
        } else {
          lm.obs.add(sumVal);
          if (apFlag) lm.obs_ap.add(sumVal);
          if (alFlag) lm.obs_al.add(sumVal);
        }
      }
    }

    b2.forEach(row => {
      const sum = normalizeSuministro(row[m2.suministro]);
      if (!sum) return;

      const rawDate = row[m2.fecha];
      const dateObj = parseDateFlexible(rawDate);
      const inYear = isDateInEvaluation(dateObj);
      const b1Info = b1Lookup.get(sum);
      const inBase1 = !!b1Info;
      // Validado en Resultado 1 (Base 1 con datos completos/cumple)
      const isValidInR1 = inBase1 && !!b1Info.cumple;
      const isParaAnalizar = isValidInR1;
      const isValid = isParaAnalizar && inYear;

      // Classification is based SOLELY on Base 1 installation date
      const isAntiguo = b1Info ? b1Info.isAntiguo : true;
      const isNuevo = b1Info ? b1Info.isNuevo : false;
      const tipoUsuario = b1Info ? b1Info.tipoUsuario : 'ANTIGUO';
      const tipoUsuarioLabel = b1Info ? b1Info.tipoUsuarioLabel : 'No en Base 1';

      let obsType = '';
      let obsLabel = '';
      if (!inBase1) {
        obsType = 'OBS_B1';
        obsLabel = 'No figura en Base 1 (Excluido de análisis)';
        obsR1Count++;
      } else if (!b1Info.cumple) {
        obsType = 'OBS_R1_INVALID';
        obsLabel = 'No validado en Resultado 1 (Excluido de análisis)';
        obsR1Count++;
      } else if (!inYear) {
        obsType = 'OBS_ANO';
        obsLabel = `Fecha fuera de año (${formatDateDisplay(dateObj)})`;
        obsAnoCount++;
      }

      const rawLoc = (m2.localidad ? row[m2.localidad] : null) || row['Localidad'] || row['LOCALIDAD'] || row['Sector'] || 'SIN LOCALIDAD';
      const localidad = String(rawLoc).trim().toUpperCase();
      const rawServ = m2.servicio ? (row[m2.servicio] || '') : 'Agua y alcantarillado';
      const serviciosCatastro = getRegulatoryServices(rawServ);
      const hasAP = serviciosCatastro.includes('Agua potable');
      const hasAL = serviciosCatastro.includes('Alcantarillado');
      const modalidad = row['Modalidad de catastro comercial'] || row['Modalidad'] || row[m2.estado] || 'CENSO';

      const item = {
        suministro: sum,
        localidad,
        fechaStr: formatDateDisplay(dateObj) || (rawDate ? String(rawDate) : '-'),
        inYear,
        inBase1,
        isValidInR1,
        isParaAnalizar,
        isValid,
        obsType,
        obsLabel,
        hasAP,
        hasAL,
        servicioRaw: String(rawServ).trim(),
        modalidad: String(modalidad).trim(),
        tipoUsuario,
        tipoUsuarioLabel,
        isAntiguo,
        isNuevo
      };
      items.push(item);

      addStage2Record(locMapTotal, localidad, sum, hasAP, hasAL, isParaAnalizar, inYear);
      if (isAntiguo) {
        addStage2Record(locMapAntiguos, localidad, sum, hasAP, hasAL, isParaAnalizar, inYear);
      }
      if (isNuevo) {
        addStage2Record(locMapNuevos, localidad, sum, hasAP, hasAL, isParaAnalizar, inYear);
      }
    });

    state.stages.s2.data = items;

    // Helper to render Stage 2 Matrix Rows
    function renderStage2MatrixTableRows(targetLocMap, tbodyEl, tfootEl, localitiesList) {
      if (!tbodyEl) return;
      tbodyEl.innerHTML = '';
      const grand = {
        rem_ap: 0, rem_al: 0, rem: 0,
        ana_ap: 0, ana_al: 0, ana: 0,
        obs_ap: 0, obs_al: 0, obs: 0,
        val_ap: 0, val_al: 0, val: 0
      };

      localitiesList.forEach(loc => {
        const d = targetLocMap[loc] || createStage2LocRecord();
        const rem_ap = d.rem_ap.size;
        const rem_al = d.rem_al.size;
        const rem = d.rem.size;

        const ana_ap = d.ana_ap.size;
        const ana_al = d.ana_al.size;
        const ana = d.ana.size;

        const obs_ap = d.obs_ap.size;
        const obs_al = d.obs_al.size;
        const obs = d.obs.size;

        const val_ap = d.val_ap.size;
        const val_al = d.val_al.size;
        const val = d.val.size;

        grand.rem_ap += rem_ap;
        grand.rem_al += rem_al;
        grand.rem += rem;

        grand.ana_ap += ana_ap;
        grand.ana_al += ana_al;
        grand.ana += ana;

        grand.obs_ap += obs_ap;
        grand.obs_al += obs_al;
        grand.obs += obs;

        grand.val_ap += val_ap;
        grand.val_al += val_al;
        grand.val += val;

        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><strong>${loc}</strong></td>
          <!-- Total Fichas Remitidas -->
          <td class="text-center font-mono">${rem_ap.toLocaleString()}</td>
          <td class="text-center font-mono">${rem_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold">${rem.toLocaleString()}</td>
          <!-- Total Fichas Para Analizar -->
          <td class="text-center font-mono text-primary font-semibold">${ana_ap.toLocaleString()}</td>
          <td class="text-center font-mono text-primary font-semibold">${ana_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-primary">${ana.toLocaleString()}</td>
          <!-- Fichas Observadas -->
          <td class="text-center font-mono text-amber">${obs_ap.toLocaleString()}</td>
          <td class="text-center font-mono text-amber">${obs_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-amber">${obs.toLocaleString()}</td>
          <!-- Fichas Validadas -->
          <td class="text-center font-mono text-emerald">${val_ap.toLocaleString()}</td>
          <td class="text-center font-mono text-emerald">${val_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-emerald">${val.toLocaleString()}</td>
        `;
        tbodyEl.appendChild(tr);
      });

      if (tfootEl) {
        tfootEl.innerHTML = `
          <tr class="tr-total-footer">
            <td><strong>Total General</strong></td>
            <td class="text-center font-mono font-bold">${grand.rem_ap.toLocaleString()}</td>
            <td class="text-center font-mono font-bold">${grand.rem_al.toLocaleString()}</td>
            <td class="text-center font-mono font-bold">${grand.rem.toLocaleString()}</td>

            <td class="text-center font-mono font-bold text-primary">${grand.ana_ap.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-primary">${grand.ana_al.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-primary">${grand.ana.toLocaleString()}</td>

            <td class="text-center font-mono font-bold text-amber">${grand.obs_ap.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-amber">${grand.obs_al.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-amber">${grand.obs.toLocaleString()}</td>

            <td class="text-center font-mono font-bold text-emerald">${grand.val_ap.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-emerald">${grand.val_al.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-emerald">${grand.val.toLocaleString()}</td>
          </tr>
        `;
      }
    }

    const sortedLocs = Object.keys(locMapTotal).sort();

    // 1. Matriz de Fichas - Usuarios Antiguos
    renderStage2MatrixTableRows(
      locMapAntiguos,
      document.getElementById('s2-matrix-body-antiguos'),
      document.getElementById('s2-matrix-footer-antiguos'),
      sortedLocs
    );

    // 2. Matriz de Fichas - Usuarios Nuevos
    renderStage2MatrixTableRows(
      locMapNuevos,
      document.getElementById('s2-matrix-body-nuevos'),
      document.getElementById('s2-matrix-footer-nuevos'),
      sortedLocs
    );

    // 3. Matriz Consolidada Total
    renderStage2MatrixTableRows(
      locMapTotal,
      document.getElementById('s2-matrix-body'),
      document.getElementById('s2-matrix-footer'),
      sortedLocs
    );

    // Populate Stage 2 Locality Filter
    const filterLocSel = document.getElementById('s2-filter-locality');
    if (filterLocSel) {
      filterLocSel.innerHTML = '<option value="ALL">Todas las localidades</option>';
      sortedLocs.forEach(loc => {
        const opt = document.createElement('option');
        opt.value = loc;
        opt.textContent = `${loc} (${locMapTotal[loc].rem.size} fichas)`;
        filterLocSel.appendChild(opt);
      });
    }

    state.stages.s2.page = 1;
    filterAndRenderStage2Table();
  }

  function filterAndRenderStage2Table() {
    const s2 = state.stages.s2;
    let list = [...s2.data];

    if (s2.search) {
      const q = s2.search.toLowerCase();
      list = list.filter(i =>
        i.suministro.toLowerCase().includes(q) ||
        i.localidad.toLowerCase().includes(q)
      );
    }
    if (s2.localityFilter !== 'ALL') {
      list = list.filter(i => i.localidad === s2.localityFilter);
    }
    if (s2.userTypeFilter && s2.userTypeFilter !== 'ALL') {
      if (s2.userTypeFilter === 'ANTIGUO') {
        list = list.filter(i => i.isAntiguo);
      } else if (s2.userTypeFilter === 'NUEVO') {
        list = list.filter(i => i.isNuevo);
      }
    }
    if (s2.statusFilter === 'ANALIZAR') {
      list = list.filter(i => i.isParaAnalizar);
    } else if (s2.statusFilter === 'VALID') {
      list = list.filter(i => i.isValid);
    } else if (s2.statusFilter === 'OBS_ALL') {
      list = list.filter(i => i.isParaAnalizar && !i.isValid);
    } else if (s2.statusFilter === 'OBS_ANO') {
      list = list.filter(i => i.obsType === 'OBS_ANO');
    } else if (s2.statusFilter === 'EXCLUIDAS') {
      list = list.filter(i => !i.isParaAnalizar);
    }

    const total = list.length;
    const start = (s2.page - 1) * s2.pageSize;
    const pageRows = list.slice(start, start + s2.pageSize);

    const tbody = document.getElementById('s2-table-body');
    if (tbody) {
      tbody.innerHTML = '';
      if (pageRows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" class="text-center py-4 text-muted">No se encontraron fichas para los filtros seleccionados</td></tr>';
      } else {
        pageRows.forEach(r => {
          const tr = document.createElement('tr');
          let statusHtml = '';
          if (r.isValid) {
            statusHtml = '<span class="status-pill pill-cumple"><i data-lucide="check-circle-2"></i> VALIDADA</span>';
          } else if (!r.isParaAnalizar) {
            statusHtml = '<span class="status-pill pill-fuera" style="background:rgba(225,29,72,0.12); color:#e11d48;"><i data-lucide="alert-octagon"></i> EXCLUIDA (NO EN R1)</span>';
          } else {
            statusHtml = '<span class="status-pill pill-soloficha"><i data-lucide="calendar-x"></i> OBSERVADA (FUERA AÑO)</span>';
          }

          tr.innerHTML = `
            <td><span class="cell-suministro">${r.suministro}</span></td>
            <td><span class="badge-tag" style="background:var(--bg-subtle); color:var(--text-main); font-size:0.75rem;">${r.localidad}</span></td>
            <td><span class="badge" style="${r.isAntiguo ? 'background:#dbeafe; color:#1e40af; border:1px solid #bfdbfe;' : 'background:#d1fae5; color:#065f46; border:1px solid #a7f3d0;'} font-size:0.72rem; padding:2px 7px; border-radius:10px; font-weight:600;">${r.tipoUsuarioLabel}</span></td>
            <td><strong>${r.fechaStr}</strong></td>
            <td class="text-center">${r.isParaAnalizar ? '<span class="check-badge check-ok"><i data-lucide="check"></i> SÍ (Validado R1)</span>' : '<span class="check-badge check-alert"><i data-lucide="x"></i> NO (Excluido)</span>'}</td>
            <td class="text-center">${r.inYear ? `<span class="check-badge check-ok"><i data-lucide="check-check"></i> SÍ (${state.evalYear})</span>` : `<span class="check-badge check-alert"><i data-lucide="x"></i> NO</span>`}</td>
            <td><small class="text-muted">${r.servicioRaw || '-'}</small></td>
            <td><small>${r.modalidad || '-'}</small></td>
            <td class="text-center">${statusHtml}</td>
          `;
          tbody.appendChild(tr);
        });
      }
    }

    const pageInfo = document.getElementById('s2-page-info');
    const pageNum = document.getElementById('s2-page-num');
    const prevBtn = document.getElementById('s2-page-prev');
    const nextBtn = document.getElementById('s2-page-next');
    const totalPages = Math.ceil(total / s2.pageSize) || 1;

    if (pageInfo) pageInfo.textContent = `Mostrando ${pageRows.length} de ${total.toLocaleString()} fichas (Pág. ${s2.page}/${totalPages})`;
    if (pageNum) pageNum.textContent = s2.page;
    if (prevBtn) prevBtn.disabled = s2.page <= 1;
    if (nextBtn) nextBtn.disabled = s2.page >= totalPages;

    if (window.lucide) window.lucide.createIcons();
  }

  // ------------------------------------------
  // ------------------------------------------
  // ETAPA 3: Plataforma QGIS (Base 3)
  // ------------------------------------------
  function computeAndRenderStage3() {
    const b3 = state.bases[3].data;
    const m3 = state.mappings.base3;
    // Suministros validados en Base 2 (fichas válidas en año y en Base 1 / Resultado 1)
    const validFichasSet = new Set(state.stages.s2.data.filter(i => i.isValid).map(i => i.suministro));

    // Lookup de Base 1 para clasificar en Antiguos y Nuevos por fecha de instalación de conexión
    const b1Lookup = new Map();
    state.stages.s1.data.forEach(r => {
      b1Lookup.set(r.suministro, r);
    });

    const items = [];
    const locMapTotal = {};
    const locMapAntiguos = {};
    const locMapNuevos = {};
    let obsAnoCount = 0;
    let obsFichasCount = 0;

    function createStage3LocRecord() {
      return {
        rem_ap: new Set(), rem_al: new Set(), rem: new Set(),
        ana_ap: new Set(), ana_al: new Set(), ana: new Set(),
        obs_ap: new Set(), obs_al: new Set(), obs: new Set(),
        val_ap: new Set(), val_al: new Set(), val: new Set()
      };
    }

    function addStage3Record(targetMap, locName, sumVal, apFlag, alFlag, isParaAnalizar, inYear) {
      if (!targetMap[locName]) {
        targetMap[locName] = createStage3LocRecord();
      }
      const lm = targetMap[locName];
      // 1. Total registros GIS remitidos (universo bruto enviado por la EPS)
      lm.rem.add(sumVal);
      if (apFlag) lm.rem_ap.add(sumVal);
      if (alFlag) lm.rem_al.add(sumVal);

      // 2. Total registros GIS para analizar (suministros validados en RESULTADO 2)
      if (isParaAnalizar) {
        lm.ana.add(sumVal);
        if (apFlag) lm.ana_ap.add(sumVal);
        if (alFlag) lm.ana_al.add(sumVal);

        // 3. Registros GIS Observados vs Validados (según año regulatorio)
        if (inYear) {
          lm.val.add(sumVal);
          if (apFlag) lm.val_ap.add(sumVal);
          if (alFlag) lm.val_al.add(sumVal);
        } else {
          lm.obs.add(sumVal);
          if (apFlag) lm.obs_ap.add(sumVal);
          if (alFlag) lm.obs_al.add(sumVal);
        }
      }
    }

    b3.forEach(row => {
      const sum = normalizeSuministro(row[m3.suministro]);
      if (!sum) return;

      const rawDate = row[m3.fecha];
      const dateObj = parseDateFlexible(rawDate);
      const inYear = isDateInEvaluation(dateObj);
      const inValidFichas = validFichasSet.has(sum);
      // Para analizar: registros de Base 3 validados previamente en Resultado 2
      const isParaAnalizar = inValidFichas;
      const isValid = isParaAnalizar && inYear;

      let obsType = '';
      let obsLabel = '';
      if (!isParaAnalizar) {
        obsType = 'OBS_FICHA';
        obsLabel = 'No validado en Fichas Catastrales (Resultado 2) - Excluido de análisis';
        obsFichasCount++;
      } else if (!inYear) {
        obsType = 'OBS_ANO';
        obsLabel = `Fecha fuera de año (${formatDateDisplay(dateObj)})`;
        obsAnoCount++;
      }

      const codigo = m3.codigo ? (row[m3.codigo] || '') : '';
      const rawLoc = (m3.localidad ? row[m3.localidad] : null) || row['Localidad'] || row['LOCALIDAD'] || row['Sector'] || 'SIN LOCALIDAD';
      const localidad = String(rawLoc).trim().toUpperCase();
      const rawServ = m3.servicio ? (row[m3.servicio] || '') : 'Agua y alcantarillado';
      const serviciosCatastro = getRegulatoryServices(rawServ);
      const hasAP = serviciosCatastro.includes('Agua potable');
      const hasAL = serviciosCatastro.includes('Alcantarillado');

      const b1Info = b1Lookup.get(sum);
      // Clasificación de usuarios antiguos y nuevos estrictamente según Base 1
      const isAntiguo = b1Info ? b1Info.isAntiguo : true;
      const isNuevo = b1Info ? b1Info.isNuevo : false;
      const tipoUsuario = b1Info ? b1Info.tipoUsuario : 'ANTIGUO';
      const tipoUsuarioLabel = b1Info ? (b1Info.isNuevo ? 'Nuevo' : 'Antiguo') : 'No en Base 1';

      const item = {
        suministro: sum,
        codigo,
        localidad,
        fechaStr: formatDateDisplay(dateObj) || (rawDate ? String(rawDate) : '-'),
        inYear,
        inValidFichas,
        isParaAnalizar,
        isValid,
        obsType,
        obsLabel,
        hasAP,
        hasAL,
        servicioRaw: String(rawServ).trim(),
        tipoUsuario,
        tipoUsuarioLabel,
        isAntiguo,
        isNuevo
      };
      items.push(item);

      addStage3Record(locMapTotal, localidad, sum, hasAP, hasAL, isParaAnalizar, inYear);
      if (isAntiguo) {
        addStage3Record(locMapAntiguos, localidad, sum, hasAP, hasAL, isParaAnalizar, inYear);
      }
      if (isNuevo) {
        addStage3Record(locMapNuevos, localidad, sum, hasAP, hasAL, isParaAnalizar, inYear);
      }
    });

    state.stages.s3.data = items;

    // Stage 3 KPIs
    const total = items.length;
    const valid = items.filter(i => i.isValid).length;
    const pct = total > 0 ? ((valid / total) * 100).toFixed(1) : 0;

    const elTotal = document.getElementById('s3-kpi-total');
    const elValid = document.getElementById('s3-kpi-valid');
    const elValidPct = document.getElementById('s3-kpi-valid-pct');
    const elObsAno = document.getElementById('s3-kpi-obs-ano');
    const elObsFichas = document.getElementById('s3-kpi-obs-fichas');

    if (elTotal) elTotal.textContent = total.toLocaleString();
    if (elValid) elValid.textContent = valid.toLocaleString();
    if (elValidPct) elValidPct.textContent = `${pct}%`;
    if (elObsAno) elObsAno.textContent = obsAnoCount.toLocaleString();
    if (elObsFichas) elObsFichas.textContent = obsFichasCount.toLocaleString();

    // Helper to render Stage 3 Matrix Rows
    function renderStage3MatrixTableRows(targetLocMap, tbodyEl, tfootEl, localitiesList) {
      if (!tbodyEl) return;
      tbodyEl.innerHTML = '';
      const grand = {
        rem_ap: 0, rem_al: 0, rem: 0,
        ana_ap: 0, ana_al: 0, ana: 0,
        obs_ap: 0, obs_al: 0, obs: 0,
        val_ap: 0, val_al: 0, val: 0
      };

      localitiesList.forEach(loc => {
        const d = targetLocMap[loc] || createStage3LocRecord();
        const rem_ap = d.rem_ap.size;
        const rem_al = d.rem_al.size;
        const rem = d.rem.size;

        const ana_ap = d.ana_ap.size;
        const ana_al = d.ana_al.size;
        const ana = d.ana.size;

        const obs_ap = d.obs_ap.size;
        const obs_al = d.obs_al.size;
        const obs = d.obs.size;

        const val_ap = d.val_ap.size;
        const val_al = d.val_al.size;
        const val = d.val.size;

        grand.rem_ap += rem_ap;
        grand.rem_al += rem_al;
        grand.rem += rem;

        grand.ana_ap += ana_ap;
        grand.ana_al += ana_al;
        grand.ana += ana;

        grand.obs_ap += obs_ap;
        grand.obs_al += obs_al;
        grand.obs += obs;

        grand.val_ap += val_ap;
        grand.val_al += val_al;
        grand.val += val;

        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><strong>${loc}</strong></td>
          <!-- Total Registros GIS Remitidos -->
          <td class="text-center font-mono">${rem_ap.toLocaleString()}</td>
          <td class="text-center font-mono">${rem_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold">${rem.toLocaleString()}</td>
          <!-- Total Registros GIS Para Analizar -->
          <td class="text-center font-mono text-primary font-semibold">${ana_ap.toLocaleString()}</td>
          <td class="text-center font-mono text-primary font-semibold">${ana_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-primary">${ana.toLocaleString()}</td>
          <!-- Registros GIS Observados -->
          <td class="text-center font-mono text-amber">${obs_ap.toLocaleString()}</td>
          <td class="text-center font-mono text-amber">${obs_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-amber">${obs.toLocaleString()}</td>
          <!-- Registros GIS Validados -->
          <td class="text-center font-mono text-emerald">${val_ap.toLocaleString()}</td>
          <td class="text-center font-mono text-emerald">${val_al.toLocaleString()}</td>
          <td class="text-center font-mono font-bold text-emerald">${val.toLocaleString()}</td>
        `;
        tbodyEl.appendChild(tr);
      });

      if (tfootEl) {
        tfootEl.innerHTML = `
          <tr class="tr-total-footer">
            <td><strong>Total General</strong></td>
            <td class="text-center font-mono font-bold">${grand.rem_ap.toLocaleString()}</td>
            <td class="text-center font-mono font-bold">${grand.rem_al.toLocaleString()}</td>
            <td class="text-center font-mono font-bold">${grand.rem.toLocaleString()}</td>

            <td class="text-center font-mono font-bold text-primary">${grand.ana_ap.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-primary">${grand.ana_al.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-primary">${grand.ana.toLocaleString()}</td>

            <td class="text-center font-mono font-bold text-amber">${grand.obs_ap.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-amber">${grand.obs_al.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-amber">${grand.obs.toLocaleString()}</td>

            <td class="text-center font-mono font-bold text-emerald">${grand.val_ap.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-emerald">${grand.val_al.toLocaleString()}</td>
            <td class="text-center font-mono font-bold text-emerald">${grand.val.toLocaleString()}</td>
          </tr>
        `;
      }
    }

    const sortedLocs = Object.keys(locMapTotal).sort();

    // 1. Matriz QGIS - Usuarios Antiguos
    renderStage3MatrixTableRows(
      locMapAntiguos,
      document.getElementById('s3-matrix-body-antiguos'),
      document.getElementById('s3-matrix-footer-antiguos'),
      sortedLocs
    );

    // 2. Matriz QGIS - Usuarios Nuevos
    renderStage3MatrixTableRows(
      locMapNuevos,
      document.getElementById('s3-matrix-body-nuevos'),
      document.getElementById('s3-matrix-footer-nuevos'),
      sortedLocs
    );

    // 3. Matriz Consolidada Total
    renderStage3MatrixTableRows(
      locMapTotal,
      document.getElementById('s3-matrix-body'),
      document.getElementById('s3-matrix-footer'),
      sortedLocs
    );

    // Populate Stage 3 Locality Filter
    const filterLocSel = document.getElementById('s3-filter-locality');
    if (filterLocSel) {
      filterLocSel.innerHTML = '<option value="ALL">Todas las localidades</option>';
      sortedLocs.forEach(loc => {
        const opt = document.createElement('option');
        opt.value = loc;
        opt.textContent = `${loc} (${locMapTotal[loc].rem.size})`;
        filterLocSel.appendChild(opt);
      });
    }

    state.stages.s3.page = 1;
    filterAndRenderStage3Table();
  }

  function filterAndRenderStage3Table() {
    const s3 = state.stages.s3;
    let list = [...s3.data];

    if (s3.search) {
      const q = s3.search.toLowerCase();
      list = list.filter(i =>
        i.suministro.toLowerCase().includes(q) ||
        i.codigo.toLowerCase().includes(q) ||
        i.localidad.toLowerCase().includes(q)
      );
    }
    if (s3.localityFilter !== 'ALL') {
      list = list.filter(i => i.localidad === s3.localityFilter);
    }
    if (s3.userTypeFilter && s3.userTypeFilter !== 'ALL') {
      if (s3.userTypeFilter === 'ANTIGUO') {
        list = list.filter(i => i.isAntiguo);
      } else if (s3.userTypeFilter === 'NUEVO') {
        list = list.filter(i => i.isNuevo);
      }
    }
    if (s3.statusFilter === 'ANALIZAR') {
      list = list.filter(i => i.isParaAnalizar);
    } else if (s3.statusFilter === 'VALID') {
      list = list.filter(i => i.isValid);
    } else if (s3.statusFilter === 'OBS_ALL') {
      list = list.filter(i => i.isParaAnalizar && !i.isValid);
    } else if (s3.statusFilter === 'OBS_ANO') {
      list = list.filter(i => i.isParaAnalizar && !i.inYear);
    } else if (s3.statusFilter === 'EXCLUIDAS') {
      list = list.filter(i => !i.isParaAnalizar);
    }

    const total = list.length;
    const start = (s3.page - 1) * s3.pageSize;
    const pageRows = list.slice(start, start + s3.pageSize);

    const tbody = document.getElementById('s3-table-body');
    if (tbody) {
      tbody.innerHTML = '';
      if (pageRows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" class="text-center py-4 text-muted">No se encontraron registros GIS para los filtros seleccionados</td></tr>';
      } else {
        pageRows.forEach(r => {
          const tr = document.createElement('tr');
          let statusHtml = '';
          if (r.isValid) {
            statusHtml = '<span class="status-pill pill-cumple"><i data-lucide="check-circle-2"></i> VALIDADO</span>';
          } else if (!r.isParaAnalizar) {
            statusHtml = '<span class="status-pill pill-fuera" style="background:rgba(225,29,72,0.12); color:#e11d48;"><i data-lucide="shield-alert"></i> EXCLUIDO (NO EN R2)</span>';
          } else {
            statusHtml = '<span class="status-pill pill-soloficha"><i data-lucide="calendar-x"></i> OBSERVADO (FUERA AÑO)</span>';
          }

          tr.innerHTML = `
            <td><span class="cell-suministro">${r.suministro}</span></td>
            <td><strong>${r.codigo || '<span class="text-muted">Sin código</span>'}</strong></td>
            <td><span class="badge-tag" style="background:var(--bg-subtle); color:var(--text-main); font-size:0.75rem;">${r.localidad}</span></td>
            <td><span class="badge" style="${r.isAntiguo ? 'background:#dbeafe; color:#1e40af; border:1px solid #bfdbfe;' : 'background:#d1fae5; color:#065f46; border:1px solid #a7f3d0;'} font-size:0.72rem; padding:2px 7px; border-radius:10px; font-weight:600;">${r.tipoUsuarioLabel}</span></td>
            <td><strong>${r.fechaStr}</strong></td>
            <td class="text-center">${r.isParaAnalizar ? '<span class="check-badge check-ok"><i data-lucide="check"></i> SÍ (Validado R2)</span>' : '<span class="check-badge check-alert"><i data-lucide="x"></i> NO (Excluido)</span>'}</td>
            <td class="text-center">${r.inYear ? `<span class="check-badge check-ok"><i data-lucide="check-check"></i> SÍ (${state.evalYear})</span>` : `<span class="check-badge check-alert"><i data-lucide="x"></i> NO</span>`}</td>
            <td><small class="text-muted">${r.servicioRaw || '-'}</small></td>
            <td class="text-center">${statusHtml}</td>
          `;
          tbody.appendChild(tr);
        });
      }
    }

    const pageInfo = document.getElementById('s3-page-info');
    const pageNum = document.getElementById('s3-page-num');
    const prevBtn = document.getElementById('s3-page-prev');
    const nextBtn = document.getElementById('s3-page-next');
    const totalPages = Math.ceil(total / s3.pageSize) || 1;

    if (pageInfo) pageInfo.textContent = `Mostrando ${pageRows.length} de ${total.toLocaleString()} registros GIS (Pág. ${s3.page}/${totalPages})`;
    if (pageNum) pageNum.textContent = s3.page;
    if (prevBtn) prevBtn.disabled = s3.page <= 1;
    if (nextBtn) nextBtn.disabled = s3.page >= totalPages;

    if (window.lucide) window.lucide.createIcons();
  }

  // Setup Event Listeners for Stage Controls
  function setupStageEventListeners() {
    // Stage 1 Listeners
    const s1Search = document.getElementById('s1-search-input');
    const s1Loc = document.getElementById('s1-filter-locality');
    const s1UserType = document.getElementById('s1-filter-user-type');
    const s1Stat = document.getElementById('s1-filter-status');
    const s1Prev = document.getElementById('s1-page-prev');
    const s1Next = document.getElementById('s1-page-next');

    if (s1Search) s1Search.addEventListener('input', (e) => {
      state.stages.s1.search = e.target.value.trim();
      state.stages.s1.page = 1;
      filterAndRenderStage1Table();
    });
    if (s1Loc) s1Loc.addEventListener('change', (e) => {
      state.stages.s1.localityFilter = e.target.value;
      state.stages.s1.page = 1;
      filterAndRenderStage1Table();
    });
    if (s1UserType) s1UserType.addEventListener('change', (e) => {
      state.stages.s1.userTypeFilter = e.target.value;
      state.stages.s1.page = 1;
      filterAndRenderStage1Table();
    });
    if (s1Stat) s1Stat.addEventListener('change', (e) => {
      state.stages.s1.statusFilter = e.target.value;
      state.stages.s1.page = 1;
      filterAndRenderStage1Table();
    });
    if (s1Prev) s1Prev.addEventListener('click', () => {
      if (state.stages.s1.page > 1) {
        state.stages.s1.page--;
        filterAndRenderStage1Table();
      }
    });
    if (s1Next) s1Next.addEventListener('click', () => {
      state.stages.s1.page++;
      filterAndRenderStage1Table();
    });

    // Stage 2 Listeners
    const s2Search = document.getElementById('s2-search-input');
    const s2Loc = document.getElementById('s2-filter-locality');
    const s2UserType = document.getElementById('s2-filter-user-type');
    const s2Stat = document.getElementById('s2-filter-status');
    const s2Prev = document.getElementById('s2-page-prev');
    const s2Next = document.getElementById('s2-page-next');

    if (s2Search) s2Search.addEventListener('input', (e) => {
      state.stages.s2.search = e.target.value.trim();
      state.stages.s2.page = 1;
      filterAndRenderStage2Table();
    });
    if (s2Loc) s2Loc.addEventListener('change', (e) => {
      state.stages.s2.localityFilter = e.target.value;
      state.stages.s2.page = 1;
      filterAndRenderStage2Table();
    });
    if (s2UserType) s2UserType.addEventListener('change', (e) => {
      state.stages.s2.userTypeFilter = e.target.value;
      state.stages.s2.page = 1;
      filterAndRenderStage2Table();
    });
    if (s2Stat) s2Stat.addEventListener('change', (e) => {
      state.stages.s2.statusFilter = e.target.value;
      state.stages.s2.page = 1;
      filterAndRenderStage2Table();
    });
    if (s2Prev) s2Prev.addEventListener('click', () => {
      if (state.stages.s2.page > 1) {
        state.stages.s2.page--;
        filterAndRenderStage2Table();
      }
    });
    if (s2Next) s2Next.addEventListener('click', () => {
      state.stages.s2.page++;
      filterAndRenderStage2Table();
    });

    // Stage 3 Listeners
    const s3Search = document.getElementById('s3-search-input');
    const s3Loc = document.getElementById('s3-filter-locality');
    const s3UserType = document.getElementById('s3-filter-user-type');
    const s3Stat = document.getElementById('s3-filter-status');
    const s3Prev = document.getElementById('s3-page-prev');
    const s3Next = document.getElementById('s3-page-next');

    if (s3Search) s3Search.addEventListener('input', (e) => {
      state.stages.s3.search = e.target.value.trim();
      state.stages.s3.page = 1;
      filterAndRenderStage3Table();
    });
    if (s3Loc) s3Loc.addEventListener('change', (e) => {
      state.stages.s3.localityFilter = e.target.value;
      state.stages.s3.page = 1;
      filterAndRenderStage3Table();
    });
    if (s3UserType) s3UserType.addEventListener('change', (e) => {
      state.stages.s3.userTypeFilter = e.target.value;
      state.stages.s3.page = 1;
      filterAndRenderStage3Table();
    });
    if (s3Stat) s3Stat.addEventListener('change', (e) => {
      state.stages.s3.statusFilter = e.target.value;
      state.stages.s3.page = 1;
      filterAndRenderStage3Table();
    });
    if (s3Prev) s3Prev.addEventListener('click', () => {
      if (state.stages.s3.page > 1) {
        state.stages.s3.page--;
        filterAndRenderStage3Table();
      }
    });
    if (s3Next) s3Next.addEventListener('click', () => {
      state.stages.s3.page++;
      filterAndRenderStage3Table();
    });
  }

  setupStageEventListeners();

  // Export to Excel (.xlsx)
  btnExportConsolidated.addEventListener('click', () => {
    exportToExcel(false);
  });

  btnExportObservations.addEventListener('click', () => {
    exportToExcel(true);
  });

  function exportToExcel(onlyObservations = false) {
    if (!window.XLSX) {
      showToast('Biblioteca XLSX no disponible', 'error');
      return;
    }

    let recordsToExport = state.evaluatedRecords || [];
    if (onlyObservations) {
      recordsToExport = recordsToExport.filter(r => r.status !== 'CUMPLE');
    }

    if (recordsToExport.length === 0) {
      showToast('No hay registros para exportar', 'warning');
      return;
    }

    const activeBtn = onlyObservations ? btnExportObservations : btnExportConsolidated;
    const origHtml = activeBtn ? activeBtn.innerHTML : '';
    if (activeBtn) {
      activeBtn.disabled = true;
      activeBtn.innerHTML = '<i data-lucide="loader" class="animate-spin"></i> <span>Generando Excel...</span>';
      if (window.lucide) window.lucide.createIcons();
    }

    showToast('Generando libro Excel... Por favor espere unos segundos', 'info');

    // Usar setTimeout para permitir que el navegador actualice la interfaz de usuario antes del procesamiento intensivo
    setTimeout(() => {
      try {
        const wb = XLSX.utils.book_new();

        // 1. Sheet: Resumen Ejecutivo (Con las 4 Matrices de Evaluación Regulatoria)
        const s = state.stats || {};
        const summaryData = [
          ['EVALUACIÓN DE META DE GESTIÓN: CATASTRO COMERCIAL (CRITERIO SUNASS)'],
          ['Fecha de Reporte', new Date().toLocaleDateString('es-PE')],
          ['Empresa Prestadora (EPS)', state.epsName || 'EPS'],
          ['Año Regulatorio en Evaluación', state.evalYear || 2025],
          ['Rango de Fechas Aplicado (Fichas / GIS)', `${state.customStart} al ${state.customEnd}`],
          ['Filtro de Fecha Propio en Base 1', state.base1DateFilter && state.base1DateFilter.enabled ? `ACTIVO (${state.base1DateFilter.start} al ${state.base1DateFilter.end})` : 'DESACTIVADO (Todo el Padrón)'],
          [],
          ['INDICADOR / PARÁMETRO', 'VALOR', 'UNIDAD / %'],
          ['Conexiones Totales a Catastrar (Base 1)', s.totalPadron || 0, 'Conexiones Físicas'],
          ['Total Servicios Evaluados en Padrón (Agua y Alcantarillado)', s.totalRecords || 0, 'Servicios Regulatorios'],
          ['Servicios Validados como Catastrados (CUMPLE)', s.validCount || 0, `${s.compliancePct || 0}% de la Base 1`],
          ['Servicios con Observación (Falta GIS/Ficha/Fecha)', s.observedTotal || 0, `${s.totalRecords ? ((s.observedTotal/s.totalRecords)*100).toFixed(1) : 0}%`],
          ['Servicios No Catastrados en el Año', s.noCatastradoCount || 0, `${s.totalRecords ? ((s.noCatastradoCount/s.totalRecords)*100).toFixed(1) : 0}%`],
          [],
          ['EVALUACIÓN REGULATORIA POR TIPO DE SERVICIO (CRITERIO SUNASS)'],
          ['Tipo de Servicio', 'Total Padrón (Base 1)', 'Catastrado Válido (Cumple)', 'Con Observaciones', '% Cobertura Catastral', 'Criterio Regulatorio Aplicado'],
          ['Agua Potable', s.totalAgua || 0, s.validAgua || 0, s.obsAgua || 0, `${s.pctAgua || 0}%`, 'Incluye "Agua potable" y conexiones mixtas "Agua y Alcantarillado"'],
          ['Alcantarillado / Saneamiento', s.totalAlc || 0, s.validAlc || 0, s.obsAlc || 0, `${s.pctAlc || 0}%`, 'Incluye "Alcantarillado/Saneamiento" y conexiones mixtas "Agua y Alcantarillado"'],
          [],
          // MATRIZ 1
          ['MATRIZ 1: CONEXIONES TOTALES A CATASTRAR (PADRÓN BASE 1)'],
          [
            'Localidad',
            'Reportado AP', 'Reportado AL', 'Reportado AP y AL',
            'Conexiones Totales AP', 'Conexiones Totales AL', 'Conexiones Totales TOTAL',
            'Observadas AP', 'Observadas AL', 'Observadas TOTAL',
            'Validadas AP', 'Validadas AL', 'Validadas TOTAL'
          ],
        ];

        // Compute Stage 1 matrix rows for Excel
        const s1LocMap = {};
        (state.stages.s1.data || []).forEach(r => {
          const loc = r.localidad || 'SIN LOCALIDAD';
          if (!s1LocMap[loc]) {
            s1LocMap[loc] = {
              rep_ap: new Set(), rep_al: new Set(), rep_ambos: new Set(),
              ap: new Set(), al: new Set(), tot: new Set(),
              obs_ap: new Set(), obs_al: new Set(), obs: new Set(),
              val_ap: new Set(), val_al: new Set(), val: new Set()
            };
          }
          s1LocMap[loc].tot.add(r.suministro);
          if (r.hasAP) s1LocMap[loc].ap.add(r.suministro);
          if (r.hasAL) s1LocMap[loc].al.add(r.suministro);

          if (r.hasAP && r.hasAL) {
            s1LocMap[loc].rep_ambos.add(r.suministro);
          } else if (r.hasAP) {
            s1LocMap[loc].rep_ap.add(r.suministro);
          } else if (r.hasAL) {
            s1LocMap[loc].rep_al.add(r.suministro);
          }

          if (r.cumple) {
            s1LocMap[loc].val.add(r.suministro);
            if (r.hasAP) s1LocMap[loc].val_ap.add(r.suministro);
            if (r.hasAL) s1LocMap[loc].val_al.add(r.suministro);
          } else {
            s1LocMap[loc].obs.add(r.suministro);
            if (r.hasAP) s1LocMap[loc].obs_ap.add(r.suministro);
            if (r.hasAL) s1LocMap[loc].obs_al.add(r.suministro);
          }
        });

        let s1GrandRepAp = 0, s1GrandRepAl = 0, s1GrandRepAmbos = 0;
        let s1GrandAp = 0, s1GrandAl = 0, s1GrandTot = 0;
        let s1GrandObsAp = 0, s1GrandObsAl = 0, s1GrandObs = 0;
        let s1GrandValAp = 0, s1GrandValAl = 0, s1GrandVal = 0;

        Object.keys(s1LocMap).sort().forEach(loc => {
          const d = s1LocMap[loc];
          s1GrandRepAp += d.rep_ap.size; s1GrandRepAl += d.rep_al.size; s1GrandRepAmbos += d.rep_ambos.size;
          s1GrandAp += d.ap.size; s1GrandAl += d.al.size; s1GrandTot += d.tot.size;
          s1GrandObsAp += d.obs_ap.size; s1GrandObsAl += d.obs_al.size; s1GrandObs += d.obs.size;
          s1GrandValAp += d.val_ap.size; s1GrandValAl += d.val_al.size; s1GrandVal += d.val.size;

          summaryData.push([
            loc,
            d.rep_ap.size, d.rep_al.size, d.rep_ambos.size,
            d.ap.size, d.al.size, d.tot.size,
            d.obs_ap.size, d.obs_al.size, d.obs.size,
            d.val_ap.size, d.val_al.size, d.val.size
          ]);
        });
        summaryData.push([
          'TOTAL GENERAL',
          s1GrandRepAp, s1GrandRepAl, s1GrandRepAmbos,
          s1GrandAp, s1GrandAl, s1GrandTot,
          s1GrandObsAp, s1GrandObsAl, s1GrandObs,
          s1GrandValAp, s1GrandValAl, s1GrandVal
        ]);
        summaryData.push([]);

        // MATRIZ 2
        summaryData.push(['MATRIZ 2: FICHAS CATASTRALES EN CAMPO (BASE 2)']);
        summaryData.push([
          'Localidad',
          'Remitidas AP', 'Remitidas AL', 'Remitidas TOTAL',
          'Para Analizar AP', 'Para Analizar AL', 'Para Analizar TOTAL',
          'Observadas AP', 'Observadas AL', 'Observadas TOTAL',
          'Validadas AP', 'Validadas AL', 'Validadas TOTAL'
        ]);

        const s2LocMap = {};
        (state.stages.s2.data || []).forEach(r => {
          const loc = r.localidad || 'SIN LOCALIDAD';
          if (!s2LocMap[loc]) {
            s2LocMap[loc] = {
              rem_ap: new Set(), rem_al: new Set(), rem: new Set(),
              ana_ap: new Set(), ana_al: new Set(), ana: new Set(),
              tot_ap: new Set(), tot_al: new Set(), tot: new Set(),
              obs_ap: new Set(), obs_al: new Set(), obs: new Set(),
              val_ap: new Set(), val_al: new Set(), val: new Set()
            };
          }
          const m = s2LocMap[loc];
          m.rem.add(r.suministro);
          m.tot.add(r.suministro);
          if (r.hasAP) { m.rem_ap.add(r.suministro); m.tot_ap.add(r.suministro); }
          if (r.hasAL) { m.rem_al.add(r.suministro); m.tot_al.add(r.suministro); }

          if (r.isParaAnalizar) {
            m.ana.add(r.suministro);
            if (r.hasAP) m.ana_ap.add(r.suministro);
            if (r.hasAL) m.ana_al.add(r.suministro);

            if (r.isValid) {
              m.val.add(r.suministro);
              if (r.hasAP) m.val_ap.add(r.suministro);
              if (r.hasAL) m.val_al.add(r.suministro);
            } else {
              m.obs.add(r.suministro);
              if (r.hasAP) m.obs_ap.add(r.suministro);
              if (r.hasAL) m.obs_al.add(r.suministro);
            }
          }
        });

        const s2G = {
          rap: 0, ral: 0, rtot: 0,
          aap: 0, aal: 0, atot: 0,
          oap: 0, oal: 0, otot: 0,
          vap: 0, val: 0, vtot: 0
        };
        Object.keys(s2LocMap).sort().forEach(loc => {
          const d = s2LocMap[loc];
          s2G.rap += d.rem_ap.size; s2G.ral += d.rem_al.size; s2G.rtot += d.rem.size;
          s2G.aap += d.ana_ap.size; s2G.aal += d.ana_al.size; s2G.atot += d.ana.size;
          s2G.oap += d.obs_ap.size; s2G.oal += d.obs_al.size; s2G.otot += d.obs.size;
          s2G.vap += d.val_ap.size; s2G.val += d.val_al.size; s2G.vtot += d.val.size;
          summaryData.push([
            loc,
            d.rem_ap.size, d.rem_al.size, d.rem.size,
            d.ana_ap.size, d.ana_al.size, d.ana.size,
            d.obs_ap.size, d.obs_al.size, d.obs.size,
            d.val_ap.size, d.val_al.size, d.val.size
          ]);
        });
        summaryData.push([
          'TOTAL GENERAL',
          s2G.rap, s2G.ral, s2G.rtot,
          s2G.aap, s2G.aal, s2G.atot,
          s2G.oap, s2G.oal, s2G.otot,
          s2G.vap, s2G.val, s2G.vtot
        ]);
        summaryData.push([]);

        // MATRIZ 3
        summaryData.push(['MATRIZ 3: PLATAFORMA QGIS (BASE 3)']);
        summaryData.push([
          'Localidad',
          'Remitidos AP', 'Remitidos AL', 'Remitidos TOTAL',
          'Para Analizar AP', 'Para Analizar AL', 'Para Analizar TOTAL',
          'Observados AP', 'Observados AL', 'Observados TOTAL',
          'Validados AP', 'Validados AL', 'Validados TOTAL'
        ]);

        const s3LocMap = {};
        (state.stages.s3.data || []).forEach(r => {
          const loc = r.localidad || 'SIN LOCALIDAD';
          if (!s3LocMap[loc]) {
            s3LocMap[loc] = {
              rem_ap: new Set(), rem_al: new Set(), rem: new Set(),
              ana_ap: new Set(), ana_al: new Set(), ana: new Set(),
              tot_ap: new Set(), tot_al: new Set(), tot: new Set(),
              obs_ap: new Set(), obs_al: new Set(), obs: new Set(),
              val_ap: new Set(), val_al: new Set(), val: new Set()
            };
          }
          const m = s3LocMap[loc];
          m.rem.add(r.suministro);
          m.tot.add(r.suministro);
          if (r.hasAP) { m.rem_ap.add(r.suministro); m.tot_ap.add(r.suministro); }
          if (r.hasAL) { m.rem_al.add(r.suministro); m.tot_al.add(r.suministro); }

          if (r.isParaAnalizar) {
            m.ana.add(r.suministro);
            if (r.hasAP) m.ana_ap.add(r.suministro);
            if (r.hasAL) m.ana_al.add(r.suministro);

            if (r.isValid) {
              m.val.add(r.suministro);
              if (r.hasAP) m.val_ap.add(r.suministro);
              if (r.hasAL) m.val_al.add(r.suministro);
            } else {
              m.obs.add(r.suministro);
              if (r.hasAP) m.obs_ap.add(r.suministro);
              if (r.hasAL) m.obs_al.add(r.suministro);
            }
          }
        });

        const s3G = {
          rap: 0, ral: 0, rtot: 0,
          aap: 0, aal: 0, atot: 0,
          oap: 0, oal: 0, otot: 0,
          vap: 0, val: 0, vtot: 0
        };
        Object.keys(s3LocMap).sort().forEach(loc => {
          const d = s3LocMap[loc];
          s3G.rap += d.rem_ap.size; s3G.ral += d.rem_al.size; s3G.rtot += d.rem.size;
          s3G.aap += d.ana_ap.size; s3G.aal += d.ana_al.size; s3G.atot += d.ana.size;
          s3G.oap += d.obs_ap.size; s3G.oal += d.obs_al.size; s3G.otot += d.obs.size;
          s3G.vap += d.val_ap.size; s3G.val += d.val_al.size; s3G.vtot += d.val.size;
          summaryData.push([
            loc,
            d.rem_ap.size, d.rem_al.size, d.rem.size,
            d.ana_ap.size, d.ana_al.size, d.ana.size,
            d.obs_ap.size, d.obs_al.size, d.obs.size,
            d.val_ap.size, d.val_al.size, d.val.size
          ]);
        });
        summaryData.push([
          'TOTAL GENERAL',
          s3G.rap, s3G.ral, s3G.rtot,
          s3G.aap, s3G.aal, s3G.atot,
          s3G.oap, s3G.oal, s3G.otot,
          s3G.vap, s3G.val, s3G.vtot
        ]);
        summaryData.push([]);

        // MATRIZ 4: CRUCE FINAL CONSOLIDADO
        summaryData.push(['MATRIZ 4: CRUCE FINAL CONSOLIDADO (SUPERVISIÓN SUNASS)']);
        summaryData.push([
          'Localidad',
          'Base 1 AP', 'Base 1 AL', 'Base 1 TOTAL',
          'Fichas Totales AP', 'Fichas Totales AL', 'Fichas Totales TOTAL',
          'Fichas Obs AP', 'Fichas Obs AL', 'Fichas Obs TOTAL',
          'Fichas Val AP', 'Fichas Val AL', 'Fichas Val TOTAL',
          'QGIS Totales AP', 'QGIS Totales AL', 'QGIS Totales TOTAL',
          'QGIS Obs AP', 'QGIS Obs AL', 'QGIS Obs TOTAL',
          'QGIS Val AP', 'QGIS Val AL', 'QGIS Val TOTAL',
          'CUMPLEN META', '% AVANCE'
        ]);

        // Preindexar suministros que cumplen por localidad para evitar bucles O(N*M)
        const cumpleByLoc = new Map();
        (state.evaluatedRecords || []).forEach(r => {
          if (r.status === 'CUMPLE') {
            const loc = r.localidad || 'SIN LOCALIDAD';
            if (!cumpleByLoc.has(loc)) cumpleByLoc.set(loc, new Set());
            cumpleByLoc.get(loc).add(r.suministro);
          }
        });

        const masterLocs = new Set([...Object.keys(s1LocMap), ...Object.keys(s2LocMap), ...Object.keys(s3LocMap)]);
        const mg = {
          b1_ap: 0, b1_al: 0, b1_tot: 0,
          b2_tap: 0, b2_tal: 0, b2_tot: 0,
          b2_oap: 0, b2_oal: 0, b2_otot: 0,
          b2_vap: 0, b2_val: 0, b2_vtot: 0,
          b3_tap: 0, b3_tal: 0, b3_tot: 0,
          b3_oap: 0, b3_oal: 0, b3_otot: 0,
          b3_vap: 0, b3_val: 0, b3_vtot: 0,
          cumple: 0
        };

        const createEmptyLocSet = () => ({
          ap: new Set(), al: new Set(), tot: new Set(),
          tot_ap: new Set(), tot_al: new Set(), rem_ap: new Set(), rem_al: new Set(), rem: new Set(),
          obs_ap: new Set(), obs_al: new Set(), obs: new Set(),
          val_ap: new Set(), val_al: new Set(), val: new Set()
        });

        Array.from(masterLocs).sort().forEach(loc => {
          const d1 = s1LocMap[loc] || createEmptyLocSet();
          const d2 = s2LocMap[loc] || createEmptyLocSet();
          const d3 = s3LocMap[loc] || createEmptyLocSet();

          const d1_ap = (d1.ap ? d1.ap.size : 0);
          const d1_al = (d1.al ? d1.al.size : 0);
          const d1_tot = (d1.tot ? d1.tot.size : 0);

          const d2_tap = (d2.tot_ap ? d2.tot_ap.size : (d2.rem_ap ? d2.rem_ap.size : 0));
          const d2_tal = (d2.tot_al ? d2.tot_al.size : (d2.rem_al ? d2.rem_al.size : 0));
          const d2_tot = (d2.tot ? d2.tot.size : (d2.rem ? d2.rem.size : 0));
          const d2_oap = (d2.obs_ap ? d2.obs_ap.size : 0);
          const d2_oal = (d2.obs_al ? d2.obs_al.size : 0);
          const d2_otot = (d2.obs ? d2.obs.size : 0);
          const d2_vap = (d2.val_ap ? d2.val_ap.size : 0);
          const d2_val = (d2.val_al ? d2.val_al.size : 0);
          const d2_vtot = (d2.val ? d2.val.size : 0);

          const d3_tap = (d3.tot_ap ? d3.tot_ap.size : (d3.rem_ap ? d3.rem_ap.size : 0));
          const d3_tal = (d3.tot_al ? d3.tot_al.size : (d3.rem_al ? d3.rem_al.size : 0));
          const d3_tot = (d3.tot ? d3.tot.size : (d3.rem ? d3.rem.size : 0));
          const d3_oap = (d3.obs_ap ? d3.obs_ap.size : 0);
          const d3_oal = (d3.obs_al ? d3.obs_al.size : 0);
          const d3_otot = (d3.obs ? d3.obs.size : 0);
          const d3_vap = (d3.val_ap ? d3.val_ap.size : 0);
          const d3_val = (d3.val_al ? d3.val_al.size : 0);
          const d3_vtot = (d3.val ? d3.val.size : 0);

          const cumpleCount = cumpleByLoc.has(loc) ? cumpleByLoc.get(loc).size : 0;
          const pct = d1_tot > 0 ? ((cumpleCount / d1_tot) * 100).toFixed(1) + '%' : '0%';

          mg.b1_ap += d1_ap; mg.b1_al += d1_al; mg.b1_tot += d1_tot;
          mg.b2_tap += d2_tap; mg.b2_tal += d2_tal; mg.b2_tot += d2_tot;
          mg.b2_oap += d2_oap; mg.b2_oal += d2_oal; mg.b2_otot += d2_otot;
          mg.b2_vap += d2_vap; mg.b2_val += d2_val; mg.b2_vtot += d2_vtot;
          mg.b3_tap += d3_tap; mg.b3_tal += d3_tal; mg.b3_tot += d3_tot;
          mg.b3_oap += d3_oap; mg.b3_oal += d3_oal; mg.b3_otot += d3_otot;
          mg.b3_vap += d3_vap; mg.b3_val += d3_val; mg.b3_vtot += d3_vtot;
          mg.cumple += cumpleCount;

          summaryData.push([
            loc,
            d1_ap, d1_al, d1_tot,
            d2_tap, d2_tal, d2_tot,
            d2_oap, d2_oal, d2_otot,
            d2_vap, d2_val, d2_vtot,
            d3_tap, d3_tal, d3_tot,
            d3_oap, d3_oal, d3_otot,
            d3_vap, d3_val, d3_vtot,
            cumpleCount, pct
          ]);
        });

        const mGrandPct = mg.b1_tot > 0 ? ((mg.cumple / mg.b1_tot) * 100).toFixed(1) + '%' : '0%';
        summaryData.push([
          'TOTAL GENERAL',
          mg.b1_ap, mg.b1_al, mg.b1_tot,
          mg.b2_tap, mg.b2_tal, mg.b2_tot,
          mg.b2_oap, mg.b2_oal, mg.b2_otot,
          mg.b2_vap, mg.b2_val, mg.b2_vtot,
          mg.b3_tap, mg.b3_tal, mg.b3_tot,
          mg.b3_oap, mg.b3_oal, mg.b3_otot,
          mg.b3_vap, mg.b3_val, mg.b3_vtot,
          mg.cumple, mGrandPct
        ]);

        const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
        XLSX.utils.book_append_sheet(wb, wsSummary, 'Resumen Ejecutivo');

        // 2. Sheet: Detalle Consolidado con Columnas Requeridas usando aoa_to_sheet (altamente eficiente para >80k filas)
        const detailHeaders = [
          'N°',
          'Número de Suministro',
          'Usuario / Titular',
          'Dirección / Predio',
          'Localidad / Sector',
          'Tipo de Usuario',
          'Fecha Instalación (Base 1)',
          'Tipo de servicio (Real)',
          'Tipo de servicio (Catastro comercial)',
          'En Padrón Conexiones Totales',
          'En Fichas Catastrales',
          'Fecha Ficha Catastral',
          '¿Ficha en Año Evaluado?',
          'En Tabla Atributiva QGIS',
          'Fecha Registro QGIS',
          '¿QGIS en Año Evaluado?',
          'Código Catastral QGIS',
          'ESTADO DE VALIDACIÓN',
          'CUMPLE META',
          'Observación / Diagnóstico Regulatorio'
        ];

        const detailRows = [detailHeaders];
        const nRec = recordsToExport.length;
        for (let i = 0; i < nRec; i++) {
          const r = recordsToExport[i];
          detailRows.push([
            i + 1,
            r.suministro || '',
            r.usuario || '',
            r.direccion || '',
            r.localidad || '',
            r.tipoUsuarioLabel || (r.tipoUsuario === 'NUEVO' ? 'Usuario Nuevo' : 'Usuario Antiguo'),
            r.fechaInstalacion || r.b1Date || '-',
            r.servicioReal || '',
            r.servicioCatastro || '',
            r.inBase1 ? 'SÍ' : 'NO',
            r.inBase2 ? 'SÍ' : 'NO',
            r.fichaDate || '-',
            r.fichaInYear ? 'SÍ' : 'NO',
            r.inBase3 ? 'SÍ' : 'NO',
            r.gisDate || '-',
            r.gisInYear ? 'SÍ' : 'NO',
            r.codigoCatastral || '-',
            r.statusLabel || '',
            r.status === 'CUMPLE' ? 'CUMPLE' : 'NO CUMPLE',
            r.observation || ''
          ]);
        }

        const wsDetail = XLSX.utils.aoa_to_sheet(detailRows);
        XLSX.utils.book_append_sheet(wb, wsDetail, onlyObservations ? 'Observaciones EPS' : 'Cruce Consolidado');

        const cleanEps = (state.epsName || 'EPS').replace(/[^a-zA-Z0-9]/g, '_');
        const filename = onlyObservations 
          ? `OBSERVACIONES_Catastro_Comercial_${cleanEps}_${state.evalYear || 2025}.xlsx`
          : `EVALUACION_CONSOLIDADA_Catastro_Comercial_${cleanEps}_${state.evalYear || 2025}.xlsx`;

        // Descarga robusta usando Blob con fallback a XLSX.writeFile
        triggerWorkbookDownload(wb, filename);
        showToast(`Descarga completada: ${filename}`, 'success');
      } catch (err) {
        console.error('Error generando archivo Excel:', err);
        showToast(`Error al generar Excel: ${err.message}`, 'error');
      } finally {
        if (activeBtn) {
          activeBtn.disabled = false;
          activeBtn.innerHTML = origHtml;
          if (window.lucide) window.lucide.createIcons();
        }
      }
    }, 60);
  }

  // Helper de descarga directa con Blob compatible con navegadores modernos y entornos de archivo local
  function triggerWorkbookDownload(wb, filename) {
    try {
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      if (window.navigator && window.navigator.msSaveOrOpenBlob) {
        window.navigator.msSaveOrOpenBlob(blob, filename);
        return;
      }
      const blobUrl = URL.createObjectURL(blob);
      const downloadLink = document.createElement('a');
      downloadLink.style.display = 'none';
      downloadLink.href = blobUrl;
      downloadLink.download = filename;
      document.body.appendChild(downloadLink);
      downloadLink.click();
      setTimeout(() => {
        if (downloadLink.parentNode) downloadLink.parentNode.removeChild(downloadLink);
        URL.revokeObjectURL(blobUrl);
      }, 2000);
    } catch (e) {
      console.warn('Fallback a XLSX.writeFile:', e);
      XLSX.writeFile(wb, filename);
    }
  }

  // Load Demo Data
  btnLoadDemo.addEventListener('click', () => {
    generateAndLoadDemoData();
  });

  function generateAndLoadDemoData() {
    showToast('Generando bases de datos realistas de prueba con localidades oficiales...', 'info');

    const evalYear = state.evalYear;
    const localities = ['ASIA', 'CERRO AZUL', 'CHILCA', 'IMPERIAL', 'LUNAHUANA', 'MALA', 'NUEVO IMPERIAL'];

    const names = [
      'RAMOS FRANCIA DEMETRIO JUSTO', 'FLORES CAMPOS ROSMERY', 'ARIAS ABURTO JUSTA AYDEE',
      'MENDOZA QUISPE CARLOS ALBERTO', 'HUAMAN CONDORI MARIA ELENA', 'RODRIGUEZ CHAVEZ JUAN JOSE',
      'TORRES VEGA ANA PATRICIA', 'CASTILLO PAREDES JORGE LUIS', 'VARGAS LOPEZ ROSA LUZ',
      'SANCHEZ GUTIERREZ MIGUEL ANGEL', 'QUISPE MAMANI PEDRO', 'ALVAREZ DIAZ CARMEN ROSA',
      'ESPINOZA ROJAS CESAR AUGUSTO', 'HERRERA SALAZAR GLORIA ISABEL', 'CASTRO MORALES VICTOR RAUL',
      'ROMERO NAVARRO DANIEL OMAR', 'GUTIERREZ SILVA MIRIAN LILIANA', 'MEDINA SOTO EDGAR MANUEL',
      'CAMPOS RUIZ DIANA CAROLINA', 'VILLANUEVA ORTIZ HUGO ENRIQUE'
    ];

    const streets = [
      'CAL. RAMON CASTILLA MZ. A LT. 10A', 'AV. PEDRO LA ROSA MZ. A1 LT. 16', 'CAL. LIMA MZ. G LT. 06',
      'JR. BOLOGNESI N° 450', 'AV. LOS PROCERES MZ. C LT. 02', 'CAL. LAS PALMERAS N° 124',
      'AV. SAN MARTIN N° 780', 'JR. GRAU MZ. H LT. 15', 'CAL. TUPAC AMARU N° 320',
      'AV. LIBERTAD MZ. E LT. 08', 'CAL. SUCRE N° 210', 'AV. HUASCAR N° 550'
    ];

    const base1Rows = [];
    const base2Rows = [];
    const base3Rows = [];

    const totalBase1 = 140; // 20 per locality
    const startSupply = 28000000;

    for (let i = 0; i < totalBase1; i++) {
      const sum = (startSupply + i).toString();
      const user = names[i % names.length];
      const street = streets[i % streets.length];
      const locality = localities[i % localities.length];
      const medidor = i % 3 === 0 ? `FA21001${750 + i}` : (i % 3 === 1 ? `QA25423${465 + i}` : '');
      const service = i % 7 === 0 ? 'AGUA POTABLE' : (i % 11 === 0 ? 'ALCANTARILLADO' : 'AGUA Y SANEAMIENTO');

      // Generar 30% como Usuarios Nuevos (instalados durante evalYear) y 70% como Usuarios Antiguos (instalados antes de evalYear)
      const isDemoNuevo = (i % 3 === 0);
      const installYear = isDemoNuevo ? evalYear : (evalYear - 2 - (i % 4));
      const installMonth = String((i % 12) + 1).padStart(2, '0');
      const installDay = String((i % 28) + 1).padStart(2, '0');
      const installDate = `${installDay}/${installMonth}/${installYear}`;

      base1Rows.push({
        'N°': i + 1,
        'Numero de suministro': sum,
        'Usuario': user,
        'Direccion': street,
        'Localidad': locality,
        'N° de medidor': medidor,
        'Tipo de servicio': service,
        'Fecha de instalacion': installDate,
        'Condicion de la conexion': isDemoNuevo ? 'NUEVO' : 'EXISTENTE',
        'Estado de la facturacion': 'ACTIVA'
      });

      // Simulation distribution:
      // i % 10 < 6 (60%): CUMPLE META (Valid Ficha in evalYear + Valid QGIS in evalYear)
      // i % 10 === 6: SOLO FICHA (Valid Ficha in evalYear, missing in QGIS)
      // i % 10 === 7: SOLO GIS (Valid QGIS in evalYear, missing in Fichas)
      // i % 10 === 8: FUERA DE AÑO (Ficha and QGIS in old year evalYear - 2)
      // i % 10 === 9: NO CATASTRADO (Only in Base 1)

      const month = String((i % 12) + 1).padStart(2, '0');
      const day = String((i % 28) + 1).padStart(2, '0');
      const validDate = `${day}/${month}/${evalYear}`;
      const oldDate = `${day}/${month}/${evalYear - 2}`;
      const codCatastral = `2-1-${340 + (i % 50)}-${100 + i}-1`;

      const mod = i % 10;
      if (mod < 6) {
        // CUMPLE
        base2Rows.push({
          'N°': base2Rows.length + 1,
          'Numero de suministro': sum,
          'Usuario': user,
          'Localidad': locality,
          'Modalidad de catastro comercial': 'ASIGNADO',
          'Estado de la actualizacion del catastro comercial': 'ACTUALIZADO',
          'Fecha de ficha catastral': validDate,
          'Tipo de servicio': service,
          'Condicion catastral': 'ACTUALIZADO'
        });

        base3Rows.push({
          'N°': base3Rows.length + 1,
          'Numero de suministro': sum,
          'Usuario': user,
          'Localidad': locality,
          'Codigo catastral Nuevo': codCatastral,
          'Fecha de registro en el sistema grafico': validDate,
          'Tipo de servicio': service,
          'Estado': 'GEORREFERENCIADO'
        });
      } else if (mod === 6) {
        // SOLO FICHA
        base2Rows.push({
          'N°': base2Rows.length + 1,
          'Numero de suministro': sum,
          'Usuario': user,
          'Localidad': locality,
          'Modalidad de catastro comercial': 'ASIGNADO',
          'Estado de la actualizacion del catastro comercial': 'ACTUALIZADO',
          'Fecha de ficha catastral': validDate,
          'Tipo de servicio': service,
          'Condicion catastral': 'ACTUALIZADO'
        });
      } else if (mod === 7) {
        // SOLO GIS (In QGIS, but not in Fichas -> will trigger OBS_FICHA in Stage 3)
        base3Rows.push({
          'N°': base3Rows.length + 1,
          'Numero de suministro': sum,
          'Usuario': user,
          'Localidad': locality,
          'Codigo catastral Nuevo': codCatastral,
          'Fecha de registro en el sistema grafico': validDate,
          'Tipo de servicio': service,
          'Estado': 'GEORREFERENCIADO'
        });
      } else if (mod === 8) {
        // FUERA DE AÑO (Both present but old dates -> triggers OBS_ANO in Stage 2 and Stage 3)
        base2Rows.push({
          'N°': base2Rows.length + 1,
          'Numero de suministro': sum,
          'Usuario': user,
          'Localidad': locality,
          'Modalidad de catastro comercial': 'ASIGNADO',
          'Estado de la actualizacion del catastro comercial': 'ACTUALIZADO',
          'Fecha de ficha catastral': oldDate,
          'Tipo de servicio': service,
          'Condicion catastral': 'ACTUALIZADO'
        });

        base3Rows.push({
          'N°': base3Rows.length + 1,
          'Numero de suministro': sum,
          'Usuario': user,
          'Localidad': locality,
          'Codigo catastral Nuevo': codCatastral,
          'Fecha de registro en el sistema grafico': oldDate,
          'Tipo de servicio': service,
          'Estado': 'GEORREFERENCIADO'
        });
      }
    }

    // Add 4 extra supplies to Base 2 that do NOT exist in Base 1 (to demonstrate OBS_B1: "No figura en Base 1")
    for (let k = 1; k <= 4; k++) {
      const extraSum = (startSupply + 9000 + k).toString();
      base2Rows.push({
        'N°': base2Rows.length + 1,
        'Numero de suministro': extraSum,
        'Usuario': `USUARIO NO REGISTRADO ${k}`,
        'Localidad': localities[k % localities.length],
        'Modalidad de catastro comercial': 'DIRECTO',
        'Estado de la actualizacion del catastro comercial': 'ACTUALIZADO',
        'Fecha de ficha catastral': `15/06/${evalYear}`,
        'Tipo de servicio': 'AGUA Y SANEAMIENTO',
        'Condicion catastral': 'NUEVO'
      });
    }

    // Split demo into 2 locality files per base to showcase multi-file support
    const b1_f1 = base1Rows.slice(0, 70);
    const b1_f2 = base1Rows.slice(70);

    const b2_f1 = base2Rows.slice(0, Math.ceil(base2Rows.length / 2));
    const b2_f2 = base2Rows.slice(Math.ceil(base2Rows.length / 2));

    const b3_f1 = base3Rows.slice(0, Math.ceil(base3Rows.length / 2));
    const b3_f2 = base3Rows.slice(Math.ceil(base3Rows.length / 2));

    state.bases[1].files = [
      { id: 'f_b1_1', name: '01_SECTOR_NORTE_Conexiones.xlsx', rows: b1_f1, rowCount: b1_f1.length, columns: Object.keys(base1Rows[0]) },
      { id: 'f_b1_2', name: '02_SECTOR_SUR_Conexiones.xlsx', rows: b1_f2, rowCount: b1_f2.length, columns: Object.keys(base1Rows[0]) }
    ];

    state.bases[2].files = [
      { id: 'f_b2_1', name: '01_SECTOR_NORTE_Fichas_Catastrales.xlsx', rows: b2_f1, rowCount: b2_f1.length, columns: Object.keys(base2Rows[0]) },
      { id: 'f_b2_2', name: '02_SECTOR_SUR_Fichas_Catastrales.xlsx', rows: b2_f2, rowCount: b2_f2.length, columns: Object.keys(base2Rows[0]) }
    ];

    state.bases[3].files = [
      { id: 'f_b3_1', name: '01_SECTOR_NORTE_Tabla_QGIS.xlsx', rows: b3_f1, rowCount: b3_f1.length, columns: Object.keys(base3Rows[0]) },
      { id: 'f_b3_2', name: '02_SECTOR_SUR_Tabla_QGIS.xlsx', rows: b3_f2, rowCount: b3_f2.length, columns: Object.keys(base3Rows[0]) }
    ];

    // Rebuild data and update visual summaries for all 3 bases
    [1, 2, 3].forEach(id => {
      rebuildBaseData(id);
      updateFilesSummaryUI(id);
      autoDetectMapping(id, state.bases[id].columns);
    });

    updateMappingView();
    checkProcessReady();

    showToast('¡Bases multi-archivo de ejemplo cargadas! Procesando cruce...', 'success');
    executeCrossEvaluation();
  }

  // Reset All
  btnResetAll.addEventListener('click', () => {
    if (confirm('¿Deseas reiniciar la evaluación y limpiar las bases cargadas?')) {
      location.reload();
    }
  });

  // Toast Notification helper
  function showToast(message, type = 'info') {
    const toast = document.getElementById('toast-notification');
    const toastMsg = document.getElementById('toast-message');
    const toastIcon = document.getElementById('toast-icon');

    toast.className = `toast-notification toast-${type}`;
    toastMsg.textContent = message;

    if (type === 'success') {
      toastIcon.innerHTML = '<i data-lucide="check-circle-2"></i>';
    } else if (type === 'error') {
      toastIcon.innerHTML = '<i data-lucide="alert-octagon"></i>';
    } else {
      toastIcon.innerHTML = '<i data-lucide="info"></i>';
    }

    if (window.lucide) window.lucide.createIcons();

    toast.classList.remove('hidden');
    setTimeout(() => {
      toast.classList.add('hidden');
    }, 4500);
  }
});
