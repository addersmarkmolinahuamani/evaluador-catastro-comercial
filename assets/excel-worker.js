/**
 * Web Worker para lectura y procesamiento de archivos Excel (.xlsx, .xls, .xlsb, .csv)
 * Evita bloquear el hilo principal (UI) del navegador y recorta filas/columnas fantasma.
 */
importScripts('vendor/xlsx.full.min.js?v=2');

function clampWorksheetRange(ws) {
  if (!ws || !ws['!ref']) return;
  const range = XLSX.utils.decode_range(ws['!ref']);

  const data = Array.isArray(ws) ? ws : (Array.isArray(ws['!data']) ? ws['!data'] : null);
  if (data) {
    let maxRow = 0;
    let maxCol = 0;
    const rowLimit = Math.min(data.length - 1, range.e.r);
    for (let r = 0; r <= rowLimit; r++) {
      const row = data[r];
      if (!row) continue;
      let rowHasVal = false;
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
          rowHasVal = true;
          if (c > maxCol) maxCol = c;
        }
      }
      if (rowHasVal && r > maxRow) maxRow = r;
    }
    range.e.r = Math.min(range.e.r, maxRow);
    range.e.c = Math.min(range.e.c, Math.max(maxCol, 0));
    ws['!ref'] = XLSX.utils.encode_range(range);
    if (data.length > range.e.r + 1) {
      data.length = range.e.r + 1;
    }
    return;
  }

  let maxRow = 0;
  let maxCol = 0;
  const keys = Object.keys(ws);
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    if (k[0] === '!') continue;
    const cell = ws[k];
    if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
      const dec = XLSX.utils.decode_cell(k);
      if (dec.r > maxRow) maxRow = dec.r;
      if (dec.c > maxCol) maxCol = dec.c;
    }
  }
  range.e.r = Math.min(range.e.r, maxRow);
  range.e.c = Math.min(range.e.c, Math.max(maxCol, 0));
  ws['!ref'] = XLSX.utils.encode_range(range);
}

self.onmessage = function (e) {
  const { id, fileName, buffer } = e.data;
  try {
    const workbook = XLSX.read(new Uint8Array(buffer), {
      type: 'array',
      dense: true,
      sheets: 0,
      cellFormula: false,
      cellHTML: false,
      cellStyles: false,
      cellNF: false,
      cellText: false,
      cellDates: false
    });

    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Recortar filas/columnas vacías (ej. hojas .xlsb/.xlsx con 1,000,000 de filas fantasma)
    clampWorksheetRange(worksheet);

    const rawRows = XLSX.utils.sheet_to_json(worksheet, {
      defval: '',
      raw: true,
      blankrows: false
    });

    if (!rawRows || rawRows.length === 0) {
      self.postMessage({ id, ok: true, rows: [], columns: [] });
      return;
    }

    // Filtrar columnas vacías (__EMPTY*) que no tengan datos reales en las primeras filas
    const allCols = Object.keys(rawRows[0]);
    const validCols = allCols.filter(col => {
      if (!col.startsWith('__EMPTY')) return true;
      // Conservar solo si alguna fila tiene valor en esa columna
      for (let i = 0; i < Math.min(rawRows.length, 200); i++) {
        const v = rawRows[i][col];
        if (v !== '' && v !== null && v !== undefined) return true;
      }
      return false;
    });

    // Limpiar localidad desde nombre de archivo si no existe columna explícita de localidad
    const cleanFileLocality = fileName
      .replace(/\.[^/.]+$/, '')
      .replace(/^\d+[\s_-]*/, '')
      .replace(/[_-]+/g, ' ')
      .trim()
      .toUpperCase();

    const hasLocCol = validCols.some(c => /localidad|distrito|sector|zona|anexo/i.test(c));
    const columns = [...validCols];
    if (!hasLocCol && cleanFileLocality && !columns.includes('Localidad')) {
      columns.push('Localidad');
    }

    // Compactar objetos de fila omitiendo columnas __EMPTY descartadas y filas totalmente en blanco
    const needColPrune = validCols.length < allCols.length;
    const rows = [];
    for (let i = 0; i < rawRows.length; i++) {
      const r = rawRows[i];
      let hasData = false;
      const out = needColPrune ? {} : r;

      for (let c = 0; c < validCols.length; c++) {
        const colName = validCols[c];
        const val = r[colName];
        if (val !== '' && val !== null && val !== undefined) {
          hasData = true;
        }
        if (needColPrune) {
          out[colName] = val;
        }
      }

      if (!hasData) continue;
      if (!hasLocCol && cleanFileLocality && !out['Localidad']) {
        out['Localidad'] = cleanFileLocality;
      }
      rows.push(out);
    }

    self.postMessage({ id, ok: true, rows, columns });
  } catch (err) {
    self.postMessage({ id, ok: false, error: err ? err.message : 'Error desconocido al leer Excel' });
  }
};
