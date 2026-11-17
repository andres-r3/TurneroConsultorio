/***** CONFIG *****/
const SHEET_ID = '1LPnBVjCFvm9z1pT0aXC0xBXo09g1taOwxopoEL0yEv8';
const SHEET_NAME = 'Legajos';
const MAX_COL_AI = 15; // A:O

/***** DESTINO ENTREGAS *****/
const ENTREGAS_SHEET_ID = '1CCSbfARDI7gMWJz6rDpxZBe0f2Epn1Y6YbAeBxM-xFI';
const ENTREGAS_TAB_NAME = 'RENDICION ENTREGAS';
const CONSULTA_MONTO = 25000; // monto fijo

/***** HELPERS *****/
function openSheet_() {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) throw new Error(`No existe la hoja "${SHEET_NAME}".`);
    return { ss, sh };
}

/***** WEB APP *****/
function doGet(e) {
    return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('Legajos – ListView')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/***** UTIL *****/
function getConfigInfo() {
    const { ss, sh } = openSheet_();
    return { sheetId: SHEET_ID, fileName: ss.getName(), sheetName: sh.getName() };
}

/***** PRUEBA DE CANAL *****/
function echo10() { return { headers: ['Col1', 'Col2'], rows: [[1, 2], [3, 4]] }; }

/***** DATOS (A:I) *****/
function srv_getTopLegajos10() {
    const { sh } = openSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = Math.min(sh.getLastColumn(), MAX_COL_AI);
    if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [] };
    const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    const n = Math.min(10, lastRow - 1);
    const rows = n > 0 ? sh.getRange(2, 1, n, lastCol).getDisplayValues() : [];
    return { headers, rows };
}

function pasteToActiveCell(key, value) {
    const rng = SpreadsheetApp.getActiveRange();
    if (!rng) throw new Error('No hay celda activa.');
    rng.setValue(value);
    return true;
}

function srv_getLegajosInfo9() {
    const { ss, sh } = openSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = Math.min(sh.getLastColumn(), MAX_COL_AI);
    const headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0] : [];
    return {
        fileName: ss.getName(),
        sheetId: SHEET_ID,
        sheetName: sh.getName(),
        headers,
        totalRows: Math.max(0, lastRow - 1),
        totalCols: lastCol
    };
}

function srv_getStartupData() {
    const { ss, sh } = openSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = Math.min(sh.getLastColumn(), MAX_COL_AI);
    const headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0] : [];
    
    const total = Math.max(0, lastRow - 1);
    const limit = 2000;
    const n = Math.min(limit, total);
    const rows = n > 0 ? sh.getRange(2, 1, n, lastCol).getDisplayValues() : [];
    const nextStart = (n < total) ? n : null;
    const remaining = (n < total) ? (total - n) : 0;
    
    const allReservas = srv_getAllReservas();
    
    return {
        headers,
        totalRows: total,
        rows,
        nextStart,
        remaining,
        allReservas
    };
}

function srv_getLegajosPage9(start, limit) {
    const { sh } = openSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = Math.min(sh.getLastColumn(), MAX_COL_AI);
    const total = Math.max(0, lastRow - 1);
    const s = Math.max(0, Math.min(start || 0, total));
    const n = Math.max(0, Math.min(limit || 200, total - s));
    if (n === 0) return { rows: [], start: s, nextStart: null, remaining: 0 };
    const rows = sh.getRange(2 + s, 1, n, lastCol).getDisplayValues();
    const nextStart = (s + n < total) ? s + n : null;
    const remaining = (s + n < total) ? (total - (s + n)) : 0;
    return { rows, start: s, nextStart, remaining };
}

/***** REPETICIONES *****/
function srv_appendToRepeticiones(headers, row) {
    const { ss } = openSheet_();
    let sh = ss.getSheetByName('Repeticiones');
    if (!sh) sh = ss.insertSheet('Repeticiones');

    const REP_COLS = ['DNI', 'Apellido', 'Medicacion 1', 'Medicacion 2', 'Medicacion 3', 'Medicacion 4']; // [] = todo

    if (Array.isArray(REP_COLS) && REP_COLS.length > 0 && Array.isArray(headers)) {
        const idxs = REP_COLS.map(name => headers.indexOf(name));
        const notFound = idxs.some(i => i < 0);
        if (notFound) throw new Error('Alguna columna de REP_COLS no existe: ' + JSON.stringify(REP_COLS));

        if (sh.getLastRow() === 0) sh.appendRow([...REP_COLS, 'Fecha']);
        const values = [...idxs.map(i => row[i]), new Date()];
        sh.appendRow(values);
        return { ok: true, mode: 'subset', rowNumber: sh.getLastRow() };
    } else {
        if (sh.getLastRow() === 0) sh.appendRow(headers || []);
        sh.appendRow(row || []);
        return { ok: true, mode: 'full', rowNumber: sh.getLastRow() };
    }
}

/***** ENTREGAS *****/
function getPrecioDesdeNombreYCodigo(medicamentoRaw, codigoRaw) {  //NUEVO
    const norm = s => String(s || "").toUpperCase().trim().replace(/\s+/g, " ");

    const med = norm(medicamentoRaw);
    const cod = norm(codigoRaw);

    // ---- Detectores ----
    const esX60 = /\bX\s*60\s*COMP\b/.test(med) || /\bX\s*60\s*COMP\b/.test(cod);

    // ¿Código común/residual? (603, 603/..., 605, 205, 204, 230, 19, 19B, etc.)
    const esComunCode = (s) => {
        if (!s) return false;
        const x = norm(s);
        // prefijos listados + fin o "/" o espacio
        if (/^(603|332|605|205|204|230|19A|19B?)($|[\/\s])/.test(x)) return true;
        // patrón genérico: "NNN/...."
        if (/^\d{3}\/[A-Z0-9/]+$/.test(x)) return true;
        return false;
    };

    const codEsComun = esComunCode(cod);
    const medEsSerieM = /^M[0-4](?:\b|$)/.test(med);
    const medEsFentermina = med.includes("FENTERMINA 30MG");
    const medEsFTD = med.startsWith("FTD");
    const medEsFT = (med === "FT" || med.startsWith("FT "));
    const medEsFastin = med.includes("FASTIN");
    const medEsComun = esComunCode(med);

    // Acompañada: cuando el CÓDIGO es una medicación común/residual
    const acompanada = !!codEsComun;

    // ---- Tabla de precios por tipo ----
    const precioEspecial = (solo30, solo60, acomp30, acomp60) => {
        if (acompanada) return esX60 ? acomp60 : acomp30;
        return esX60 ? solo60 : solo30;
    };


    // MECANISMO UTILIZADO: SOLA, SOLA X 60, ACOMPAÑADA, ACOMPAÑADA X 60
    // 1) FENTERMINA 30MG y FTD (mismo esquema de precios)
  if (medEsFentermina || medEsFTD) {
    return precioEspecial(26000, 52000, 38000, 76000);
  }

  // 2) FASTIN
  if (medEsFastin) {
    return precioEspecial(18000, 36000, 28000, 56000);
  }

  // 3) FT (cuidar orden: FTD ya tratado arriba)
  if (medEsFT) {
    return precioEspecial(20000, 40000, 30000, 60000);
  }

  // 4) Serie M (M0..M4)
  if (medEsSerieM) {
    return precioEspecial(16000, 32000, 28000, 56000);
  }

  // 5) COMÚN/RESIDUAL (603, 603/…, 205/1, 204, 230, 19, 19B, 605, etc.)
  if (medEsComun) {
    return precioEspecial(16000, 28000, 28000, 56000);
    //45=22000
    //15=9000
  }

  // 6) Default (si aparece algo fuera de catálogo, conservamos tu default histórico)
  return 28000;
}

/*function getPrecioDesdeNombreYCodigo(medicamentoRaw, codigoRaw) { //VIEJO
  const norm = s => String(s || "").toUpperCase().trim().replace(/\s+/g, " ");

  const med = norm(medicamentoRaw);
  const cod = norm(codigoRaw);

  // ---- Detectores ----
  const esX60 = /\bX\s*60\s*COMP\b/.test(med) || /\bX\s*60\s*COMP\b/.test(cod);

  // ¿Código común/residual? (603, 603/..., 605, 205, 204, 230, 19, 19B, etc.)
  const esComunCode = (s) => {
    if (!s) return false;
    const x = norm(s);
    // prefijos listados + fin o "/" o espacio
    if (/^(603|332|605|205|204|230|19A|19B?)($|[\/\s])/.test(x)) return true;
    // patrón genérico: "NNN/...."
    if (/^\d{3}\/[A-Z0-9/]+$/.test(x)) return true;
    return false;
  };

  const codEsComun     = esComunCode(cod);
  const medEsSerieM    = /^M[0-4](?:\b|$)/.test(med);
  const medEsFentermina= med.includes("FENTERMINA 30MG");
  const medEsFTD       = med.startsWith("FTD");
  const medEsFT        = (med === "FT" || med.startsWith("FT "));
  const medEsFastin    = med.includes("FASTIN");
  const medEsComun     = esComunCode(med);

  // Acompañada: cuando el CÓDIGO es una medicación común/residual
  const acompanada = !!codEsComun;

  // ---- Tabla de precios por tipo ----
  const precioEspecial = (solo30, solo60, acomp30, acomp60) => {
    if (acompanada) return esX60 ? acomp60 : acomp30;
    return esX60 ? solo60 : solo30;
  };


    // MECANISMO UTILIZADO: SOLA, SOLA X 60, ACOMPAÑADA, ACOMPAÑADA X 60
  // 1) FENTERMINA 30MG y FTD
  if (medEsFentermina || medEsFTD) {
    return precioEspecial(20000, 40000, 32000, 64000);
  }

  // 2) FASTIN
  if (medEsFastin) {
    return precioEspecial(13000, 26000, 23000, 46000);
  }

  // 3) FT (poner después de FTD)
  if (medEsFT) {
    return precioEspecial(14000, 28000, 25000, 50000);
  }

  // 4) Serie M (M0..M4)
  if (medEsSerieM) {
    return precioEspecial(12000, 24000, 22000, 44000);
  }

  // 5) COMÚN/RESIDUAL (603, 603/…, 205/1, 204, 230, 19, 19B, 605, etc.)
  if (medEsComun) {
    return precioEspecial(12000, 24000, 24000, 48000);
    //45=20000
    //15=6000
  }

  // 6) Default
  return 22000;
}*/


function srv_appendToEntregas(headers, row) {
    const idx = (h) => headers.indexOf(h);
    const pick = (h) => { const i = idx(h); return i >= 0 ? row[i] : ''; };

    const dni = pick('DNI');
    const nombre = pick('Apellido') || pick('Nombre');
    const fr1 = pick('Medicacion 1');
    const fr2 = pick('Medicacion 2');
    const fr3 = pick('Medicacion 3');
    const fr4 = pick('Medicacion 4');

    let precio = 0;
    // BLOQUE PRINCIPAL (Medicacion 1 + 2)
    if (fr1 || fr2) {
        precio += getPrecioDesdeNombreYCodigo(fr1, fr2);
    }

    // BLOQUE ADICIONAL (Medicacion 3 + 4)
    if (fr3 || fr4) {
        precio += getPrecioDesdeNombreYCodigo(fr3, fr4);
    }

    const ss = SpreadsheetApp.openById(ENTREGAS_SHEET_ID);
    let sh = ss.getSheetByName(ENTREGAS_TAB_NAME);
    if (!sh) sh = ss.insertSheet(ENTREGAS_TAB_NAME);

    if (sh.getLastRow() === 0) {
        sh.appendRow(['DNI', 'NOMBRE', 'FRASCO 1', 'FRASCO 2', 'FRASCO 3', 'FRASCO 4', 'CONSULTA', 'PRECIO', 'HORARIO']);
    }

    const values = [dni, nombre, fr1, fr2, fr3, fr4, CONSULTA_MONTO, precio, new Date()];
    sh.appendRow(values);

    return {
        ok: true, rowNumber: sh.getLastRow(),
        dni, nombre, precio, consulta: CONSULTA_MONTO, total: Number(CONSULTA_MONTO) + Number(precio)
    };
}

function srv_appendToEntregasSinConsulta(headers, row) {
    const idx = (h) => headers.indexOf(h);
    const pick = (h) => { const i = idx(h); return i >= 0 ? row[i] : ''; };

    const dni = pick('DNI');
    const nombre = pick('Apellido') || pick('Nombre');
    const fr1 = pick('Medicacion 1');
    const fr2 = pick('Medicacion 2');
    const fr3 = pick('Medicacion 3');
    const fr4 = pick('Medicacion 4');

    let precio = 0;
    // BLOQUE PRINCIPAL (Medicacion 1 + 2)
    if (fr1 || fr2) {
        precio += getPrecioDesdeNombreYCodigo(fr1, fr2);
    }

    // BLOQUE ADICIONAL (Medicacion 3 + 4)
    if (fr3 || fr4) {
        precio += getPrecioDesdeNombreYCodigo(fr3, fr4);
    }

    const consulta = ""; // ← SIN consulta

    const ss = SpreadsheetApp.openById(ENTREGAS_SHEET_ID);
    let sh = ss.getSheetByName(ENTREGAS_TAB_NAME);
    if (!sh) sh = ss.insertSheet(ENTREGAS_TAB_NAME);

    if (sh.getLastRow() === 0) {
        sh.appendRow(['DNI', 'NOMBRE', 'FRASCO 1', 'FRASCO 2', 'FRASCO 3', 'FRASCO 4', 'CONSULTA', 'PRECIO', 'HORARIO']);
    }

    const values = [dni, nombre, fr1, fr2, fr3, fr4, consulta, precio, new Date()];
    sh.appendRow(values);

    return {
        ok: true, rowNumber: sh.getLastRow(),
        dni, nombre, precio, consulta, total: Number(consulta) + Number(precio)
    };
}


function srv_previewPrecio(headers, row, opts) {
    const idx = (h) => headers.indexOf(h);
    const pick = (h) => { const i = idx(h); return i >= 0 ? row[i] : ''; };

    const dni = pick('DNI');
    const nombre = pick('Apellido') || pick('Nombre');
    const fr1 = pick('Medicacion 1');
    const fr2 = pick('Medicacion 2');
    const fr3 = pick('Medicacion 3');
    const fr4 = pick('Medicacion 4');

    let precio = 0;
    // BLOQUE PRINCIPAL (Medicacion 1 + 2)
    if (fr1 || fr2) {
        precio += getPrecioDesdeNombreYCodigo(fr1, fr2);
    }

    // BLOQUE ADICIONAL (Medicacion 3 + 4)
    if (fr3 || fr4) {
        precio += getPrecioDesdeNombreYCodigo(fr3, fr4);
    }

    // ⚠️ tomar la consulta desde el FRONT (opts)
    opts = opts || {};
    const consultaFront = Number(opts.consultaMonto || 0);

    // (opcional) fallback a constante del server si la tenés definida aquí también
    const consultaSrv = (typeof CONSULTA_MONTO !== 'undefined') ? Number(CONSULTA_MONTO) : 0;

    const consulta = consultaFront || consultaSrv; // prioridad al front
    const total = Number(precio) + Number(consulta);

    return {
        ok: true,
        dni, nombre, precio, consulta, total,
        fr1, fr2, fr3, fr4 // por si querés mostrar qué tomó
    };
}




/***** Alta de Paciente (A:I) *****/
function srv_addPaciente(p) {
    const { sh } = openSheet_();
    const acceso = String(p?.acceso || '').toLowerCase().trim();
    const validos = ['no permitido', 'permitido', 'en tratamiento'];
    if (!validos.includes(acceso)) throw new Error('Acceso inválido.');

    const dni = String(p?.dni || '').trim();
    const apellido = String(p?.apellido || '').trim();
    const tel = String(p?.telefono || '').trim();
    if (!dni || !apellido) throw new Error('Faltan datos obligatorios: DNI y Apellido/Nombre.');

    let fecha = new Date();
    if (p?.passwordDate) {
        const parts = String(p.passwordDate).split('-');
        if (parts.length === 3) fecha = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    }

    const row = [acceso, fecha, dni, apellido, tel, String(p?.m1 || ''), String(p?.m2 || ''), String(p?.m3 || ''), String(p?.m4 || '')];
    sh.appendRow(row);
    return { ok: true, rowNumber: sh.getLastRow() };
}

/***** Turnero (Disponibilidad/Reservas) *****/
const DISP_TAB_NAME = 'Disponibilidad';
const RESERVAS_TAB = 'Reservas';

function _getDispSheet_() {
    const { ss } = openSheet_();
    const sh = ss.getSheetByName(DISP_TAB_NAME);
    if (!sh) throw new Error(`No existe la hoja "${DISP_TAB_NAME}".`);
    return sh;
}
function _getReservasSheet_() {
    const { ss } = openSheet_();
    let sh = ss.getSheetByName(RESERVAS_TAB);
    if (!sh) { sh = ss.insertSheet(RESERVAS_TAB); sh.appendRow(['Timestamp', 'Fecha', 'Hora', 'DNI', 'NOMBRE']); }
    return sh;
}
function _normDni(x) { return String(x || '').replace(/\D+/g, '').replace(/^0+/, ''); }
function _sameYMD(a, b) { return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate(); }
function _fmtDMY_(d) { const p = n => String(n).padStart(2, '0'); return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()}`; }
function _normHora(x) {
    if (!x) return '';
    if (x instanceof Date) return Utilities.formatDate(x, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm');
    const s = String(x).trim();
    const m = s.match(/(\d{1,2}):(\d{2})/);
    if (m) return ('0' + m[1]).slice(-2) + ':' + m[2];
    return s;
}

function _findColByDateMs(targetMs) {
    const sh = _getDispSheet_();
    const lastCol = sh.getLastColumn();
    if (lastCol < 2) return null;
    const headerDates = sh.getRange(1, 2, 1, lastCol - 1).getValues()[0];
    const target = new Date(Number(targetMs));
    for (let i = 0; i < headerDates.length; i++) {
        const v = headerDates[i];
        if (v instanceof Date && !isNaN(v.getTime()) && _sameYMD(v, target)) return 2 + i;
    }
    return null;
}

function _findHoraRow(horaText) {
    const sh = _getDispSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return null;
    const horasDisp = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues().map(r => _normHora(r[0]));
    const targetHora = _normHora(horaText);
    const idx = horasDisp.findIndex(x => x === targetHora);
    return (idx >= 0) ? (2 + idx) : null;
}

function srv_getDisponibilidadCompleta() {
    const sh = _getDispSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 2) return { fechas: [], horas: [], grilla: {} };
    
    // 1. Obtener fechas de cabecera (fila 1, cols 2 a lastCol)
    const hdrVals = sh.getRange(1, 2, 1, lastCol - 1).getValues()[0];
    const labels = sh.getRange(1, 2, 1, lastCol - 1).getDisplayValues()[0];
    const fechas = [];
    for (let i = 0; i < hdrVals.length; i++) {
        const v = hdrVals[i];
        if (v instanceof Date && !isNaN(v.getTime())) {
            fechas.push({ col: 2 + i, dateMs: v.getTime(), label: labels[i] || _fmtDMY_(v) });
        }
    }
    
    // 2. Obtener horas (columna A, filas 2 a lastRow)
    const horas = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues().map(r => _normHora(r[0]));
    
    // 3. Obtener matriz de estados (filas 2 a lastRow, cols 2 a lastCol)
    const matrix = sh.getRange(2, 2, lastRow - 1, lastCol - 1).getDisplayValues();
    
    // 4. Mapear grilla por fechaMs y hora
    const grilla = {};
    fechas.forEach((f, cIdx) => {
        const dateMs = f.dateMs;
        grilla[dateMs] = [];
        for (let rIdx = 0; rIdx < horas.length; rIdx++) {
            const h = horas[rIdx];
            if (!h) continue;
            const st = (matrix[rIdx][cIdx] || '').toString().trim().toLowerCase();
            const esLibre = (st === '' || st === 'nuevos' || st === 'cancelado' || st === 'confirmar');
            grilla[dateMs].push({ hora: h, libre: esLibre, row: 2 + rIdx, estado: st });
        }
    });
    
    return { fechas, horas, grilla };
}

function srv_listFechasDisponibilidad() {
    const sh = _getDispSheet_();
    const lastCol = sh.getLastColumn();
    if (lastCol < 2) return [];
    const hdrVals = sh.getRange(1, 2, 1, lastCol - 1).getValues()[0];
    const labels = sh.getRange(1, 2, 1, lastCol - 1).getDisplayValues()[0];
    const out = [];
    for (let i = 0; i < hdrVals.length; i++) {
        const v = hdrVals[i];
        if (v instanceof Date && !isNaN(v.getTime())) {
            out.push({ col: 2 + i, dateMs: v.getTime(), label: labels[i] || _fmtDMY_(v) });
        }
    }
    return out;
}

function srv_listHorasParaFecha(dateMs) {
    const sh = _getDispSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 2) return [];
    const headerDates = sh.getRange(1, 2, 1, lastCol - 1).getValues()[0];
    const target = new Date(Number(dateMs));
    let col = null;
    for (let i = 0; i < headerDates.length; i++) {
        const v = headerDates[i];
        if (v instanceof Date && !isNaN(v.getTime()) && _sameYMD(v, target)) { col = 2 + i; break; }
    }
    if (!col) return [];
    const horas = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues().map(r => r[0]);
    const estados = sh.getRange(2, col, lastRow - 1, 1).getDisplayValues().map(r => r[0]);
    const libres = [];
    for (let r = 0; r < horas.length; r++) {
        const h = (horas[r] || '').toString().trim(); if (!h) continue;
        const st = (estados[r] || '').toString().trim().toLowerCase();
        if (st === '' || st === 'nuevos' || st === 'cancelado' || st === 'confirmar') libres.push({ row: 2 + r, hora: h });
    }
    return libres;
}
function srv_sugerirPrimerLibre(dateMs) {
    const slots = srv_listHorasParaFecha(dateMs);
    return slots.length ? slots[0] : null;
}
function srv_reservarTurno(dateMs, hora, dniRaw, nombreRaw, esNuevoFlag) { //(quitar esnuevoflag para quitar reserva por neuvos)
    const lock = LockService.getScriptLock();
    lock.waitLock(10 * 1000);
    try {
        const target = new Date(Number(dateMs));
        const sh = _getDispSheet_();
        const col = _findColByDateMs(dateMs);
        if (!col) throw new Error('No se encontró la columna para esa fecha.');
        const row = _findHoraRow(hora);
        if (!row) throw new Error('No se encontró esa hora en la columna A.');

        const st = (sh.getRange(row, col).getDisplayValue() || '').toString().trim().toLowerCase();
        if (!(st === '' || st === 'nuevos' || st === 'cancelado' || st === 'confirmar')) throw new Error('Ese turno ya no está disponible.');
        const esNuevo = esNuevoFlag === true;
        sh.getRange(row, col).setValue(esNuevo ? 'Reservado Nuevo' : 'Reservado');


        const resSh = _getReservasSheet_();
        const now = new Date(); const dni = _normDni(dniRaw); const nombre = String(nombreRaw || '').trim();
        
        // Buscar saldo a favor (fila de este DNI con fecha vacía y pago TRUE)
        const lastRowRes = resSh.getLastRow();
        const lastColRes = Math.max(resSh.getLastColumn(), 9);
        let filaSaldo = -1;
        
        if (lastRowRes >= 2) {
            const dataRes = resSh.getRange(2, 1, lastRowRes - 1, lastColRes).getValues();
            for (let i = 0; i < dataRes.length; i++) {
                const rDni = _normDni(dataRes[i][3]);
                const rFecha = dataRes[i][1];
                const rPago = dataRes[i][8]; // Columna I
                const tienePago = (rPago === true || String(rPago).toUpperCase() === 'TRUE');
                
                // Fecha vacía y mismo DNI y con pago
                if (rDni === dni && (!rFecha || rFecha === "") && tienePago) {
                    filaSaldo = i + 2;
                    break;
                }
            }
        }
        
        const nuevaFecha = new Date(target.getFullYear(), target.getMonth(), target.getDate());
        const horaNorm = _normHora(hora);
        if (filaSaldo > 0) {
            // Reutilizar la fila del saldo a favor
            resSh.getRange(filaSaldo, 2).setValue(nuevaFecha);
            resSh.getRange(filaSaldo, 3).setValue(horaNorm);
        } else {
            // Nueva reserva normal
            resSh.appendRow([now, nuevaFecha, horaNorm, dni, nombre]);
        }
        
        const filaReserva = (filaSaldo > 0) ? filaSaldo : resSh.getLastRow();
        const allReservas = srv_getAllReservas();
        return { ok: true, allReservas, col, row, filaReserva, fechaLabel: _fmtDMY_(target), hora, dni, nombre, reasignado: (filaSaldo > 0) };
    } finally { try { lock.releaseLock(); } catch (e) { } }
}




function srv_listarTurnosPaciente(dniRaw) {
    const dni = _normDni(dniRaw);
    if (!dni) return [];
    const sh = _getReservasSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = Math.max(sh.getLastColumn(), 9); // Asegurarnos de llegar hasta la I (9)
    if (lastRow < 2) return [];
    
    const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const result = [];
    
    for (let i = 0; i < data.length; i++) {
        const rowDni = _normDni(data[i][3]);
        if (rowDni === dni) {
            const fechaVal = data[i][1];
            const horaVal = _normHora(data[i][2]);
            const pagoVal = data[i][8]; // Columna I (índice 8)
            
            let fechaMs = null;
            let fechaStr = "";
            if (fechaVal instanceof Date && !isNaN(fechaVal.getTime())) {
                fechaMs = fechaVal.getTime();
                fechaStr = _fmtDMY_(fechaVal);
            }
            
            result.push({
                fila: i + 2,
                fechaMs: fechaMs,
                fechaStr: fechaStr,
                hora: horaVal,
                pago: (pagoVal === true || String(pagoVal).toUpperCase() === 'TRUE'),
                nombre: data[i][4] || ''
            });
        }
    }
    return result;
}

function srv_getAllReservas() {
    const sh = _getReservasSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = Math.max(sh.getLastColumn(), 9);
    if (lastRow < 2) return [];
    
    const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const result = [];
    
    for (let i = 0; i < data.length; i++) {
        const rowDni = _normDni(data[i][3]);
        const fechaVal = data[i][1];
        const horaVal = _normHora(data[i][2]);
        const pagoVal = data[i][8]; // Columna I (índice 8)
        
        let fechaMs = null;
        let fechaStr = "";
        if (fechaVal instanceof Date && !isNaN(fechaVal.getTime())) {
            fechaMs = fechaVal.getTime();
            fechaStr = _fmtDMY_(fechaVal);
        }
        
        result.push({
            fila: i + 2,
            dni: rowDni,
            fechaMs: fechaMs,
            fechaStr: fechaStr,
            hora: horaVal,
            pago: (pagoVal === true || String(pagoVal).toUpperCase() === 'TRUE'),
            nombre: data[i][4] || ''
        });
    }
    return result;
}

function srv_cancelarTurno(filaReserva, fechaMs, hora, dniRaw) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10 * 1000);
    try {
        const resSh = _getReservasSheet_();
        const lastRow = resSh.getLastRow();
        const lastCol = Math.max(resSh.getLastColumn(), 9);
        if (filaReserva < 2 || filaReserva > lastRow) throw new Error("Fila de reserva inválida.");
        
        const rowData = resSh.getRange(filaReserva, 1, 1, lastCol).getValues()[0];
        const rowDni = _normDni(rowData[3]);
        const dni = _normDni(dniRaw);
        
        if (rowDni !== dni) throw new Error("El DNI no coincide con el registro.");
        
        const pagoVal = rowData[8]; // Columna I
        const tienePago = (pagoVal === true || String(pagoVal).toUpperCase() === 'TRUE');
        
        // 1. Liberar disponibilidad
        if (fechaMs && hora) {
            const colDisp = _findColByDateMs(fechaMs);
            const rowDisp = _findHoraRow(hora);
            if (colDisp && rowDisp) {
                _getDispSheet_().getRange(rowDisp, colDisp).clearContent();
            }
        }
        
        // 2. Manejar la fila de reservas
        if (tienePago) {
            // Limpiar fecha y hora
            resSh.getRange(filaReserva, 2).clearContent(); // Fecha
            resSh.getRange(filaReserva, 3).clearContent(); // Hora
            const allReservas = srv_getAllReservas();
            return { ok: true, allReservas, accion: "limpiado", mensaje: "Turno cancelado, pero el pago se mantiene como saldo a favor." };
        } else {
            // Eliminar la fila
            resSh.deleteRow(filaReserva);
            const allReservas = srv_getAllReservas();
            return { ok: true, allReservas, accion: "eliminado", mensaje: "Turno eliminado completamente." };
        }
        
    } finally {
        try { lock.releaseLock(); } catch (e) { }
    }
}

function srv_eliminarSaldo(filaReserva, dniRaw) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10 * 1000);
    try {
        const resSh = _getReservasSheet_();
        const lastRow = resSh.getLastRow();
        if (filaReserva < 2 || filaReserva > lastRow) throw new Error("Fila de reserva inválida.");
        
        const rowDni = _normDni(resSh.getRange(filaReserva, 4).getValue());
        const dni = _normDni(dniRaw);
        if (rowDni !== dni) throw new Error("El DNI no coincide.");
        
        resSh.deleteRow(filaReserva);
        const allReservas = srv_getAllReservas();
        return { ok: true, allReservas, mensaje: "Saldo a favor eliminado." };
    } finally {
        try { lock.releaseLock(); } catch (e) { }
    }
}

function srv_reasignarTurno(filaReserva, viejaFechaMs, viejaHora, nuevaFechaMs, nuevaHora, dniRaw) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15 * 1000);
    try {
        const resSh = _getReservasSheet_();
        const dni = _normDni(dniRaw);
        const rowDni = _normDni(resSh.getRange(filaReserva, 4).getValue());
        if (rowDni !== dni) throw new Error("El DNI no coincide con el registro.");

        const shDisp = _getDispSheet_();

        // 1. Liberar el turno viejo
        if (viejaFechaMs && viejaHora) {
            const colOld = _findColByDateMs(viejaFechaMs);
            const rowOld = _findHoraRow(viejaHora);
            if (colOld && rowOld) {
                shDisp.getRange(rowOld, colOld).clearContent();
            }
        }

        // 2. Reservar el turno nuevo
        const colNew = _findColByDateMs(nuevaFechaMs);
        const rowNew = _findHoraRow(nuevaHora);
        if (!colNew || !rowNew) throw new Error("Fecha o hora nueva no válida.");

        const stNew = (shDisp.getRange(rowNew, colNew).getDisplayValue() || '').toString().trim().toLowerCase();
        if (!(stNew === '' || stNew === 'nuevos' || stNew === 'cancelado' || stNew === 'confirmar')) {
             throw new Error('El nuevo turno ya no está disponible.');
        }

        // Verificamos si antes era un turno de "Nuevo" para mantener la etiqueta si fuera necesario, 
        // pero por defecto en secretaría solemos poner "Reservado".
        shDisp.getRange(rowNew, colNew).setValue('Reservado');

        // 3. Actualizar la fila en Reservas
        const dNew = new Date(Number(nuevaFechaMs));
        const hNew = _normHora(nuevaHora);
        resSh.getRange(filaReserva, 2).setValue(dNew);
        resSh.getRange(filaReserva, 3).setValue(hNew);

        const allReservas = srv_getAllReservas();
        return { 
            ok: true, 
            allReservas,
            mensaje: "Turno reasignado correctamente.",
            nuevaFechaLabel: _fmtDMY_(dNew),
            nuevaHora: hNew
        };

    } finally {
        try { lock.releaseLock(); } catch (e) { }
    }
}

//// desde aqui eliminar //////////////////////////////////////// esto es lista de espera
function srv_appendToListaEsperaBatch(batchRows) {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName('Lista espera');
    if (!sh) {
        sh = ss.insertSheet('Lista espera');
        sh.appendRow(['A', 'B', 'C', 'D', 'E', 'J', 'K', 'Timestamp']);
    }

    if (!Array.isArray(batchRows) || batchRows.length === 0) {
        return { ok: true, appended: 0 };
    }

    const now = new Date();
    const values = batchRows.map(r => [
        r?.[0] ?? '',  // B
        r?.[1] ?? '',  // C
        r?.[2] ?? '',  // D
        r?.[3] ?? '',  // E
        r?.[4] ?? '',  // J
        r?.[5] ?? '',  // B
        r?.[6] ?? '',  // k
        now
    ]);

    sh.getRange(sh.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
    return { ok: true, appended: values.length };
}
//// hasta aqui eliminar ////////////////////////////////////////