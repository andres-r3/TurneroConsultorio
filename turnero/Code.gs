/**** CONFIGURABLES *****************************************************/

const VERSION = '2.4.7';
let cacheHeaders_ = null;

// Archivo y hojas
const SHEET_ID = '1LPnBVjCFvm9z1pT0aXC0xBXo09g1taOwxopoEL0yEv8'; // "Consultorio"
const SHEETS = {
    LEGAJOS: 'Legajos',
    DISP: 'Disponibilidad',
    RESERVAS: 'Reservas',
    REPETICIONES: 'Repeticiones',
    DIALOGOS: 'Dialogos',
    CONFIG: 'Config' // opcional (clave/valor). Ej: PUBLIC_ON | true
};

// Gates (se sobreescriben desde Disponibilidad)
let LISTA_ESPERA_ON = false; // lista de espera habilitada
let PUBLIC_ON = true; // turnero abierto/cerrado
let MED_ON = true; // solicitar medicación habilitado
let CANCEL_ON = true; // cancelar reservas habilitado

// Info caché del consultorio
let INFO_CLINIC = {
    proximaAtencion1: '',
    proximaAtencion2: '',
    fechaMedicacion: '',
    fechaEntregaTurnos: '',
    datosTransferencia: '',
    verPrecioMedOn: false,
    fechaLimiteMedicacion: '',
    horariosEntregaMedicacion: '',
    precioConsulta: 25000,
    customPrices: null
};

// Límites
let REQUIRE_PERMITIDO = true;             // exige Acceso="permitido" en Legajos (col A) para login
let MAX_RESERVAS_POR_MES = 0;             // 1 reserva por mes por DNI (0 = sin tope)
let MAX_RESERVAS_ACTIVAS_POR_DNI = 0;     // tope reservas futuras por DNI (0 = sin tope)
let RATE_SECONDS = 3;                     // antiflood por DNI (intento cada N s)

// Fechas públicas fijas: B1 y C1 en Disponibilidad (2=B, 3=C)
const PUBLIC_DATE_COLS = [2, 3];

// Nombres de encabezados Legajos (por nombre)
const COLS_LEGAJOS = {
    acceso: 'Acceso',
    password: 'Password',
    dni: 'DNI',
    nombre: 'Apellido',
    tel: 'Telefono',
    domicilio: 'Domicilio',   // opcional: si existe se completa, sino se ignora
    fr1: 'Medicacion 1',
    fr2: 'Medicacion 2',
    fr3: 'Medicacion 3',
    fr4: 'Medicacion 4',
    fr5: 'Medicacion 5'
};

/***********************************************************************/

function doGet() {
    return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('Turnero Público')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);

        if (data.action === 'reservar') {
            const { dateMs, horaText, dni, nombre } = data;
            const res = pub_reservar(dateMs, horaText, dni, nombre);
            return json_(res);
        }

        return json_({ ok: false, error: 'Acción no válida' });

    } catch (err) {
        return json_({ ok: false, error: String(err) });
    }
}

function json_(obj) {
    return ContentService
        .createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
}



/***** Helpers base *****/
function _ss_() {
    try {
        const act = SpreadsheetApp.getActiveSpreadsheet();
        if (act) return act;
    } catch (e) { }
    return SpreadsheetApp.openById(SHEET_ID);
}
function _sh_(name) {
    const sh = _ss_().getSheetByName(name);
    if (!sh) throw new Error(`No existe la hoja "${name}".`);
    return sh;
}
function _normDni_(x) { return String(x || '').replace(/\D+/g, '').replace(/^0+/, ''); }
function _sameYMD_(a, b) {
    return a.getFullYear() === b.getFullYear() &&
        a.getMonth() === b.getMonth() &&
        a.getDate() === b.getDate();
}
function _dYMD_(ms) {
    const d = new Date(ms);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function _fmtDMY_(d) {
    const p = n => String(n).padStart(2, '0');
    return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()}`;
}
function _strOrEmpty_(v) { return (v == null) ? '' : String(v); }
function _tz_() { return Session.getScriptTimeZone() || 'America/Argentina/Cordoba'; }

/* Normaliza una “hora” cualquiera (Date o String) a "HH:mm" */
function _normHoraStr_(x) {
    if (x instanceof Date) {
        return Utilities.formatDate(x, _tz_(), 'HH:mm');
    }
    const s = String(x || '').trim();
    const m = s.match(/(\d{1,2}):(\d{2})/);
    if (m) {
        const hh = ('0' + Number(m[1])).slice(-2);
        const mm = m[2];
        return `${hh}:${mm}`;
    }
    return s;
}

/***** Cargar Config desde hoja (opcional) + gates de Disponibilidad *****/
function readDispBool_(a1) {
    try {
        const sh = _sh_(SHEETS.DISP);
        const v = sh.getRange(a1).getValue();
        return (v === true || v === false) ? v : null;
    } catch (e) {
        return null;
    }
}
function loadConfig_() {
    try {
        const sh = _ss_().getSheetByName(SHEETS.CONFIG);
        if (sh) {
            const last = sh.getLastRow();
            if (last > 0) {
                const rows = sh.getRange(1, 1, last, Math.min(2, sh.getLastColumn())).getDisplayValues();
                for (const [k, v] of rows) {
                    if (!k) continue;
                    const key = String(k).trim().toUpperCase();
                    const val = String(v || '').trim();
                    const toBool = (s) => /^true$/i.test(s);
                    const toNum = (s) => Number(s);
                    const toJSON = (s) => { try { return JSON.parse(s); } catch (_) { return null; } };
                    if (key === 'PUBLIC_ON') PUBLIC_ON = toBool(val);
                    else if (key === 'REQUIRE_PERMITIDO') REQUIRE_PERMITIDO = toBool(val);
                    else if (key === 'MAX_RESERVAS_POR_MES') MAX_RESERVAS_POR_MES = toNum(val);
                    else if (key === 'MAX_RESERVAS_ACTIVAS_POR_DNI') MAX_RESERVAS_ACTIVAS_POR_DNI = toNum(val);
                    else if (key === 'RATE_SECONDS') RATE_SECONDS = toNum(val);
                    else if (key === 'PUBLIC_DATE_COLS') {
                        const arr = toJSON(val);
                        if (Array.isArray(arr) && arr.every(n => typeof n === 'number')) {
                            PUBLIC_DATE_COLS.length = 0;
                            for (const n of arr) PUBLIC_DATE_COLS.push(n);
                        }
                    }
                }
            }
        }
    } catch (e) { /* sin Config: defaults */ }

    // Gates e información leídos de una sola vez desde Disponibilidad!I10:I24 (Prioridad)
    try {
        const shDisp = _sh_(SHEETS.DISP);
        const range = shDisp.getRange('I10:I24');
        const vals = range.getValues();
        const dispVals = range.getDisplayValues();

        const toBool = (v) => (v === true || String(v || '').toLowerCase().trim() === 'true');

        // I10 (index 0) -> LISTA_ESPERA_ON
        if (vals[0][0] !== '') LISTA_ESPERA_ON = toBool(vals[0][0]);

        // I11 (index 1) -> PUBLIC_ON
        if (vals[1][0] !== '') PUBLIC_ON = toBool(vals[1][0]);

        // I12 (index 2) -> CANCEL_ON
        if (vals[2][0] !== '') CANCEL_ON = toBool(vals[2][0]);

        // I13 (index 3) -> MED_ON
        if (vals[3][0] !== '') MED_ON = toBool(vals[3][0]);

        // I15 (index 5) -> MAX_RESERVAS_ACTIVAS_POR_DNI
        if (vals[5][0] !== '') {
            const rawVal = vals[5][0];
            if (typeof rawVal === 'boolean') {
                MAX_RESERVAS_ACTIVAS_POR_DNI = rawVal ? 1 : 0;
            } else {
                MAX_RESERVAS_ACTIVAS_POR_DNI = Number(rawVal) || 0;
            }
        }

        // Cargar info del consultorio
        INFO_CLINIC.proximaAtencion1 = String(dispVals[6][0] || '').trim();
        INFO_CLINIC.proximaAtencion2 = String(dispVals[7][0] || '').trim();
        INFO_CLINIC.fechaMedicacion = String(dispVals[8][0] || '').trim();
        INFO_CLINIC.fechaEntregaTurnos = String(dispVals[9][0] || '').trim();
        INFO_CLINIC.datosTransferencia = String(dispVals[11][0] || '').trim();

        // Nuevas celdas para Medicación (I22, I23, I24)
        if (vals.length >= 13 && vals[12][0] !== '') {
            INFO_CLINIC.verPrecioMedOn = toBool(vals[12][0]); // I22
        }
        if (dispVals.length >= 14) {
            INFO_CLINIC.fechaLimiteMedicacion = String(dispVals[13][0] || '').trim(); // I23
        }
        if (dispVals.length >= 15) {
            INFO_CLINIC.horariosEntregaMedicacion = String(dispVals[14][0] || '').trim(); // I24
        }

        // Cargar o auto-crear hoja de PreciosMedicacion durante la carga de configuración inicial
        try {
            const { precioConsulta, customPrices } = getPreciosFromSheet_();
            INFO_CLINIC.precioConsulta = precioConsulta;
            INFO_CLINIC.customPrices = customPrices;
        } catch (e) {
            console.error("Error al precargar PreciosMedicacion en loadConfig_: " + e);
        }
    } catch (e) {
        console.error("Error al cargar Disponibilidad en bloque I10:I24: " + e);
    }
}

/***** Log en Dialogos *****/
function pub_log(action, meta) {
    try {
        const sh = _ss_().getSheetByName(SHEETS.DIALOGOS) || _ss_().insertSheet(SHEETS.DIALOGOS);
        if (sh.getLastRow() === 0) {
            sh.appendRow(['TS', 'ACCION', 'DNI', 'META_JSON', 'RESULT']);
        }
        
        // 1. Si el registro tiene DNI, buscamos tanto el Nombre como el Acceso
        if (meta && meta.dni) {
            const nd = _normDni_(meta.dni);
            const leg = getLegajoByDni_(nd);
            if (leg) {
                meta.nombre = leg.nombre || ''; // Agrega el nombre (ej: "Juan Pérez")
                meta.acceso = leg.acceso || ''; // Agrega el acceso (ej: "no permitido", "permitido", "nuevo")
            }
        }
        
        // 2. Si el registro tiene fechaMs, le agregamos la fecha legible
        if (meta && meta.fechaMs) {
            meta.fecha_legible = _fmtDMY_(new Date(meta.fechaMs));
        }

        const dni = _strOrEmpty_(meta?.dni);
        const result = _strOrEmpty_(meta?.result);
        const json = JSON.stringify(meta || {});
        sh.appendRow([new Date(), action, dni, json, result]);
        return true;
    } catch (e) { return false; }
}


/***** Rate limit por DNI *****/
function rateLimit_(dni, key, seconds) {
    const cache = CacheService.getScriptCache();
    const k = `rate:${key}:${_normDni_(dni)}`;
    if (cache.get(k)) return false;
    cache.put(k, '1', Math.max(1, Math.min(120, seconds | 0)));
    return true;
}

/***** LEGAJOS: helpers *****/
function getLegajosHeaderIndexes_() {
    if (cacheHeaders_) return cacheHeaders_;
    const sh = _sh_(SHEETS.LEGAJOS);
    const lastCol = sh.getLastColumn();
    if (lastCol < 1) throw new Error('Hoja Legajos vacía.');
    const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => String(h || '').trim());
    const idx = (name) => headers.findIndex(h => h.toLowerCase() === String(name).toLowerCase());
    cacheHeaders_ = { headers, idx, lastCol, sh };
    return cacheHeaders_;
}
function findLegajoRowByDni_(dni) {
    const { sh, idx } = getLegajosHeaderIndexes_();
    const colDni = idx(COLS_LEGAJOS.dni) + 1;
    if (colDni < 1) throw new Error('No se encontró la columna DNI en Legajos.');
    const last = sh.getLastRow();
    if (last < 2) return null;
    const dnis = sh.getRange(2, colDni, last - 1, 1).getDisplayValues().map(r => _normDni_(r[0]));
    const target = _normDni_(dni);
    const i = dnis.findIndex(x => x === target);
    if (i < 0) return null;
    return { row: 2 + i };
}
function getLegajoByDni_(dni) {
    const { headers, idx, sh } = getLegajosHeaderIndexes_();
    const rowInfo = findLegajoRowByDni_(dni);
    if (!rowInfo) return null;
    const row = rowInfo.row;
    const vals = sh.getRange(row, 1, 1, headers.length).getDisplayValues()[0];
    const pick = (hname) => {
        const i = idx(hname);
        return i >= 0 ? vals[i] : '';
    };
    return {
        row,
        acceso: pick(COLS_LEGAJOS.acceso),
        dni: pick(COLS_LEGAJOS.dni),
        nombre: pick(COLS_LEGAJOS.nombre),
        telefono: pick(COLS_LEGAJOS.tel),
        domicilio: pick(COLS_LEGAJOS.domicilio),
        fr1: pick(COLS_LEGAJOS.fr1),
        fr2: pick(COLS_LEGAJOS.fr2),
        fr3: pick(COLS_LEGAJOS.fr3),
        fr4: pick(COLS_LEGAJOS.fr4),
        fr5: pick(COLS_LEGAJOS.fr5)
    };
}

/***** DISPONIBILIDAD: fechas públicas (B1, C1 por defecto) *****/
function getPublicDates_() {
    const sh = _sh_(SHEETS.DISP);
    const lastCol = sh.getLastColumn();
    if (lastCol < 2) return [];

    const range = sh.getRange(1, 1, 1, lastCol);
    const vals = range.getValues()[0];
    const dispVals = range.getDisplayValues()[0];

    const out = [];
    for (const c of PUBLIC_DATE_COLS) {
        if (c < 2 || c > lastCol) continue;
        const v = vals[c - 1];
        const lbl = dispVals[c - 1];
        if (v instanceof Date && !isNaN(v.getTime())) {
            out.push({ col: c, dateMs: v.getTime(), label: lbl || _fmtDMY_(v) });
        }
    }
    return out;
}

//avisos por horarios
function listHorasForCol_(col, isNuevo) {
    const sh = _sh_(SHEETS.DISP);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return [];

    const vals = sh.getRange(2, 1, lastRow - 1, col).getDisplayValues();
    const libres = [];

    for (let i = 0; i < vals.length; i++) {
        const h = vals[i][0];
        if (!h) continue;

        const st = vals[i][col - 1];
        const stClean = String(st || '').toLowerCase().trim();
        if (isNuevo) {
            if (stClean !== 'nuevos') continue;
        } else {
            if (stClean !== '') continue;
        }

        const horaNorm = _normHoraStr_(h);   // "HH:mm"
        const requiereTransferencia = isNuevo ? true : (horaNorm <= '10:00');

        libres.push({
            row: 2 + i,
            hora: horaNorm,
            requiereTransferencia
        });
    }

    return libres;
}

function findColByDateMs_(targetMs) {
    const sh = _sh_(SHEETS.DISP);
    const lastCol = sh.getLastColumn();
    if (lastCol < 2) return null;
    const vals = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const target = _dYMD_(targetMs);
    for (const col of PUBLIC_DATE_COLS) {
        if (col < 2 || col > lastCol) continue;
        const v = vals[col - 1];
        if (v instanceof Date && _sameYMD_(v, target)) return col;
    }
    return null;
}
function findHoraRow_(horaText) {
    const sh = _sh_(SHEETS.DISP);
    const lastRow = sh.getLastRow();
    const horasDisp = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues().map(r => _normHoraStr_(r[0]));
    const idx = horasDisp.findIndex(x => x === _normHoraStr_(horaText));
    return (idx >= 0) ? (2 + idx) : null;
}

/***** RESERVAS helpers *****/
function appendReserva_(fechaMs, horaText, dni, nombre) {
    const sh = _sh_(SHEETS.RESERVAS);
    if (sh.getLastRow() === 0) {
        sh.appendRow(['Marca Horaria', 'Fecha', 'Hora', 'DNI', 'Nombre Paciente', 'PAGO']);
    }
    const d = _dYMD_(fechaMs);
    const hora = _normHoraStr_(horaText);
    const nd = _normDni_(dni);

    // Buscar saldo a favor (fila de este DNI con fecha vacía y pago TRUE)
    const lastRow = sh.getLastRow();
    const lastCol = Math.max(sh.getLastColumn(), 9);
    let filaSaldo = -1;

    if (lastRow >= 2) {
        const dataRes = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
        for (let i = 0; i < dataRes.length; i++) {
            const rDni = _normDni_(dataRes[i][3]);
            const rFecha = dataRes[i][1];
            const rPago = dataRes[i][8]; // Columna I
            const tienePago = (rPago === true || String(rPago).toUpperCase() === 'TRUE');

            if (rDni === nd && (!rFecha || rFecha === "") && tienePago) {
                filaSaldo = i + 2;
                break;
            }
        }
    }

    if (filaSaldo > 0) {
        sh.getRange(filaSaldo, 2).setValue(d);
        sh.getRange(filaSaldo, 3).setValue(hora);
        return { ok: true, reasignado: true };
    } else {
        sh.appendRow([new Date(), d, hora, nd, String(nombre || ''), null]);
        return { ok: true, reasignado: false };
    }
}
function removeReservaRow_(dni, fechaMs, horaText) {
    const sh = _sh_(SHEETS.RESERVAS);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return false;
    const lastCol = Math.max(sh.getLastColumn(), 9);
    const vals = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();           // A..I
    const targetDni = _normDni_(dni);
    const targetDate = _dYMD_(fechaMs);
    const targetHora = _normHoraStr_(horaText);

    for (let i = vals.length - 1; i >= 0; i--) {
        const B = vals[i][1];                   // Fecha (Date)
        const Cval = _normHoraStr_(vals[i][2]); // Hora normalizada directamente del valor crudo
        const D = _normDni_(vals[i][3]);
        if (B instanceof Date && _sameYMD_(B, targetDate) && Cval === targetHora && D === targetDni) {
            const pagoVal = vals[i][8]; // Columna I
            const tienePago = (pagoVal === true || String(pagoVal).toUpperCase() === 'TRUE');
            if (tienePago) {
                sh.getRange(2 + i, 2).clearContent(); // Fecha
                sh.getRange(2 + i, 3).clearContent(); // Hora
                return { ok: true, accion: 'limpiado' };
            } else {
                sh.deleteRow(2 + i);
                return { ok: true, accion: 'eliminado' };
            }
        }
    }
    return false;
}
function countReservasMesPorDni_(dni, year, month) {
    if (MAX_RESERVAS_POR_MES <= 0) return 0;
    const sh = _sh_(SHEETS.RESERVAS);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return 0;
    const vals = sh.getRange(2, 1, lastRow - 1, 5).getValues();
    const targetDni = _normDni_(dni);
    let n = 0;
    for (const r of vals) {
        const B = r[1]; const D = _normDni_(r[3]);
        if (D !== targetDni) continue;
        if (B instanceof Date && B.getFullYear() === year && B.getMonth() === month) n++;
    }
    return n;
}
function countReservasActivasPorDni_(dni) {
    if (MAX_RESERVAS_ACTIVAS_POR_DNI <= 0) return 0;
    const sh = _sh_(SHEETS.RESERVAS);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return 0;
    const vals = sh.getRange(2, 1, lastRow - 1, 5).getValues();
    const today = _dYMD_(Date.now());
    const target = _normDni_(dni);
    let n = 0;
    for (const r of vals) {
        const B = r[1]; const D = _normDni_(r[3]);
        if (D !== target) continue;
        if (B instanceof Date && _dYMD_(B.getTime()) >= today) n++;
    }
    return n;
}

/***** API: Config + Login *****/
function pub_getConfig() {
    loadConfig_();
    const fechas = getPublicDates_();

    return {
        ok: true,
        version: VERSION,
        publicOn: !!PUBLIC_ON,
        requirePermitido: !!REQUIRE_PERMITIDO,
        medOn: !!MED_ON,
        cancelOn: !!CANCEL_ON,
        listaEsperaOn: !!LISTA_ESPERA_ON,
        fechas,
        limits: {
            porMes: MAX_RESERVAS_POR_MES,
            activas: MAX_RESERVAS_ACTIVAS_POR_DNI,
            rateSeconds: RATE_SECONDS
        },
        info: INFO_CLINIC
    };
}
function pub_login(dni) {
    loadConfig_();
    const nd = _normDni_(dni);
    if (!nd) return { ok: false, error: 'DNI inválido' };
    const leg = getLegajoByDni_(nd);
    if (!leg) {
        pub_log('PUB_LOGIN_FAIL', { dni: nd, result: 'not_found' });
        return { ok: false, error: 'No estás registrado en Legajos.' };
    }
    const acc = String(leg.acceso || '').toLowerCase().trim();
    if (REQUIRE_PERMITIDO && acc !== 'permitido' && acc !== 'nuevo' && acc !== 'no permitido') {
        pub_log('PUB_LOGIN_DENY', { dni: nd, result: 'access_denied' });
        return { ok: false, error: 'Acceso no permitido por el momento.' };
    }
    pub_log('PUB_LOGIN_OK', { dni: nd, result: 'ok' });
    return {
        ok: true,
        perfil: {
            acceso: leg.acceso,
            dni: _normDni_(leg.dni),
            nombre: leg.nombre,
            telefono: leg.telefono,
            domicilio: leg.domicilio,
            fr1: leg.fr1, fr2: leg.fr2, fr3: leg.fr3, fr4: leg.fr4, fr5: leg.fr5
        }
    };
}

/***** API: Horarios por fecha (columna) *****/
function pub_listHorasForCol(col, dni) {
    loadConfig_();
    if (!PUBLIC_ON) return { ok: false, error: 'Turnero cerrado por el momento.' };
    if (!PUBLIC_DATE_COLS.includes(Number(col))) return { ok: false, error: 'Fecha no habilitada.' };
    
    const nd = _normDni_(dni);
    const leg = getLegajoByDni_(nd);
    const isNuevo = leg && String(leg.acceso || '').toLowerCase().trim() === 'nuevo';
    
    const horas = listHorasForCol_(Number(col), isNuevo);
    return { ok: true, horas };
}

/***** API: Reservar (Lock + validaciones) *****/
function pub_reservar(dateMs, horaText, dni, nombre) {
    loadConfig_();
    const nd = _normDni_(dni);
    if (!PUBLIC_ON) return { ok: false, error: 'Turnero cerrado por el momento.' };
    if (!nd) return { ok: false, error: 'DNI inválido.' };
    if (!rateLimit_(nd, 'reserve', RATE_SECONDS)) return { ok: false, error: 'Demasiados intentos, probá en unos segundos.' };

    // límites
    const d = _dYMD_(dateMs);
    if (MAX_RESERVAS_POR_MES > 0) {
        const nMes = countReservasMesPorDni_(nd, d.getFullYear(), d.getMonth());
        if (nMes >= MAX_RESERVAS_POR_MES) {
            pub_log('PUB_RESERVAR_FAIL', { dni: nd, result: 'limit_mes' });
            return { ok: false, error: 'Alcanzaste el límite mensual de reservas.' };
        }
    }
    if (MAX_RESERVAS_ACTIVAS_POR_DNI > 0) {
        const nAct = countReservasActivasPorDni_(nd);
        if (nAct >= MAX_RESERVAS_ACTIVAS_POR_DNI) {
            pub_log('PUB_RESERVAR_FAIL', { dni: nd, result: 'limit_activas' });
            return { ok: false, error: 'Alcanzaste el límite de reservas activas.' };
        }
    }

    const leg = getLegajoByDni_(nd);
    const isNuevo = leg && String(leg.acceso || '').toLowerCase().trim() === 'nuevo';

    // LOCK + revalidación de slot
    const lock = LockService.getScriptLock();
    lock.waitLock(10 * 1000);
    try {
        const col = findColByDateMs_(dateMs);
        if (!col) return { ok: false, error: 'Fecha inválida.' };
        const row = findHoraRow_(_normHoraStr_(horaText));
        if (!row) return { ok: false, error: 'Hora inválida.' };

        const sh = _sh_(SHEETS.DISP);
        const st = String(sh.getRange(row, col).getDisplayValue() || '').trim();
        const stClean = st.toLowerCase().trim();

        if (isNuevo) {
            if (stClean !== 'nuevos') {
                pub_log('PUB_RESERVAR_FAIL', { dni: nd, result: 'no_permitido_nuevo', fechaMs: dateMs, hora: _normHoraStr_(horaText) });
                return { ok: false, error: 'Este horario no está disponible para nuevos pacientes.' };
            }
        } else {
            if (stClean !== '') {
                pub_log('PUB_RESERVAR_FAIL', { dni: nd, result: 'ocupado', fechaMs: dateMs, hora: _normHoraStr_(horaText) });
                return { ok: false, error: 'Ese horario ya fue tomado o no está disponible.' };
            }
        }

        sh.getRange(row, col).setValue('Reservado');
        const resAppend = appendReserva_(dateMs, _normHoraStr_(horaText), nd, String(nombre || ''));

        pub_log('PUB_RESERVAR_OK', { dni: nd, result: 'ok', fechaMs: dateMs, hora: _normHoraStr_(horaText), reasignado: resAppend.reasignado });
        return { ok: true, fechaMs: dateMs, hora: _normHoraStr_(horaText), reasignado: resAppend.reasignado };
    } finally {
        try { lock.releaseLock(); } catch (_) { }
    }
}

/***** API: Mis reservas (futuras) *****/
function pub_misReservas(dni) {
    loadConfig_();
    const nd = _normDni_(dni);
    const sh = _sh_(SHEETS.RESERVAS);
    const last = sh.getLastRow();
    if (last < 2) return { ok: true, reservas: [] };

    const vals = sh.getRange(2, 1, last - 1, 6).getValues();           // A..F

    const today = _dYMD_(Date.now());
    const out = [];
    for (let i = 0; i < vals.length; i++) {
        const fecha = vals[i][1];                 // Date
        const horaDisp = _normHoraStr_(vals[i][2]);
        const d = _normDni_(vals[i][3]);
        const nombre = _strOrEmpty_(vals[i][4]);
        if (d !== nd) continue;
        if (!(fecha instanceof Date)) continue;
        if (_dYMD_(fecha.getTime()) < today) continue;
        out.push({ fechaMs: fecha.getTime(), hora: horaDisp, nombre });
    }
    out.sort((a, b) => a.fechaMs - b.fechaMs || String(a.hora).localeCompare(String(b.hora)));
    return { ok: true, reservas: out };
}

/***** API: Cancelar (gate + Lock) *****/
function pub_cancelar(dni, fechaMs, horaText) {
    loadConfig_();
    if (!CANCEL_ON) return { ok: false, error: 'La cancelación está deshabilitada por el momento.' };
    const nd = _normDni_(dni);
    if (!rateLimit_(nd, 'cancel', RATE_SECONDS)) return { ok: false, error: 'Demasiados intentos, probá en unos segundos.' };

    const lock = LockService.getScriptLock();
    lock.waitLock(10 * 1000);
    try {
        const col = findColByDateMs_(fechaMs);
        const row = findHoraRow_(_normHoraStr_(horaText));
        if (!col || !row) return { ok: false, error: 'Fecha/Hora inválida.' };

        const disp = _sh_(SHEETS.DISP);
        const st = String(disp.getRange(row, col).getDisplayValue() || '').trim();
        if (st !== '') disp.getRange(row, col).setValue('');

        const resCancel = removeReservaRow_(nd, fechaMs, _normHoraStr_(horaText));
        if (!resCancel) {
            pub_log('PUB_CANCELAR_FAIL', { dni: nd, result: 'not_found', fechaMs, hora: _normHoraStr_(horaText) });
            return { ok: false, error: 'No se encontró la reserva.' };
        }

        pub_log('PUB_CANCELAR_OK', { dni: nd, result: 'ok', fechaMs, hora: _normHoraStr_(horaText), accion: resCancel.accion });
        return { ok: true, accion: resCancel.accion };
    } finally {
        try { lock.releaseLock(); } catch (_) { }
    }
}

/***** API: Solicitar medicación (gate + append en Repeticiones) *****/
function pub_solicitarMedicacion(dni) {
    loadConfig_();
    if (!MED_ON) return { ok: false, error: 'La solicitud de medicación está temporalmente deshabilitada.' };
    const nd = _normDni_(dni);
    const leg = getLegajoByDni_(nd);
    if (!leg) return { ok: false, error: 'No se encontró el legajo.' };

    const sh = _ss_().getSheetByName(SHEETS.REPETICIONES) || _ss_().insertSheet(SHEETS.REPETICIONES);
    if (sh.getLastRow() === 0) {
        sh.appendRow(['DNI', 'NOMBRE', 'FRASCO 1', 'FRASCO 2', 'FRASCO 3', 'FRASCO 4', 'FECHA Y HORA']);
    }
    const fr1 = leg.fr1 || ''; const fr2 = leg.fr2 || ''; const fr3 = leg.fr3 || ''; const fr4 = leg.fr4 || '';
    sh.appendRow([nd, String(leg.nombre || ''), fr1, fr2, fr3, fr4, new Date()]);
    pub_log('PUB_MEDICACION', { dni: nd, result: 'ok' });
    return { ok: true };
}

/***** API: Actualizar contacto (tel y domicilio si existe) *****/
function pub_updateContacto(dni, telefono, domicilio) {
    loadConfig_();
    const nd = _normDni_(dni);
    const { headers, idx, sh } = getLegajosHeaderIndexes_();
    const rowInfo = findLegajoRowByDni_(nd);
    if (!rowInfo) return { ok: false, error: 'No se encontró el legajo.' };

    const row = rowInfo.row;
    const colTel = idx(COLS_LEGAJOS.tel) + 1;
    if (colTel > 0) sh.getRange(row, colTel).setValue(String(telefono || ''));

    const colDom = idx(COLS_LEGAJOS.domicilio) + 1;
    if (colDom > 0 && domicilio != null) sh.getRange(row, colDom).setValue(String(domicilio || ''));

    pub_log('PUB_UPDATE_CONTACTO', { dni: nd, result: 'ok' });
    return { ok: true };
}

/***** API: Registro de nuevo usuario *****/
function pub_registerUser(payload) {
    loadConfig_();
    // rate limit suave para evitar spam
    const dniRaw = payload?.dni;
    const nd = _normDni_(dniRaw);
    if (!rateLimit_(nd || 'anon', 'register', 10)) {
        return { ok: false, error: 'Intentá nuevamente en unos segundos.' };
    }

    // Validaciones
    const err = (msg) => ({ ok: false, error: msg });

    if (!nd || nd.length < 7) return err('DNI inválido.');
    const apellido = String(payload?.apellido || '').trim().toUpperCase();
    const nombre = String(payload?.nombre || '').trim().toUpperCase();
    const contacto = String(payload?.contacto || '').replace(/\s+/g, '').replace(/[^\d+]/g, '');
    const domicilio = String(payload?.domicilio || '').trim().toUpperCase();
    const localidad = String(payload?.localidad || '').trim().toUpperCase();
    const tipo = String(payload?.tipo || '').trim().toLowerCase();

    if (!apellido) return err('Apellido es obligatorio.');
    if (!nombre) return err('Nombre es obligatorio.');
    if (!contacto || !/^\+?\d{7,15}$/.test(contacto)) return err('Contacto inválido.');
    if (!domicilio) return err('Domicilio es obligatorio.');
    if (!localidad) return err('Localidad es obligatoria.');
    if (!['primera atencion', 'paciente en tratamiento'].includes(tipo)) return err('Tipo inválido.');

    // Duplicado DNI
    if (findLegajoRowByDni_(nd)) return err('Este DNI ya está registrado.');

    const { headers, idx, sh, lastCol } = getLegajosHeaderIndexes_();
    const accesoVal = (tipo === 'primera atencion') ? 'no permitido' : 'en tratamiento';
    const nombreFull = `${apellido} ${nombre}`.replace(/\s+/g, ' ').trim();
    const domFull = `${domicilio}, ${localidad}`.replace(/\s+,/g, ',').replace(/,\s+/g, ', ').trim();

    const lock = LockService.getScriptLock();
    lock.waitLock(10 * 1000);
    try {
        // Preparar fila vacía con longitud lastCol
        const rowVals = new Array(Math.max(5, lastCol)).fill('');

        const setByHeader = (hName, value) => {
            const i = idx(hName);
            if (i >= 0) rowVals[i] = value;
            return i >= 0;
        };

        // Escribir por nombre (si no existe header, fallback por posición básica A..E)
        const wroteAcc = setByHeader(COLS_LEGAJOS.acceso, accesoVal);
        const wrotePwd = setByHeader(COLS_LEGAJOS.password, ''); // vacío
        const wroteDni = setByHeader(COLS_LEGAJOS.dni, nd);
        const wroteNom = setByHeader(COLS_LEGAJOS.nombre, nombreFull);
        const wroteTel = setByHeader(COLS_LEGAJOS.tel, contacto);
        // Domicilio si existe la columna
        setByHeader(COLS_LEGAJOS.domicilio, domFull);

        if (!(wroteAcc && wroteDni && wroteNom && wroteTel)) {
            // fallback en las primeras 5 columnas (A..E)
            rowVals[0] = accesoVal;  // A Acceso
            rowVals[1] = '';         // B Password (vacío)
            rowVals[2] = nd;         // C DNI
            rowVals[3] = nombreFull; // D Apellido (Apellido + Nombre)
            rowVals[4] = contacto;   // E Telefono
        }

        // Append
        sh.appendRow(rowVals.slice(0, lastCol));

        pub_log('PUB_REGISTER_OK', { dni: nd, result: 'ok', tipo });
        const msg = (tipo === 'primera atencion') ? 'Registro completado. Ya podés iniciar sesión con tu DNI.' : 'Registro enviado. Será evaluado por el consultorio.';
        return { ok: true, message: msg };
    } finally {
        try { lock.releaseLock(); } catch (_) { }
    }
}


function pub_getFechaMedicacion() {
    try {
        const sh = _sh_(SHEETS.DISP); // "Disponibilidad"
        const rng = sh.getRange('I18');
        const v = rng.getValue();
        const disp = String(rng.getDisplayValue() || '').trim();
        let fechaTxt = '';
        if (v instanceof Date && !isNaN(v.getTime())) {
            fechaTxt = disp || Utilities.formatDate(v, _tz_(), 'dd/MM/yyyy');
        } else {
            fechaTxt = disp || ''; // por si I18 es texto ya formateado
        }
        return { ok: true, fecha: fechaTxt };
    } catch (e) {
        return { ok: false, error: String(e && e.message || e) };
    }
}

/***** API: Precios de Medicación *****/
function getPrecioDesdeNombreYCodigo(medicamentoRaw, codigoRaw, customPrices) {
    const norm = s => String(s || "").toUpperCase().trim().replace(/\s+/g, " ");

    const med = norm(medicamentoRaw);
    const cod = norm(codigoRaw);

    if (!med && !cod) return 0;

    // ---- Detectores ----
    const esX60 = /\bX\s*60\s*COMP\b/.test(med) || /\bX\s*60\s*COMP\b/.test(cod);

    // ¿Código común/residual?
    const esComunCode = (s) => {
        if (!s) return false;
        const x = norm(s);
        return /^(603|332|605|205|204|230|19A|19B?)($|[\/\s])/.test(x);
    };

    const codEsComun = esComunCode(cod);

    const medEsSerieM = /^M[0-4](?:\b|$)/.test(med);
    const medEsFentermina = med.includes("FENTERMINA 30MG");
    const medEsFTD = med.startsWith("FTD");
    const medEsFT = (med === "FT" || med.startsWith("FT "));
    const medEsFastin = med.includes("FASTIN");
    const medEsComun = esComunCode(med) || /^\d{3}\/[A-Z0-9/]+$/.test(med); // p.ej. "603/A"

    // Acompañada = cuando el código es una medicación COMÚN/RESIDUAL
    const acompanada = !!codEsComun;

    const table = customPrices || {};
    const defRow = table.default || table.comun || table.serieM || [0, 0, 0, 0];
    const pFent = table.fentermina || defRow;
    const pFast = table.fastin || defRow;
    const pFT = table.ft || defRow;
    const pSerieM = table.serieM || defRow;
    const pComun = table.comun || defRow;
    const pDefault = defRow;

    // ---- Tabla de precios por tipo ----
    const precioEspecial = (solo30, solo60, acomp30, acomp60) => {
        if (acompanada) return esX60 ? acomp60 : acomp30;
        return esX60 ? solo60 : solo30;
    };

    // 1) FENTERMINA 30MG y FTD (mismo esquema de precios)
    if (medEsFentermina || medEsFTD) {
        return precioEspecial(pFent[0], pFent[1], pFent[2], pFent[3]);
    }

    // 2) FASTIN
    if (medEsFastin) {
        return precioEspecial(pFast[0], pFast[1], pFast[2], pFast[3]);
    }

    // 3) FT (cuidar orden: FTD ya tratado arriba)
    if (medEsFT) {
        return precioEspecial(pFT[0], pFT[1], pFT[2], pFT[3]);
    }

    // 4) Serie M (M0..M4)
    if (medEsSerieM) {
        return precioEspecial(pSerieM[0], pSerieM[1], pSerieM[2], pSerieM[3]);
    }

    // 5) COMÚN/RESIDUAL (603, 603/…, 205/1, 204, 230, 19, 19B, 605, etc.)
    if (medEsComun) {
        return precioEspecial(pComun[0], pComun[1], pComun[2], pComun[3]);
    }

    // 6) Default (si aparece algo fuera de catálogo)
    return precioEspecial(pDefault[0], pDefault[1], pDefault[2], pDefault[3]);
}

function getPreciosSheet_(ss) {
    let shP = ss.getSheetByName('PreciosMedicacion');
    if (shP) return shP;

    const sheets = ss.getSheets();
    for (const s of sheets) {
        const n = s.getName().trim().toLowerCase();
        if (n.includes('precio') || n.includes('medicacion') || n.includes('tarifa') || n.includes('costo')) {
            return s;
        }
    }
    shP = ss.insertSheet('PreciosMedicacion');
    shP.appendRow(['Categoría', 'Solo 30', 'Solo 60', 'Acompañada 30', 'Acompañada 60']);
    shP.appendRow(['CONSULTA', 25000, 25000, 25000, 25000]);
    shP.appendRow(['FENTERMINA', 50000, 50000, 50000, 50000]);
    shP.appendRow(['FASTIN', 50000, 50000, 50000, 50000]);
    shP.appendRow(['FT', 50000, 50000, 50000, 50000]);
    shP.appendRow(['SERIE M', 50000, 50000, 50000, 50000]);
    shP.appendRow(['COMUN', 50000, 50000, 50000, 50000]);
    shP.appendRow(['DEFAULT', 50000, 50000, 50000, 50000]);
    return shP;
}

function parseNum_(val, defaultVal = 0) {
    if (typeof val === 'number') return isNaN(val) ? defaultVal : val;
    if (!val) return defaultVal;
    let s = String(val).trim().replace(/\$/g, '').replace(/\s+/g, '');
    if (/^\d{1,3}(\.\d{3})+$/.test(s)) {
        s = s.replace(/\./g, '');
    } else if (/^\d{1,3}(,\d{3})+$/.test(s)) {
        s = s.replace(/,/g, '');
    } else {
        s = s.replace(/,/g, '.');
    }
    const n = parseFloat(s);
    return isNaN(n) ? defaultVal : n;
}

function getPreciosFromSheet_() {
    try {
        const ss = _ss_();
        const shP = getPreciosSheet_(ss);
        if (shP) {
            const data = shP.getDataRange().getValues();
            let customPrices = {};
            let precioConsulta = 25000;

            for (let r = 0; r < data.length; r++) {
                const row = data[r];
                if (!row || row.length === 0) continue;

                for (let c = 0; c < row.length; c++) {
                    const cat = String(row[c] || '').toUpperCase().replace(/[\s\u00A0]+/g, ' ').trim();
                    if (!cat) continue;

                    let nums = [];
                    for (let k = c + 1; k < row.length; k++) {
                        const n = parseNum_(row[k], -1);
                        if (n >= 0) nums.push(n);
                        if (nums.length >= 4) break;
                    }

                    if (nums.length === 0) continue;

                    const s30 = nums[0] || 0;
                    const s60 = nums.length > 1 ? nums[1] : s30;
                    const a30 = nums.length > 2 ? nums[2] : s30;
                    const a60 = nums.length > 3 ? nums[3] : s60;

                    if (cat.includes('CONSULTA')) {
                        precioConsulta = s30 || 25000;
                    } else if (cat.includes('FENTERMINA')) {
                        customPrices.fentermina = [s30, s60, a30, a60];
                    } else if (cat.includes('FASTIN')) {
                        customPrices.fastin = [s30, s60, a30, a60];
                    } else if (cat === 'FT' || cat === 'FTD' || cat.startsWith('FT ') || cat.startsWith('FTD')) {
                        customPrices.ft = [s30, s60, a30, a60];
                    } else if (cat.includes('SERIE') || cat.includes('M0') || cat.includes('M1') || cat.includes('M2') || cat === 'M') {
                        customPrices.serieM = [s30, s60, a30, a60];
                    } else if (cat.includes('COMUN') || cat.includes('RESIDUAL')) {
                        customPrices.comun = [s30, s60, a30, a60];
                    } else if (cat.includes('DEFAULT') || cat.includes('OTRO')) {
                        customPrices.default = [s30, s60, a30, a60];
                    }
                }
            }
            return { precioConsulta, customPrices, sheetName: shP.getName() };
        }
    } catch (e) {
        console.error("Error al leer hoja PreciosMedicacion: " + e);
    }
    return {
        precioConsulta: 25000,
        customPrices: null,
        sheetName: 'Ninguna'
    };
}

function pub_calcularPrecioMedicacion(dni) {
    loadConfig_();
    const nd = _normDni_(dni);
    if (!nd) return { ok: false, error: 'DNI inválido' };
    const leg = getLegajoByDni_(nd);
    if (!leg) return { ok: false, error: 'No se encontró el legajo.' };

    const rawMeds = [leg.fr1, leg.fr2, leg.fr3, leg.fr4, leg.fr5]
        .map(x => String(x || '').trim())
        .filter(Boolean);

    // Separar posibles múltiples medicamentos unidos por puntos (· o .), comas o punto y coma dentro del mismo frasco
    let parsedTokens = [];
    rawMeds.forEach(x => {
        const parts = x.split(/[·.,;]+/).map(p => p.trim()).filter(Boolean);
        parsedTokens.push(...parts);
    });

    const { precioConsulta, customPrices } = getPreciosFromSheet_();

    if (parsedTokens.length === 0) {
        return {
            ok: true,
            totalMedicacion: 0,
            precioConsulta: precioConsulta,
            totalFinal: precioConsulta,
            medicacionStr: '(Sin medicación registrada)',
            mensaje: 'No tenés medicación registrada en tu legajo.'
        };
    }

    const norm = s => String(s || "").toUpperCase().trim().replace(/\s+/g, " ");
    const esComunCode = (s) => {
        if (!s) return false;
        const x = norm(s);
        return /^(603|332|605|205|204|230|19A|19B?)($|[\/\s])/.test(x);
    };

    const comunCodes = parsedTokens.filter(m => esComunCode(m));
    const mainMeds = parsedTokens.filter(m => !esComunCode(m));

    let totalMedicacion = 0;
    const acompanada = comunCodes.length > 0;

    if (mainMeds.length > 0) {
        // Para la primera medicación principal, se le pasa el código común acompañante si existe
        const firstCod = acompanada ? comunCodes[0] : '';
        totalMedicacion += getPrecioDesdeNombreYCodigo(mainMeds[0], firstCod, customPrices);

        // Si hay medicaciones principales adicionales, se suman sus precios individuales
        for (let i = 1; i < mainMeds.length; i++) {
            totalMedicacion += getPrecioDesdeNombreYCodigo(mainMeds[i], '', customPrices);
        }
    } else if (comunCodes.length > 0) {
        // Si solo tiene código común (sin medicación principal)
        totalMedicacion += getPrecioDesdeNombreYCodigo(comunCodes[0], '', customPrices);
        for (let i = 1; i < comunCodes.length; i++) {
            totalMedicacion += getPrecioDesdeNombreYCodigo(comunCodes[i], '', customPrices);
        }
    }

    return {
        ok: true,
        medicacionStr: rawMeds.join(' · '),
        totalMedicacion: totalMedicacion,
        precioConsulta: precioConsulta,
        totalFinal: totalMedicacion + precioConsulta,
        acompanada: acompanada
    };
}


/***** API: Lista de Espera *****/
function pub_checkWaitingList(dni) {
    loadConfig_();
    const nd = _normDni_(dni);
    if (!nd) return { registrado: false };
    
    try {
        const ss = _ss_();
        let sh = ss.getSheetByName('Lista espera');
        if (!sh) return { registrado: false };
        
        const lastRow = sh.getLastRow();
        if (lastRow < 2) return { registrado: false };
        
        // El DNI ahora se encuentra en la columna 3 (C)
        const dnis = sh.getRange(2, 3, lastRow - 1, 1).getValues().map(r => _normDni_(r[0]));
        return { registrado: dnis.includes(nd) };
    } catch (e) {
        console.error("Error checking waiting list: " + e);
        return { registrado: false };
    }
}

function pub_joinWaitingList(dni, nombre, telefono) {
    loadConfig_();
    const nd = _normDni_(dni);
    if (!nd) return { ok: false, error: 'DNI inválido' };
    
    const lock = LockService.getScriptLock();
    lock.waitLock(10 * 1000);
    try {
        const ss = _ss_();
        let sh = ss.getSheetByName('Lista espera');
        if (!sh) {
            sh = ss.insertSheet('Lista espera');
            sh.appendRow(['Acceso', 'Registro', 'DNI', 'Nombre', 'Telefono', 'Ultima atencion', 'Estado', 'fecha solicitud']);
        }
        
        const lastRow = sh.getLastRow();
        if (lastRow >= 2) {
            // Comprobamos duplicados buscando en la columna 3 (C)
            const dnis = sh.getRange(2, 3, lastRow - 1, 1).getValues().map(r => _normDni_(r[0]));
            if (dnis.includes(nd)) {
                return { ok: true, mensaje: 'Ya anotado' };
            }
        }
        
        const leg = getLegajoByDni_(nd);
        const accesoVal = leg ? (leg.acceso || 'no permitido') : 'no permitido';
        
        // Estructura: Acceso | Registro | DNI | Nombre | Telefono
        sh.appendRow([
        accesoVal,                            // A: Acceso (permitido, no permitido, etc.)
        new Date(),                           // B: Registro (fecha actual)
        nd,                                   // C: DNI
        String(nombre || '').trim(),          // D: Nombre
        String(telefono || '').trim(),        // E: Telefono
        "",                                   // F: Ultima atencion (en blanco)
        "",                                   // G: Estado (en blanco)
        new Date()                            // H: fecha solicitud (fecha y hora exacta)
    ]);

        return { ok: true };
    } catch (e) {
        console.error("Error joining waiting list: " + e);
        return { ok: false, error: String(e && e.message || e) };
    } finally {
        try { lock.releaseLock(); } catch (_) {}
    }
}

