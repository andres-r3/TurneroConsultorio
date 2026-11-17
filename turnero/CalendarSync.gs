/***********************************************************************
 * SINCRONIZACIÓN  Sheets (Disponibilidad / Reservas)  ⇄  Google Calendar
 * Para que el bot de WhatsApp pueda ofrecer y tomar turnos.
 *
 * Va en el MISMO proyecto de Apps Script que el turnero público (Code.gs),
 * porque reutiliza sus funciones: _ss_, _sh_, loadConfig_, getLegajoByDni_,
 * appendReserva_, removeReservaRow_, countReservas*, pub_log, etc.
 *
 * Idea general
 *  - La planilla sigue siendo la fuente de verdad.
 *  - En un calendario DEDICADO se crean eventos "No disponible" que cubren
 *    todo lo que NO es un turno libre. El bot ve como libre exactamente los
 *    turnos libres de la grilla (y solo en las fechas publicadas).
 *  - Cuando el bot crea un evento (reserva por WhatsApp), este script:
 *      identifica al paciente (DNI o teléfono), aplica las MISMAS reglas que
 *      el turnero y, si todo está bien, marca "Reservado" en Disponibilidad
 *      y agrega la fila en Reservas. Si no, rechaza el evento.
 *  - Si el bot borra / mueve el evento → se cancela / reprograma en la planilla.
 *  - Si el turno se cancela en la planilla → se borra el evento del bot.
 *
 * Instalación: ver README.md (ejecutar calSync_instalar() una vez).
 ***********************************************************************/

const CAL_SYNC = {
    CALENDAR_ID: '',                         // ID de un calendario existente. Vacío = usa/crea CALENDAR_NAME
    CALENDAR_NAME: 'Turnos Consultorio (bot)',
    DIAS_HORIZONTE: 60,                      // cuántos días hacia adelante se bloquean/sincronizan
    SOLO_FECHAS_PUBLICAS: true,              // igual que el turnero: solo las fechas de PUBLIC_DATE_COLS (B1, C1)
    ESTADOS_LIBRES: [''],                    // estados de una celda que el bot puede ofrecer ('' = vacía)
    NUEVOS_PARA_BOT: false,                  // true = las celdas "nuevos" se muestran libres (solo las podrán tomar pacientes "nuevo")
    DURACION_TURNO_MIN: 0,                   // 0 = se deduce de la columna A (menor diferencia entre horas)
    ACCION_RECHAZO: 'eliminar',              // 'eliminar' = borra el evento del bot | 'marcar' = lo deja con título [RECHAZADO]
    PREFIJO_IGNORAR: '#',                    // eventos cuyo título empieza con esto no se tocan nunca
    EMAIL_AVISOS: '',                        // si se completa, se envía un mail por cada reserva rechazada / cancelación denegada
    TITULO_BLOQUE: 'No disponible',
    HOJA_SYNC: 'SyncCalendar',               // hoja de registro de reservas hechas por el bot
    TOLERANCIA_MIN: 1                        // minutos de tolerancia en la duración del evento del bot
};

const CAL_SYNC_TAG = 'turnero';              // tag privado en los eventos: 'bloque' | 'reserva' | 'rechazado'
const CAL_SYNC_COLS = ['EventId', 'Estado', 'DNI', 'Nombre', 'Fecha', 'Hora', 'InicioMs', 'Actualizado', 'Detalle'];

/***** Funciones para ejecutar a mano desde el editor *****/

/** Crea (si hace falta) el calendario, instala los activadores y hace la primera sincronización. */
function calSync_instalar() {
    const cal = calSync_getCalendar_();
    calSync_borrarActivadores_();
    ScriptApp.newTrigger('syncTurnosCalendar').timeBased().everyMinutes(5).create();
    let triggerCalendario = true;
    try {
        ScriptApp.newTrigger('syncTurnosCalendar').forUserCalendar(cal.getId()).onEventUpdated().create();
    } catch (e) {
        triggerCalendario = false;
        console.warn('No se pudo crear el activador de Calendar (queda solo el de 5 minutos): ' + e);
    }
    PropertiesService.getScriptProperties().setProperty('CAL_SYNC_ACTIVO', '1');
    const resumen = syncTurnosCalendar();
    const info = { calendario: cal.getName(), calendarId: cal.getId(), triggerCalendario, resumen };
    console.log(JSON.stringify(info, null, 2));
    return info;
}

/** Quita los activadores y desactiva la sincronización (no borra eventos). */
function calSync_desinstalar() {
    calSync_borrarActivadores_();
    PropertiesService.getScriptProperties().deleteProperty('CAL_SYNC_ACTIVO');
    console.log('Sincronización desactivada.');
}

/** Borra todos los bloques "No disponible" creados por este script (no toca reservas del bot). */
function calSync_borrarBloques() {
    const cal = calSync_getCalendar_();
    const desde = new Date(Date.now() - 400 * 864e5);
    const hasta = new Date(Date.now() + 400 * 864e5);
    let n = 0;
    cal.getEvents(desde, hasta).forEach(ev => {
        if (calSync_tag_(ev) === 'bloque') { ev.deleteEvent(); n++; }
    });
    console.log(`Bloques borrados: ${n}`);
    return n;
}

/***** Punto de entrada de los activadores *****/
function syncTurnosCalendar() {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20 * 1000)) return { ok: false, motivo: 'ocupado' };
    try {
        return calSync_ejecutar_();
    } finally {
        try { lock.releaseLock(); } catch (_) { }
    }
}

/**
 * Llamado desde pub_reservar / pub_cancelar (que ya tienen el lock tomado)
 * para que el calendario se actualice al instante. Nunca rompe la reserva.
 */
function calSync_trasCambioEnSheet_() {
    try {
        if (PropertiesService.getScriptProperties().getProperty('CAL_SYNC_ACTIVO') !== '1') return;
        calSync_ejecutar_();
    } catch (e) {
        console.error('calSync_trasCambioEnSheet_: ' + e);
    }
}

/***** Núcleo *****/
function calSync_ejecutar_() {
    loadConfig_();
    const cal = calSync_getCalendar_();
    const g = calSync_leerGrilla_();
    const tz = g.tz;
    const hoyYmd = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const desde = Utilities.parseDate(hoyYmd + ' 00:00', tz, 'yyyy-MM-dd HH:mm').getTime();
    const hasta = desde + (CAL_SYNC.DIAS_HORIZONTE + 1) * 864e5;

    const eventos = cal.getEvents(new Date(desde), new Date(hasta));
    const porId = {};
    eventos.forEach(ev => { porId[ev.getId()] = ev; });

    const mapa = calSync_leerMapa_();
    const reservas = calSync_reservasSet_(tz);
    const cubiertos = new Set();     // slots "ymd|hora" ocupados por un evento del bot ya importado
    const idsActivos = new Set();
    const reimportar = new Set();
    const r = { importadas: 0, rechazadas: 0, canceladasPorBot: 0, canceladasEnSheet: 0, bloquesCreados: 0, bloquesBorrados: 0 };

    // 1) Reservas del bot ya registradas: ¿siguen vigentes en ambos lados?
    mapa.filter(m => m.estado === 'activa').forEach(m => {
        if (m.inicioMs && m.inicioMs < desde) { calSync_setEstado_(m, 'pasada'); return; }
        const ev = porId[m.eventId];
        const key = m.ymd + '|' + m.hora;

        if (!ev) {                                        // el bot borró el evento → cancelar
            if (calSync_cancelarDesdeBot_(m, g, reservas, 'Evento eliminado en Calendar')) r.canceladasPorBot++;
            return;
        }
        if (ev.getStartTime().getTime() !== m.inicioMs) { // el bot lo movió → cancelar y reimportar
            if (calSync_cancelarDesdeBot_(m, g, reservas, 'Evento reprogramado en Calendar')) {
                r.canceladasPorBot++;
                ev.deleteTag(CAL_SYNC_TAG);
                reimportar.add(m.eventId);
            } else {
                calSync_rechazar_(ev, 'La cancelación/reprogramación está deshabilitada; se mantiene el turno original', { dni: m.dni, nombre: m.nombre }, r);
                delete porId[m.eventId];
            }
            return;
        }
        if (!reservas.has(m.dni + '|' + key)) {           // cancelado/reasignado en la planilla → borrar evento
            ev.deleteEvent();
            delete porId[m.eventId];
            calSync_setEstado_(m, 'cancelada_en_sheet');
            r.canceladasEnSheet++;
            return;
        }
        cubiertos.add(key);
        idsActivos.add(m.eventId);
    });

    // 2) Eventos nuevos del bot → validar e importar (por orden de creación)
    const ahora = Date.now();
    const nuevos = Object.values(porId).filter(ev => {
        const id = ev.getId();
        if (idsActivos.has(id)) return false;
        if (!reimportar.has(id) && calSync_tag_(ev)) return false;
        if (ev.isAllDayEvent()) return false;
        if (String(ev.getTitle() || '').startsWith(CAL_SYNC.PREFIJO_IGNORAR)) return false;
        if (ev.getStartTime().getTime() < ahora) return false;
        return true;
    }).sort((a, b) => a.getDateCreated().getTime() - b.getDateCreated().getTime());

    nuevos.forEach(ev => {
        const res = calSync_importar_(ev, g, reservas, r);
        if (res.ok) {
            cubiertos.add(res.key);
            delete porId[ev.getId()];
        }
    });

    // 3) Bloques "No disponible" = todo lo que no es turno libre para el bot
    const libres = [];
    Object.values(g.dias).forEach(dia => {
        const habilitado = !CAL_SYNC.SOLO_FECHAS_PUBLICAS || dia.publica;
        Object.values(dia.slots).forEach(s => {
            if (s.endMs <= desde || s.startMs >= hasta) return;
            const libre = habilitado && PUBLIC_ON && calSync_libreParaBot_(s.estado);
            if (libre || cubiertos.has(s.ymd + '|' + s.hora)) {
                libres.push([Math.max(s.startMs, desde), Math.min(s.endMs, hasta)]);
            }
        });
    });
    libres.sort((a, b) => a[0] - b[0]);
    const bloques = [];
    let cursor = desde;
    libres.forEach(([a, b]) => {
        if (a > cursor) bloques.push([cursor, a]);
        cursor = Math.max(cursor, b);
    });
    if (cursor < hasta) bloques.push([cursor, hasta]);

    const quiero = new Map(bloques.map(([a, b]) => [a + '-' + b, [a, b]]));
    Object.values(porId).forEach(ev => {
        if (calSync_tag_(ev) !== 'bloque') return;
        const k = ev.getStartTime().getTime() + '-' + ev.getEndTime().getTime();
        if (quiero.has(k)) quiero.delete(k);
        else { ev.deleteEvent(); r.bloquesBorrados++; }
    });
    quiero.forEach(([a, b]) => {
        const ev = cal.createEvent(CAL_SYNC.TITULO_BLOQUE, new Date(a), new Date(b));
        ev.setTag(CAL_SYNC_TAG, 'bloque');
        try { ev.removeAllReminders(); } catch (_) { }
        r.bloquesCreados++;
    });

    console.log('syncTurnosCalendar: ' + JSON.stringify(r));
    return r;
}

/***** Importar una reserva hecha por el bot *****/
function calSync_importar_(ev, g, reservas, r) {
    const tz = g.tz;
    const startMs = ev.getStartTime().getTime();
    const endMs = ev.getEndTime().getTime();
    const ymd = Utilities.formatDate(new Date(startMs), tz, 'yyyy-MM-dd');
    const hora = Utilities.formatDate(new Date(startMs), tz, 'HH:mm');
    const texto = [ev.getTitle(), ev.getDescription(), ev.getLocation()].join('\n');
    const ctx = { ymd, hora, texto };
    const rechazar = (motivo) => { calSync_rechazar_(ev, motivo, ctx, r); return { ok: false }; };

    const dia = g.dias[ymd];
    const slot = dia && dia.slots[hora];
    if (!slot || (CAL_SYNC.SOLO_FECHAS_PUBLICAS && !dia.publica)) return rechazar('El horario no corresponde a un turno publicado');
    if (endMs - startMs > (g.slotMin + CAL_SYNC.TOLERANCIA_MIN) * 60000) return rechazar(`El evento dura más que un turno (${g.slotMin} min)`);
    if (!PUBLIC_ON) return rechazar('Turnero cerrado por el momento');

    const pac = calSync_identificarPaciente_(texto);
    if (!pac.ok) return rechazar(pac.error);
    const leg = pac.leg;
    const dni = _normDni_(leg.dni);
    ctx.dni = dni;
    ctx.nombre = leg.nombre;

    const acc = String(leg.acceso || '').toLowerCase().trim();
    if (REQUIRE_PERMITIDO && acc !== 'permitido' && acc !== 'nuevo' && acc !== 'no permitido') return rechazar('Acceso no permitido');

    const isNuevo = acc === 'nuevo';
    const st = String(slot.estado || '').toLowerCase().trim();
    if (isNuevo ? st !== 'nuevos' : st !== '') {
        return rechazar(isNuevo ? 'Este horario no está disponible para nuevos pacientes' : 'Ese horario ya fue tomado o no está disponible');
    }

    const d = new Date(slot.dateMs);
    if (MAX_RESERVAS_POR_MES > 0 && countReservasMesPorDni_(dni, d.getFullYear(), d.getMonth()) >= MAX_RESERVAS_POR_MES) {
        return rechazar('Alcanzó el límite mensual de reservas');
    }
    if (MAX_RESERVAS_ACTIVAS_POR_DNI > 0 && countReservasActivasPorDni_(dni) >= MAX_RESERVAS_ACTIVAS_POR_DNI) {
        return rechazar('Alcanzó el límite de reservas activas');
    }

    // Confirmar en la planilla (mismo efecto que pub_reservar)
    _sh_(SHEETS.DISP).getRange(slot.row, slot.col).setValue('Reservado');
    slot.estado = 'Reservado';
    const resAppend = appendReserva_(slot.dateMs, hora, dni, String(leg.nombre || ''));
    reservas.add(dni + '|' + ymd + '|' + hora);

    ev.setTag(CAL_SYNC_TAG, 'reserva');
    calSync_hojaMapa_().appendRow([
        ev.getId(), 'activa', dni, String(leg.nombre || ''), "'" + ymd, "'" + hora, "'" + startMs, new Date(),
        'Identificado por ' + pac.via + (resAppend.reasignado ? ' · usó saldo a favor' : '') +
        ((isNuevo || hora <= '10:00') ? ' · requiere transferencia' : '')
    ]);
    pub_log('BOT_RESERVAR_OK', { dni, result: 'ok', fechaMs: slot.dateMs, hora, eventId: ev.getId(), reasignado: resAppend.reasignado });
    r.importadas++;
    return { ok: true, key: ymd + '|' + hora };
}

/***** Rechazo de un evento del bot *****/
function calSync_rechazar_(ev, motivo, ctx, r) {
    const titulo = String(ev.getTitle() || '');
    const fechaTxt = ctx.ymd ? `${ctx.ymd} ${ctx.hora}` : Utilities.formatDate(ev.getStartTime(), calSync_tz_(), 'yyyy-MM-dd HH:mm');
    const texto = String(ctx.texto || [titulo, ev.getDescription()].join('\n')).slice(0, 500);
    const id = ev.getId();

    if (CAL_SYNC.ACCION_RECHAZO === 'marcar') {
        ev.setTag(CAL_SYNC_TAG, 'rechazado');
        if (!titulo.startsWith('[RECHAZADO]')) ev.setTitle('[RECHAZADO] ' + titulo);
        ev.setDescription(`Motivo: ${motivo}\n\n${ev.getDescription() || ''}`);
    } else {
        ev.deleteEvent();
    }

    calSync_hojaMapa_().appendRow([
        id, 'rechazada', ctx.dni || '', ctx.nombre || '', "'" + fechaTxt.slice(0, 10), "'" + fechaTxt.slice(11), '', new Date(),
        motivo + ' | Evento: ' + texto.replace(/\s+/g, ' ')
    ]);
    pub_log('BOT_RESERVAR_RECHAZADO', { dni: ctx.dni || '', result: motivo, fecha: fechaTxt, eventId: id });
    calSync_avisar_(`Turno de WhatsApp rechazado (${fechaTxt})`,
        `Motivo: ${motivo}\nFecha/hora: ${fechaTxt}\nDNI: ${ctx.dni || '-'}\nNombre: ${ctx.nombre || '-'}\n\nDatos del evento:\n${texto}\n\n` +
        'El paciente puede creer que tiene el turno: conviene contactarlo.');
    if (r) r.rechazadas++;
}

/***** Cancelación hecha desde el bot (evento borrado o movido) *****/
function calSync_cancelarDesdeBot_(m, g, reservas, motivo) {
    if (!CANCEL_ON) {
        calSync_setEstado_(m, 'cancelacion_denegada', motivo + ' (cancelación deshabilitada en el turnero)');
        pub_log('BOT_CANCELAR_DENEGADO', { dni: m.dni, result: 'cancel_off', hora: m.hora, fecha: m.ymd });
        calSync_avisar_(`Cancelación por WhatsApp denegada (${m.ymd} ${m.hora})`,
            `${m.nombre} (DNI ${m.dni}) quitó su turno en el bot pero la cancelación está deshabilitada. El turno sigue reservado en la planilla.`);
        return false;
    }
    const dia = g.dias[m.ymd];
    const slot = dia && dia.slots[m.hora];
    if (slot && /^reservado/i.test(String(slot.estado || '').trim())) {
        _sh_(SHEETS.DISP).getRange(slot.row, slot.col).setValue('');
        slot.estado = '';
    }
    const fechaMs = slot ? slot.dateMs : Utilities.parseDate(m.ymd + ' 00:00', g.tz, 'yyyy-MM-dd HH:mm').getTime();
    const res = removeReservaRow_(m.dni, fechaMs, m.hora);
    reservas.delete(m.dni + '|' + m.ymd + '|' + m.hora);
    calSync_setEstado_(m, 'cancelada_por_bot', motivo + (res ? ` (${res.accion})` : ' (no estaba en Reservas)'));
    pub_log('BOT_CANCELAR_OK', { dni: m.dni, result: 'ok', fechaMs, hora: m.hora, accion: res ? res.accion : 'no_encontrada' });
    return true;
}

/***** Identificar al paciente a partir del texto del evento *****/
function calSync_identificarPaciente_(texto) {
    const t = String(texto || '');
    const candidatos = [];
    const etiquetado = t.match(/D\.?\s*N\.?\s*I\.?\D{0,5}(\d[\d.\s]{5,12}\d)/i);
    if (etiquetado) candidatos.push(_normDni_(etiquetado[1]));
    else (t.match(/\b\d{1,2}\.?\d{3}\.?\d{3}\b/g) || []).forEach(x => candidatos.push(_normDni_(x)));

    const unicos = [...new Set(candidatos.filter(d => d.length >= 7 && d.length <= 8))];
    const encontrados = unicos.map(d => getLegajoByDni_(d)).filter(Boolean);
    if (encontrados.length === 1) return { ok: true, leg: encontrados[0], via: 'DNI' };
    if (encontrados.length > 1) return { ok: false, error: 'El evento tiene más de un DNI registrado' };

    // Sin DNI válido: probar por teléfono (últimos 8 dígitos)
    const tels = (t.replace(/[\s\-().]/g, '').match(/\d{8,15}/g) || []).map(x => x.slice(-8));
    if (tels.length) {
        const { sh, idx } = getLegajosHeaderIndexes_();
        const cTel = idx(COLS_LEGAJOS.tel), cDni = idx(COLS_LEGAJOS.dni);
        const last = sh.getLastRow();
        if (cTel >= 0 && cDni >= 0 && last >= 2) {
            const vals = sh.getRange(2, 1, last - 1, Math.max(cTel, cDni) + 1).getDisplayValues();
            const match = new Set();
            vals.forEach(row => {
                const tel = String(row[cTel] || '').replace(/\D+/g, '');
                if (tel.length >= 8 && tels.includes(tel.slice(-8))) match.add(_normDni_(row[cDni]));
            });
            if (match.size === 1) {
                const leg = getLegajoByDni_([...match][0]);
                if (leg) return { ok: true, leg, via: 'teléfono' };
            }
            if (match.size > 1) return { ok: false, error: 'El teléfono corresponde a más de un paciente; falta el DNI' };
        }
    }
    return { ok: false, error: unicos.length ? 'El DNI no está registrado en Legajos' : 'No se encontró un DNI ni un teléfono registrado en el evento' };
}

/***** Lectura de la grilla Disponibilidad *****/
function calSync_leerGrilla_() {
    const tz = calSync_tz_();
    const sh = _sh_(SHEETS.DISP);
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    const g = { tz, slotMin: CAL_SYNC.DURACION_TURNO_MIN || 15, dias: {} };
    if (lastRow < 2 || lastCol < 2) return g;

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const grid = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    const horas = grid.map(row => _normHoraStr_(row[0]));
    const esHora = h => /^\d{2}:\d{2}$/.test(h);

    if (!CAL_SYNC.DURACION_TURNO_MIN) {
        const mins = [...new Set(horas.filter(esHora).map(h => Number(h.slice(0, 2)) * 60 + Number(h.slice(3))))].sort((a, b) => a - b);
        let min = 0;
        for (let i = 1; i < mins.length; i++) {
            const dif = mins[i] - mins[i - 1];
            if (dif > 0 && (!min || dif < min)) min = dif;
        }
        if (min) g.slotMin = min;
    }

    for (let c = 2; c <= lastCol; c++) {
        const v = header[c - 1];
        if (!(v instanceof Date) || isNaN(v.getTime())) continue;
        const ymd = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        const publica = PUBLIC_DATE_COLS.includes(c);
        const previo = g.dias[ymd];
        if (previo && (previo.publica || !publica)) continue; // ante fechas repetidas, prioriza la columna pública

        const dia = { ymd, col: c, dateMs: v.getTime(), publica, slots: {} };
        for (let i = 0; i < horas.length; i++) {
            const h = horas[i];
            if (!esHora(h) || dia.slots[h]) continue;
            const start = Utilities.parseDate(ymd + ' ' + h, tz, 'yyyy-MM-dd HH:mm').getTime();
            dia.slots[h] = {
                row: i + 2, col: c, hora: h, ymd, dateMs: dia.dateMs,
                startMs: start, endMs: start + g.slotMin * 60000,
                estado: String(grid[i][c - 1] || '').trim()
            };
        }
        g.dias[ymd] = dia;
    }
    return g;
}

function calSync_libreParaBot_(estado) {
    const st = String(estado || '').trim().toLowerCase();
    if (CAL_SYNC.ESTADOS_LIBRES.includes(st)) return true;
    return st === 'nuevos' && CAL_SYNC.NUEVOS_PARA_BOT;
}

/** Conjunto "dni|yyyy-MM-dd|HH:mm" de las reservas con fecha en la hoja Reservas. */
function calSync_reservasSet_(tz) {
    const sh = _sh_(SHEETS.RESERVAS);
    const last = sh.getLastRow();
    const set = new Set();
    if (last < 2) return set;
    const vals = sh.getRange(2, 1, last - 1, 4).getValues();
    const disp = sh.getRange(2, 3, last - 1, 1).getDisplayValues();
    vals.forEach((row, i) => {
        const f = row[1];
        if (!(f instanceof Date) || isNaN(f.getTime())) return;
        set.add(_normDni_(row[3]) + '|' + Utilities.formatDate(f, tz, 'yyyy-MM-dd') + '|' + _normHoraStr_(disp[i][0]));
    });
    return set;
}

/***** Hoja SyncCalendar (registro de reservas del bot) *****/
function calSync_hojaMapa_() {
    const ss = _ss_();
    let sh = ss.getSheetByName(CAL_SYNC.HOJA_SYNC);
    if (!sh) {
        sh = ss.insertSheet(CAL_SYNC.HOJA_SYNC);
        sh.appendRow(CAL_SYNC_COLS);
        sh.setFrozenRows(1);
    }
    return sh;
}

function calSync_leerMapa_() {
    const sh = calSync_hojaMapa_();
    const last = sh.getLastRow();
    if (last < 2) return [];
    return sh.getRange(2, 1, last - 1, CAL_SYNC_COLS.length).getDisplayValues().map((row, i) => ({
        fila: i + 2,
        eventId: String(row[0] || ''),
        estado: String(row[1] || '').trim(),
        dni: _normDni_(row[2]),
        nombre: String(row[3] || ''),
        ymd: String(row[4] || '').trim(),
        hora: _normHoraStr_(row[5]),
        inicioMs: Number(String(row[6] || '').replace(/\D+/g, '')) || 0
    }));
}

function calSync_setEstado_(m, estado, detalle) {
    const sh = calSync_hojaMapa_();
    sh.getRange(m.fila, 2).setValue(estado);
    sh.getRange(m.fila, 8).setValue(new Date());
    if (detalle) sh.getRange(m.fila, 9).setValue(detalle);
}

/***** Helpers varios *****/
function calSync_getCalendar_() {
    if (CAL_SYNC.CALENDAR_ID) {
        const c = CalendarApp.getCalendarById(CAL_SYNC.CALENDAR_ID);
        if (!c) throw new Error('No se encontró el calendario ' + CAL_SYNC.CALENDAR_ID);
        return c;
    }
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('CAL_SYNC_CALENDAR_ID');
    if (id) {
        const c = CalendarApp.getCalendarById(id);
        if (c) return c;
    }
    const existentes = CalendarApp.getCalendarsByName(CAL_SYNC.CALENDAR_NAME);
    const cal = existentes.length ? existentes[0] : CalendarApp.createCalendar(CAL_SYNC.CALENDAR_NAME, { timeZone: calSync_tz_() });
    props.setProperty('CAL_SYNC_CALENDAR_ID', cal.getId());
    return cal;
}

// Misma zona horaria que usa el turnero (la del proyecto de Apps Script).
function calSync_tz_() {
    return _tz_();
}

function calSync_tag_(ev) {
    try { return ev.getTag(CAL_SYNC_TAG) || ''; } catch (_) { return ''; }
}

function calSync_avisar_(asunto, cuerpo) {
    if (!CAL_SYNC.EMAIL_AVISOS) return;
    try { MailApp.sendEmail(CAL_SYNC.EMAIL_AVISOS, '[Turnero] ' + asunto, cuerpo); } catch (e) { console.error('Mail: ' + e); }
}

function calSync_borrarActivadores_() {
    ScriptApp.getProjectTriggers()
        .filter(t => t.getHandlerFunction() === 'syncTurnosCalendar')
        .forEach(t => ScriptApp.deleteTrigger(t));
}
