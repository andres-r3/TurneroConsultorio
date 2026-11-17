# Turnero Consultorio

| Carpeta | Qué es |
|---|---|
| `turnero/` | Turnero público (web app de Apps Script): `Code.gs`, `index.html` y **`CalendarSync.gs`** (nuevo). |
| `superusuario/` | Panel de secretaría (otro proyecto de Apps Script): `codigo.gs`, `index.html`. |

Las dos apps trabajan sobre la planilla **Consultorio** (hojas `Legajos`, `Disponibilidad`, `Reservas`, `Dialogos`, …).

---

## Sincronización con Google Calendar (bot de WhatsApp)

La planilla **sigue siendo la fuente de verdad**. `CalendarSync.gs` mantiene un calendario dedicado
(**"Turnos Consultorio (bot)"**) para que el bot de WhatsApp vea como libres **exactamente** los turnos
libres del turnero y pueda reservar.

### Cómo funciona

1. **Planilla → Calendar.** Todo lo que *no* es un turno libre se cubre con eventos **"No disponible"**:
   horarios reservados o bloqueados, días sin atención, fechas no publicadas y el turnero cerrado (`PUBLIC_ON`).
   Los huecos que quedan son los turnos libres.
2. **Calendar → Planilla.** Cuando el bot crea un evento en un hueco, el script:
   - busca al paciente por **DNI** escrito en el título o la descripción del evento
     (si no hay DNI, prueba con el **teléfono** de `Legajos`);
   - aplica **las mismas reglas que el turnero**: acceso, celda vacía / `nuevos`, tope mensual y tope de reservas activas;
   - si todo está bien, escribe `Reservado` en `Disponibilidad`, agrega la fila en `Reservas`
     (reutiliza el saldo a favor si lo hay) y la registra en la hoja **`SyncCalendar`**;
   - si algo falla, **rechaza** el evento: lo borra, o lo marca `[RECHAZADO]` si así se configura.
     El rechazo queda en `SyncCalendar`, en `Dialogos` y, si se configura, se avisa por mail.
3. **Cancelaciones y cambios.**
   - Si el bot **borra** el evento, se cancela el turno en la planilla. Si la cancelación está
     deshabilitada (`CANCEL_ON`), el turno se mantiene y queda registrado.
   - Si el bot **mueve** el evento, se cancela el turno viejo y se procesa el nuevo horario.
   - Si el turno se cancela o reasigna desde el **turnero o la secretaría**, se borra el evento del bot.

Se ejecuta en estos momentos:
- cada 5 minutos;
- cada vez que cambia el calendario;
- **al instante** cuando un paciente reserva o cancela desde el turnero web.

Los cambios hechos desde el panel de superusuario se reflejan en Calendar en un máximo de unos 5 minutos.

### Instalación (una sola vez)

1. En el proyecto de Apps Script del **turnero público**:
   - crear el archivo `CalendarSync.gs` con el contenido de `turnero/CalendarSync.gs`;
   - reemplazar `Code.gs` por la versión del repo (solo agrega 2 líneas que llaman a la sincronización).
2. Verificar en *Configuración del proyecto* que la zona horaria sea `America/Argentina/Cordoba`,
   la misma que la de la planilla.
3. En el editor, elegir la función **`calSync_instalar`** y presionar *Ejecutar*. Esto:
   - pide autorizar el acceso a Google Calendar;
   - crea el calendario;
   - crea los activadores;
   - hace la primera sincronización.
4. **Implementar → Gestionar implementaciones → Editar → Nueva versión**, para que la web app publicada
   incluya los cambios.
5. Abrir Google Calendar y revisar que el calendario "Turnos Consultorio (bot)" tenga los bloques "No disponible".

Para desactivar: ejecutar `calSync_desinstalar`.
Para limpiar los bloques: ejecutar `calSync_borrarBloques`.

### Configuración del bot de WhatsApp (importante)

- Conectarlo **solo** al calendario **"Turnos Consultorio (bot)"**, no al calendario personal.
- **Duración de la cita = duración del turno de la grilla.** Por ejemplo 15 min: es la diferencia entre
  las horas de la columna A de `Disponibilidad`. Si el evento dura más, se rechaza.
- El horario de atención del bot puede ser amplio (por ejemplo 8 a 20 h): la disponibilidad real la marcan los bloques.
- Hacer que el bot **pida el DNI** y lo escriba en el título o la descripción del evento, por ejemplo `DNI 30123456`.
- Conviene que el mensaje de confirmación diga *"tu turno queda sujeto a verificación del consultorio"*:
  la validación ocurre después de que el bot crea el evento.
- Eventos propios en ese calendario: empezar el título con `#` (por ejemplo `#reunión`) y el script no los toca.

### Opciones (`CAL_SYNC` al inicio de `CalendarSync.gs`)

| Opción | Por defecto | Para qué |
|---|---|---|
| `DIAS_HORIZONTE` | 60 | Días hacia adelante que se sincronizan. |
| `SOLO_FECHAS_PUBLICAS` | `true` | El bot solo ofrece las fechas publicadas (B1/C1), igual que el turnero. |
| `NUEVOS_PARA_BOT` | `false` | Si es `true`, las celdas `nuevos` aparecen libres en el bot (solo podrá tomarlas un paciente "nuevo"). |
| `ACCION_RECHAZO` | `'eliminar'` | `'marcar'` deja el evento con título `[RECHAZADO]` en lugar de borrarlo. |
| `EMAIL_AVISOS` | vacío | Mail que recibe un aviso por cada rechazo o cancelación denegada. |
| `CALENDAR_ID` | vacío | Para usar un calendario existente en lugar de crear uno. |

### Limitaciones a tener en cuenta

- **Las validaciones son posteriores a la reserva.** Si un paciente no registrado, o que ya llegó a su tope,
  reserva por WhatsApp, el bot le confirma y segundos o minutos después el turno se rechaza.
  Hay que contactarlo: los datos quedan en `SyncCalendar` y en el mail de aviso.
- **Pago por transferencia.** El bot no puede exigirlo (turnos de las 10:00 o antes y pacientes nuevos).
  Queda anotado como *"requiere transferencia"* en `SyncCalendar`.
- **Choque de reservas simultáneas.** Si alguien toma el mismo horario por la web en el mismo momento,
  gana la primera reserva registrada en la planilla y la del bot se rechaza.
