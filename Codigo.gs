/**
 * PastoreoGestion — Codigo.gs
 * Entry point del proyecto Apps Script.
 *
 * Responsabilidades:
 *   - onOpen: agrega el menú PastoreoGestion en Sheets
 *   - doGet: delega todos los requests HTTP a Api.gs
 *   - Funciones de utilidad accesibles desde el menú
 *
 * @version 1.0.0
 */

// ─── MENÚ EN GOOGLE SHEETS ─────────────────────────────────────────────────

/**
 * Se ejecuta automáticamente cuando el usuario abre la planilla.
 * Agrega el menú "PastoreoGestion" en la barra de Sheets.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🌿 PastoreoGestion")
    .addItem("Calcular pastoreo actual", "menuCalcularActual")
    .addSeparator()
    .addItem("Registrar ingreso de rodeo", "menuRegistrarIngreso")
    .addItem("Registrar egreso de rodeo",  "menuRegistrarEgreso")
    .addSeparator()
    .addItem("Ver estado de potreros",     "menuVerEstado")
    .addSeparator()
    .addItem("Generar nuevo token API",    "menuGenerarToken")
    .addItem("Mostrar URL de la API",      "menuMostrarUrl")
    .addToUi();
}

// ─── ACCIONES DEL MENÚ ─────────────────────────────────────────────────────

/**
 * Lee potrero_id y rodeo_id de la pestaña Calculadora y ejecuta el cálculo.
 * Escribe los resultados de vuelta en la hoja para que el productor los vea.
 */
function menuCalcularActual() {
  const ui = SpreadsheetApp.getUi();

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calculadora");
    if (!sheet) throw new Error("No se encontró la pestaña Calculadora.");

    const potreroId = sheet.getRange("B4").getValue();
    const rodeoId   = sheet.getRange("B5").getValue();
    const diasObj   = sheet.getRange("B6").getValue();

    if (!potreroId || !rodeoId) {
      ui.alert("Completá potrero_id (B4) y rodeo_id (B5) en la pestaña Calculadora.");
      return;
    }

    const resultado = calcularPastoreo(potreroId, rodeoId, diasObj || null);

    if (!resultado.ok) {
      ui.alert("Error en el cálculo:\n\n" + resultado.error);
      return;
    }

    // Las fórmulas INDEX/MATCH de la hoja ya muestran los resultados.
    // Mostramos un resumen en un popup para confirmación rápida.
    const c = resultado.calculo;
    const f = resultado.fechas;
    const alertas = resultado.alertas.map(a => `[${a.nivel.toUpperCase()}] ${a.mensaje}`).join("\n");

    const resumen =
      `Potrero: ${resultado.potrero.nombre} (${resultado.potrero.ha} ha)\n` +
      `Rodeo:   ${resultado.rodeo.nombre} (${resultado.rodeo.cabezas} cabezas)\n\n` +
      `Oferta:  ${c.ofertaTotalKgMs.toLocaleString()} kg MS\n` +
      `Demanda: ${c.demandaDiariaKgMs} kg MS/día\n\n` +
      `Días posibles:  ${c.diasPastoreoPosibles}\n` +
      `Días finales:   ${c.diasPastoreoFinal}\n` +
      `Carga máx:      ${c.cargaMaximaCabezas} cabezas\n\n` +
      `Entrada:  ${f.entrada}\n` +
      `Salida:   ${f.salidaEst}\n` +
      `Retorno:  ${f.retornoPrev}\n\n` +
      `─── Alertas ───\n${alertas}`;

    ui.alert("Resultado del cálculo", resumen, ui.ButtonSet.OK);

  } catch (err) {
    ui.alert("Error inesperado: " + err.message);
  }
}

/**
 * Diálogo para registrar el ingreso de un rodeo a un potrero.
 * Actualiza el estado del potrero y escribe en la pestaña Pastoreos.
 */
function menuRegistrarIngreso() {
  const ui = SpreadsheetApp.getUi();

  const resP = ui.prompt(
    "Registrar ingreso",
    "Ingresá el ID del potrero (ej: POT-001):",
    ui.ButtonSet.OK_CANCEL
  );
  if (resP.getSelectedButton() !== ui.Button.OK) return;

  const resR = ui.prompt(
    "Registrar ingreso",
    "Ingresá el ID del rodeo (ej: ROD-001):",
    ui.ButtonSet.OK_CANCEL
  );
  if (resR.getSelectedButton() !== ui.Button.OK) return;

  const potreroId = resP.getResponseText().trim();
  const rodeoId   = resR.getResponseText().trim();

  try {
    registrarIngreso(potreroId, rodeoId, new Date());
    ui.alert(`Ingreso registrado correctamente.\nPotrero ${potreroId} → estado: ocupado.`);
  } catch (err) {
    ui.alert("Error al registrar ingreso:\n" + err.message);
  }
}

/**
 * Diálogo para registrar el egreso de un rodeo y calcular la fecha de retorno.
 */
function menuRegistrarEgreso() {
  const ui = SpreadsheetApp.getUi();

  const res = ui.prompt(
    "Registrar egreso",
    "Ingresá el ID del potrero del que salen (ej: POT-001):",
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const potreroId = res.getResponseText().trim();

  try {
    registrarEgreso(potreroId, new Date());
    ui.alert(`Egreso registrado correctamente.\nPotrero ${potreroId} → estado: descanso.`);
  } catch (err) {
    ui.alert("Error al registrar egreso:\n" + err.message);
  }
}

/**
 * Muestra un resumen del estado actual de todos los potreros del establecimiento activo.
 */
function menuVerEstado() {
  const ui = SpreadsheetApp.getUi();

  try {
    const config = _leerConfig();
    const estId  = config["est_id_activo"];
    const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Potreros");
    const data   = sheet.getDataRange().getValues().slice(3);

    const lineas = [`Estado de potreros — ${estId}\n`];
    let disponibles = 0, ocupados = 0, descanso = 0;

    data.forEach(f => {
      if (!f[0] || f[1] !== estId) return;
      const estado = f[5];
      const fecha  = f[6] ? ` (disponible: ${_formatFecha(f[6])})` : "";
      lineas.push(`${f[0]}  ${f[2].padEnd(20)} [${estado}]${fecha}`);
      if (estado === "disponible") disponibles++;
      else if (estado === "ocupado") ocupados++;
      else descanso++;
    });

    lineas.push(`\nResumen: ${disponibles} disponibles · ${ocupados} ocupados · ${descanso} en descanso`);

    ui.alert("Estado de potreros", lineas.join("\n"), ui.ButtonSet.OK);

  } catch (err) {
    ui.alert("Error: " + err.message);
  }
}

/**
 * Genera un nuevo token aleatorio y lo guarda en Config.
 * El productor debe copiar este token al archivo api.js del frontend.
 */
function menuGenerarToken() {
  const ui = SpreadsheetApp.getUi();

  const confirmar = ui.alert(
    "Generar nuevo token API",
    "¿Estás seguro? El token anterior dejará de funcionar y tendrás que actualizar el frontend.",
    ui.ButtonSet.YES_NO
  );
  if (confirmar !== ui.Button.YES) return;

  const nuevoToken = _generarToken();
  _guardarEnConfig("api_token", nuevoToken);

  ui.alert(
    "Token generado",
    `Nuevo token:\n\n${nuevoToken}\n\nGuardalo en el archivo api.js de tu frontend como la constante API_TOKEN.`,
    ui.ButtonSet.OK
  );
}

/**
 * Muestra la URL del Web App desplegado para copiar al frontend.
 */
function menuMostrarUrl() {
  const ui  = SpreadsheetApp.getUi();
  const url = ScriptApp.getService().getUrl();

  if (!url) {
    ui.alert(
      "La app no está desplegada aún.\n\n" +
      "Hacé clic en Implementar → Nueva implementación → Aplicación web.\n" +
      "Ejecutar como: Yo. Quién puede acceder: Cualquier usuario."
    );
    return;
  }

  ui.alert("URL de la API", url + "\n\nAgregar ?token=TU_TOKEN&action=estado para probarla.", ui.ButtonSet.OK);
}

// ─── REGISTRO DE INGRESOS Y EGRESOS ────────────────────────────────────────

/**
 * Registra el ingreso de un rodeo a un potrero.
 * - Actualiza estado del potrero a "ocupado"
 * - Actualiza potrero_actual en el rodeo
 * - Crea una nueva fila en Pastoreos
 *
 * @param {string} potreroId
 * @param {string} rodeoId
 * @param {Date}   fechaEntrada
 */
function registrarIngreso(potreroId, rodeoId, fechaEntrada) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Actualizar estado del potrero
  _actualizarEstadoPotrero(potreroId, "ocupado", null);

  // 2. Actualizar potrero_actual en el rodeo
  _actualizarPotreroActualRodeo(rodeoId, potreroId);

  // 3. Generar ID para el pastoreo y escribir en Pastoreos
  const pastoreoId = _generarId("PAS", "ultimo_pastoreo_id", 5);
  const config     = _leerConfig();

  const sheetP = ss.getSheetByName("Pastoreos");
  if (!sheetP) throw new Error("No se encontró la pestaña Pastoreos.");

  const ultimaFila = _ultimaFilaConDatos(sheetP) + 1;
  const fila = [
    pastoreoId,
    config["est_id_activo"] || "",
    potreroId,
    rodeoId,
    Utilities.formatDate(fechaEntrada, Session.getScriptTimeZone(), "yyyy-MM-dd"),
    "",       // fecha_salida — se completa al egresar
    "",       // dias_ocupacion — calculado al egresar
    "",       // cabezas — se copia desde Rodeos automáticamente
    "",       // kg_ms_ofrecido — se calcula al egresar
    "",       // kg_ms_consumido_est
    "",       // observaciones
  ];

  sheetP.getRange(ultimaFila, 1, 1, fila.length).setValues([fila]);
  Logger.log(`Ingreso registrado: ${pastoreoId} | ${potreroId} | ${rodeoId}`);
}

/**
 * Registra el egreso de un rodeo del potrero actualmente ocupado.
 * - Actualiza estado del potrero a "descanso"
 * - Calcula y guarda la fecha de retorno (estado → disponible)
 * - Completa la fila de Pastoreos con fecha_salida y días
 *
 * @param {string} potreroId
 * @param {Date}   fechaSalida
 */
function registrarEgreso(potreroId, fechaSalida) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Leer período de descanso del potrero
  const potrero = _getPotrero(potreroId);

  // 2. Calcular fecha de retorno
  const fechaRetorno = new Date(fechaSalida);
  fechaRetorno.setDate(fechaRetorno.getDate() + potrero.periodo_descanso_dias);

  // 3. Actualizar potrero: estado = descanso, fecha_disponible = fechaRetorno
  _actualizarEstadoPotrero(potreroId, "descanso", fechaRetorno);

  // 4. Limpiar potrero_actual del rodeo
  _limpiarPotreroActualRodeo(potreroId);

  // 5. Completar la última fila de Pastoreos para este potrero
  _completarRegistroPastoreo(potreroId, fechaSalida, potrero);

  Logger.log(`Egreso registrado: ${potreroId} | salida: ${fechaSalida} | retorno: ${fechaRetorno}`);
}

// ─── HELPERS DE ESCRITURA EN SHEETS ────────────────────────────────────────

function _actualizarEstadoPotrero(potreroId, nuevoEstado, fechaDisponible) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Potreros");
  const data  = sheet.getDataRange().getValues();

  for (let i = 3; i < data.length; i++) {
    if (data[i][0] === potreroId) {
      // Col F (índice 5) = estado, Col G (índice 6) = fecha_disponible
      sheet.getRange(i + 1, 6).setValue(nuevoEstado);
      if (fechaDisponible) {
        sheet.getRange(i + 1, 7).setValue(
          Utilities.formatDate(fechaDisponible, Session.getScriptTimeZone(), "yyyy-MM-dd")
        );
      }
      return;
    }
  }
  throw new Error(`Potrero "${potreroId}" no encontrado.`);
}

function _actualizarPotreroActualRodeo(rodeoId, potreroId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rodeos");
  const data  = sheet.getDataRange().getValues();

  for (let i = 3; i < data.length; i++) {
    if (data[i][0] === rodeoId) {
      sheet.getRange(i + 1, 9).setValue(potreroId); // Col I = potrero_actual
      return;
    }
  }
  throw new Error(`Rodeo "${rodeoId}" no encontrado.`);
}

function _limpiarPotreroActualRodeo(potreroId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rodeos");
  const data  = sheet.getDataRange().getValues();

  for (let i = 3; i < data.length; i++) {
    if (data[i][8] === potreroId) { // Col I = potrero_actual
      sheet.getRange(i + 1, 9).setValue("");
    }
  }
}

function _completarRegistroPastoreo(potreroId, fechaSalida, potrero) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pastoreos");
  const data  = sheet.getDataRange().getValues();
  const fmtFecha = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");

  // Buscar la última fila abierta (sin fecha_salida) para este potrero
  for (let i = data.length - 1; i >= 3; i--) {
    if (data[i][2] === potreroId && !data[i][5]) {
      const fechaEntrada = new Date(data[i][4]);
      const dias = Math.round((fechaSalida - fechaEntrada) / (1000 * 60 * 60 * 24));

      sheet.getRange(i + 1, 6).setValue(fmtFecha(fechaSalida));   // fecha_salida
      sheet.getRange(i + 1, 7).setValue(dias);                     // dias_ocupacion
      return;
    }
  }
  Logger.log(`Advertencia: no se encontró registro abierto de pastoreo para ${potreroId}`);
}

// ─── GENERACIÓN DE IDs ──────────────────────────────────────────────────────

/**
 * Genera el próximo ID incremental del tipo indicado.
 * Lee el último ID de Config, lo incrementa, guarda el nuevo valor.
 *
 * @param {string} prefijo    - "POT", "ROD", "PAS"
 * @param {string} configKey  - clave en Config que guarda el último ID
 * @param {number} digits     - dígitos del número (POT→3, PAS→5)
 */
function _generarId(prefijo, configKey, digits) {
  const config    = _leerConfig();
  const ultimoId  = config[configKey] || `${prefijo}-${"0".repeat(digits)}`;

  // Extraer el número del último ID y sumar 1
  const match     = ultimoId.match(/(\d+)$/);
  const numero    = match ? parseInt(match[1]) + 1 : 1;
  const nuevoId   = `${prefijo}-${String(numero).padStart(digits, "0")}`;

  _guardarEnConfig(configKey, nuevoId);
  _configCache = null; // invalidar caché

  return nuevoId;
}

/**
 * Actualiza una clave en la pestaña Config.
 */
function _guardarEnConfig(clave, valor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 3; i < data.length; i++) {
    if (data[i][0] === clave) {
      sheet.getRange(i + 1, 2).setValue(valor);
      return;
    }
  }
  // Si no existe la clave, agregarla al final
  const nextRow = _ultimaFilaConDatos(sheet) + 1;
  sheet.getRange(nextRow, 1).setValue(clave);
  sheet.getRange(nextRow, 2).setValue(valor);
}

/**
 * Genera un token aleatorio de 32 caracteres alfanuméricos.
 */
function _generarToken() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  let token = "";
  for (let i = 0; i < 32; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

/**
 * Devuelve el índice de la última fila con datos en una hoja.
 */
function _ultimaFilaConDatos(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i].some(cell => cell !== "")) return i + 1;
  }
  return 1;
}
