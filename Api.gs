/**
 * PastoreoGestion — Api.gs v1.1.0
 * Agrega doPost con guardar_potrero, guardar_rodeo,
 * registrar_ingreso, registrar_egreso.
 */

function doGet(e) {
  try {
    const params = e.parameter || {};
    const tokenError = _validarToken(params.token);
    if (tokenError) return _respuestaError(tokenError, 401);
    const action = (params.action || "").toLowerCase().trim();
    switch (action) {
      case "calcular":          return _handleCalcular(params);
      case "potreros":          return _handlePotreros(params);
      case "rodeos":            return _handleRodeos(params);
      case "estado":            return _handleEstado();
      case "guardar_potrero":   return _handleGuardarPotrero(params);
      case "guardar_rodeo":     return _handleGuardarRodeo(params);
      case "registrar_ingreso": return _handleRegistrarIngreso(params);
      case "registrar_egreso":  return _handleRegistrarEgreso(params);
      default: return _respuestaError(`Acción desconocida: "${action}"`, 400);
    }
  } catch (err) {
    Logger.log("doGet ERROR: " + err.message);
    return _respuestaError("Error interno: " + err.message, 500);
  }
}

function doPost(e) {
  try {
    let body = {};
    try { body = JSON.parse(e.postData.contents); } catch(x) {}
    const tokenError = _validarToken(body.token);
    if (tokenError) return _respuestaError(tokenError, 401);
    const action = (body.action || "").toLowerCase().trim();
    switch (action) {
      case "guardar_potrero":   return _handleGuardarPotrero(body);
      case "guardar_rodeo":     return _handleGuardarRodeo(body);
      case "registrar_ingreso": return _handleRegistrarIngreso(body);
      case "registrar_egreso":  return _handleRegistrarEgreso(body);
      default: return _respuestaError(`Acción POST desconocida: "${action}"`, 400);
    }
  } catch (err) {
    Logger.log("doPost ERROR: " + err.message);
    return _respuestaError("Error interno: " + err.message, 500);
  }
}

function _handleCalcular(params) {
  const potreroId = (params.potrero_id || "").trim();
  const rodeoId   = (params.rodeo_id   || "").trim();
  const dias      = params.dias ? parseInt(params.dias) : null;
  if (!potreroId) return _respuestaError("Falta potrero_id.", 400);
  if (!rodeoId)   return _respuestaError("Falta rodeo_id.", 400);
  const resultado = calcularPastoreo(potreroId, rodeoId, dias);
  if (!resultado.ok) return _respuestaError(resultado.error, 422);
  return _respuestaOk(resultado);
}

function _handlePotreros(params) {
  const estId = (params.est_id || "").trim();
  const soloDisponibles = params.solo_disponibles === "true";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Potreros");
  if (!sheet) return _respuestaError("No se encontró Potreros.", 500);
  const data = sheet.getDataRange().getValues();
  const potreros = [];
  for (let i = 3; i < data.length; i++) {
    const f = data[i];
    if (!f[0]) continue;
    if (estId && f[1] !== estId) continue;
    if (soloDisponibles && f[5] !== "disponible") continue;
    potreros.push({
      potrero_id: f[0], est_id: f[1], nombre: f[2], ha: f[3],
      tipo_recurso: f[4], estado: f[5], fecha_disponible: _formatFecha(f[6]),
      kg_ms_ha: f[7], aprovec_pct: f[8], periodo_descanso_dias: f[9],
      altura_entrada_cm: f[10], altura_salida_cm: f[11], notas: f[12] || "",
    });
  }
  return _respuestaOk({ potreros, total: potreros.length });
}

function _handleRodeos(params) {
  const estId = (params.est_id || "").trim();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rodeos");
  if (!sheet) return _respuestaError("No se encontró Rodeos.", 500);
  const data = sheet.getDataRange().getValues();
  const rodeos = [];
  for (let i = 3; i < data.length; i++) {
    const f = data[i];
    if (!f[0]) continue;
    if (estId && f[1] !== estId) continue;
    rodeos.push({
      rodeo_id: f[0], est_id: f[1], nombre: f[2], categoria: f[3],
      cabezas: f[4], peso_prom_kg: f[5], consumo_pct_pv: f[6],
      demanda_ms_dia_kg: Math.round(f[4] * f[5] * (f[6]/100) * 10) / 10,
      potrero_actual: f[8] || "", notas: f[9] || "",
    });
  }
  return _respuestaOk({ rodeos, total: rodeos.length });
}

function _handleEstado() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = _leerConfig();
  const sheetP = ss.getSheetByName("Potreros");
  const sheetR = ss.getSheetByName("Rodeos");
  const dataP  = sheetP ? sheetP.getDataRange().getValues().slice(3) : [];
  const dataR  = sheetR ? sheetR.getDataRange().getValues().slice(3) : [];
  const conteo = { disponible: 0, ocupado: 0, descanso: 0 };
  dataP.forEach(f => { if (f[0] && conteo.hasOwnProperty(f[5])) conteo[f[5]]++; });
  return _respuestaOk({
    app: "PastoreoGestion", version: config["version_plantilla"] || "1.0.0",
    est_id_activo: config["est_id_activo"] || "",
    nombre_est: _getNombreEstablecimiento(config["est_id_activo"]),
    potreros: conteo, total_rodeos: dataR.filter(f => f[0]).length,
    timestamp: new Date().toISOString(),
  });
}

function _handleGuardarPotrero(body) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Potreros");
  if (!sheet) return _respuestaError("No se encontró Potreros.", 500);
  if (!body.nombre)       return _respuestaError("nombre requerido.", 400);
  if (!body.ha)           return _respuestaError("ha requerido.", 400);
  if (!body.tipo_recurso) return _respuestaError("tipo_recurso requerido.", 400);
  if (!body.kg_ms_ha)     return _respuestaError("kg_ms_ha requerido.", 400);

  const config = _leerConfig();
  const estId  = body.est_id || config["est_id_activo"] || "EST-001";
  const data   = sheet.getDataRange().getValues();

  if (body.potrero_id) {
    for (let i = 3; i < data.length; i++) {
      if (data[i][0] === body.potrero_id) {
        const r = i + 1;
        sheet.getRange(r, 3).setValue(body.nombre);
        sheet.getRange(r, 4).setValue(parseFloat(body.ha));
        sheet.getRange(r, 5).setValue(body.tipo_recurso);
        sheet.getRange(r, 8).setValue(parseFloat(body.kg_ms_ha));
        sheet.getRange(r, 9).setValue(parseFloat(body.aprovec_pct || 70));
        sheet.getRange(r, 10).setValue(parseInt(body.periodo_descanso_dias || 45));
        sheet.getRange(r, 11).setValue(parseInt(body.altura_entrada_cm || 20));
        sheet.getRange(r, 12).setValue(parseInt(body.altura_salida_cm  || 6));
        sheet.getRange(r, 13).setValue(body.notas || "");
        SpreadsheetApp.flush();
        return _respuestaOk({ potrero_id: body.potrero_id, accion: "editado" });
      }
    }
    return _respuestaError(`Potrero "${body.potrero_id}" no encontrado.`, 404);
  } else {
    const nuevoId = _generarId("POT", "ultimo_potrero_id", 3);
    const f = _ultimaFilaConDatos(sheet) + 1;
    sheet.getRange(f, 1, 1, 13).setValues([[
      nuevoId, estId, body.nombre, parseFloat(body.ha), body.tipo_recurso,
      "disponible", "", parseFloat(body.kg_ms_ha), parseFloat(body.aprovec_pct || 70),
      parseInt(body.periodo_descanso_dias || 45), parseInt(body.altura_entrada_cm || 20),
      parseInt(body.altura_salida_cm || 6), body.notas || "",
    ]]);
    SpreadsheetApp.flush();
    return _respuestaOk({ potrero_id: nuevoId, accion: "creado" });
  }
}

function _handleGuardarRodeo(body) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rodeos");
  if (!sheet) return _respuestaError("No se encontró Rodeos.", 500);
  if (!body.nombre)       return _respuestaError("nombre requerido.", 400);
  if (!body.categoria)    return _respuestaError("categoria requerido.", 400);
  if (!body.cabezas)      return _respuestaError("cabezas requerido.", 400);
  if (!body.peso_prom_kg) return _respuestaError("peso_prom_kg requerido.", 400);

  const config = _leerConfig();
  const estId  = body.est_id || config["est_id_activo"] || "EST-001";
  const data   = sheet.getDataRange().getValues();

  if (body.rodeo_id) {
    for (let i = 3; i < data.length; i++) {
      if (data[i][0] === body.rodeo_id) {
        const r = i + 1;
        sheet.getRange(r, 3).setValue(body.nombre);
        sheet.getRange(r, 4).setValue(body.categoria);
        sheet.getRange(r, 5).setValue(parseInt(body.cabezas));
        sheet.getRange(r, 6).setValue(parseFloat(body.peso_prom_kg));
        sheet.getRange(r, 7).setValue(parseFloat(body.consumo_pct_pv || 2.5));
        sheet.getRange(r, 10).setValue(body.notas || "");
        SpreadsheetApp.flush();
        return _respuestaOk({ rodeo_id: body.rodeo_id, accion: "editado" });
      }
    }
    return _respuestaError(`Rodeo "${body.rodeo_id}" no encontrado.`, 404);
  } else {
    const nuevoId = _generarId("ROD", "ultimo_rodeo_id", 3);
    const f = _ultimaFilaConDatos(sheet) + 1;
    sheet.getRange(f, 1, 1, 10).setValues([[
      nuevoId, estId, body.nombre, body.categoria,
      parseInt(body.cabezas), parseFloat(body.peso_prom_kg),
      parseFloat(body.consumo_pct_pv || 2.5), "", "", body.notas || "",
    ]]);
    SpreadsheetApp.flush();
    return _respuestaOk({ rodeo_id: nuevoId, accion: "creado" });
  }
}

function _handleRegistrarIngreso(body) {
  if (!body.potrero_id) return _respuestaError("Falta potrero_id.", 400);
  if (!body.rodeo_id)   return _respuestaError("Falta rodeo_id.", 400);
  try {
    registrarIngreso(body.potrero_id, body.rodeo_id, new Date());
    return _respuestaOk({ mensaje: `Ingreso registrado. ${body.potrero_id} → ocupado.` });
  } catch(e) { return _respuestaError(e.message, 422); }
}

function _handleRegistrarEgreso(body) {
  if (!body.potrero_id) return _respuestaError("Falta potrero_id.", 400);
  try {
    registrarEgreso(body.potrero_id, new Date());
    return _respuestaOk({ mensaje: `Egreso registrado. ${body.potrero_id} → descanso.` });
  } catch(e) { return _respuestaError(e.message, 422); }
}

function _validarToken(tokenRecibido) {
  if (!tokenRecibido) return "Token requerido.";
  const config = _leerConfig();
  const esperado = config["api_token"];
  if (!esperado) return "api_token no configurado.";
  if (tokenRecibido !== esperado) return "Token inválido.";
  return null;
}

function _respuestaOk(data) {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function _respuestaError(mensaje, statusCode) {
  return ContentService.createTextOutput(JSON.stringify({ ok: false, status: statusCode || 400, error: mensaje }))
    .setMimeType(ContentService.MimeType.JSON);
}

let _apiConfigCache = null;
function _leerConfig() {
  if (_apiConfigCache) return _apiConfigCache;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 3; i < data.length; i++) { if (data[i][0]) config[data[i][0]] = data[i][1]; }
  _apiConfigCache = config;
  return config;
}

function _getNombreEstablecimiento(estId) {
  if (!estId) return "";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Establecimiento");
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  for (let i = 3; i < data.length; i++) { if (data[i][0] === estId) return data[i][2]; }
  return "";
}

function _formatFecha(valor) {
  if (!valor) return "";
  if (valor instanceof Date)
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return String(valor);
}

function _generarId(prefijo, configKey, digits) {
  const config  = _leerConfig();
  const ultimo  = config[configKey] || `${prefijo}-${"0".repeat(digits)}`;
  const match   = ultimo.match(/(\d+)$/);
  const numero  = match ? parseInt(match[1]) + 1 : 1;
  const nuevoId = `${prefijo}-${String(numero).padStart(digits, "0")}`;
  _guardarEnConfig(configKey, nuevoId);
  _apiConfigCache = null;
  return nuevoId;
}

function _guardarEnConfig(clave, valor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 3; i < data.length; i++) {
    if (data[i][0] === clave) { sheet.getRange(i + 1, 2).setValue(valor); return; }
  }
  const next = _ultimaFilaConDatos(sheet) + 1;
  sheet.getRange(next, 1).setValue(clave);
  sheet.getRange(next, 2).setValue(valor);
}

function _ultimaFilaConDatos(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i].some(c => c !== "")) return i + 1;
  }
  return 1;
}
