/**
 * PastoreoGestion — Calculadora.gs
 * Módulo de cálculo agronómico central.
 *
 * Responsabilidades:
 *   - calcularPastoreo(potrero_id, rodeo_id, diasObjetivo?)
 *   - calcularCargaMaxima(potrero_id, rodeo_id, diasObjetivo)
 *   - obtenerTasaCrecimiento(tipo_recurso, zona, mes)
 *   - escribirResultadosEnCalculadora(resultado)
 *
 * Todas las funciones son puras (reciben IDs, leen Sheets, devuelven objetos).
 * No escriben a Sheets salvo escribirResultadosEnCalculadora().
 *
 * @version 1.0.0
 */

// ─── CONSTANTES AGRONÓMICAS ────────────────────────────────────────────────

const CONSUMO_POR_CATEGORIA = {
  "Vaca de cría":   { min: 2.0, tipico: 2.5 },
  "Vaquillona":     { min: 2.2, tipico: 2.8 },
  "Novillo":        { min: 2.0, tipico: 2.5 },
  "Ternero":        { min: 3.0, tipico: 3.5 },
  "Toro":           { min: 1.8, tipico: 2.0 },
};

const APROVECHAMIENTO_POR_SISTEMA = {
  "rotativo_intensivo": 75,
  "rotativo_extensivo": 60,
  "continuo":           45,
};

// ─── FUNCIÓN PRINCIPAL ─────────────────────────────────────────────────────

/**
 * Calcula todos los parámetros de un evento de pastoreo.
 *
 * @param {string} potreroId   - ID del potrero (ej: "POT-001")
 * @param {string} rodeoId     - ID del rodeo   (ej: "ROD-001")
 * @param {number} [diasObj]   - Días objetivo fijos (opcional)
 * @returns {Object} resultado con oferta, demanda, días, fechas y alertas
 */
function calcularPastoreo(potreroId, rodeoId, diasObj) {
  try {
    const potrero = _getPotrero(potreroId);
    const rodeo   = _getRodeo(rodeoId);

    _validarEntradas(potrero, rodeo);

    const ofertaTotal   = _calcularOferta(potrero);
    const demandaDiaria = _calcularDemanda(rodeo);
    const diasPosibles  = _calcularDias(ofertaTotal, demandaDiaria);
    const diasFinal     = diasObj ? Math.min(diasObj, diasPosibles) : diasPosibles;
    const cargaMaxima   = _calcularCargaMaxima(ofertaTotal, rodeo, diasFinal);
    const fechas        = _calcularFechas(potrero, diasFinal);
    const alertas       = _generarAlertas(potrero, rodeo, diasPosibles, diasObj);

    return {
      ok: true,
      potrero: {
        id:         potrero.id,
        nombre:     potrero.nombre,
        ha:         potrero.ha,
        recurso:    potrero.tipo_recurso,
        estado:     potrero.estado,
      },
      rodeo: {
        id:         rodeo.id,
        nombre:     rodeo.nombre,
        categoria:  rodeo.categoria,
        cabezas:    rodeo.cabezas,
        pesoPromKg: rodeo.peso_prom_kg,
        consumoPct: rodeo.consumo_pct_pv,
      },
      calculo: {
        ofertaTotalKgMs:      Math.round(ofertaTotal),
        demandaDiariaKgMs:    Math.round(demandaDiaria * 10) / 10,
        diasPastoreoPosibles: diasPosibles,
        diasPastoreoFinal:    diasFinal,
        cargaMaximaCabezas:   cargaMaxima,
      },
      fechas: {
        entrada:      fechas.entrada,
        salidaEst:    fechas.salida,
        retornoPrev:  fechas.retorno,
      },
      alertas: alertas,
      generadoEn: new Date().toISOString(),
    };

  } catch (e) {
    Logger.log("calcularPastoreo ERROR: " + e.message);
    return { ok: false, error: e.message };
  }
}

// ─── CÁLCULOS INTERNOS ─────────────────────────────────────────────────────

/**
 * Oferta forrajera total disponible en el potrero.
 * Oferta (kg MS) = ha × kg_ms_ha × (aprovechamiento% / 100)
 */
function _calcularOferta(potrero) {
  return potrero.ha * potrero.kg_ms_ha * (potrero.aprovec_pct / 100);
}

/**
 * Demanda diaria total del rodeo.
 * Demanda (kg MS/día) = cabezas × peso_prom_kg × (consumo_pct_pv / 100)
 */
function _calcularDemanda(rodeo) {
  return rodeo.cabezas * rodeo.peso_prom_kg * (rodeo.consumo_pct_pv / 100);
}

/**
 * Días de pastoreo posibles con la oferta disponible.
 * Días = oferta_total / demanda_diaria  (redondeado hacia abajo)
 */
function _calcularDias(ofertaTotal, demandaDiaria) {
  if (demandaDiaria <= 0) throw new Error("Demanda diaria es cero o negativa.");
  return Math.floor(ofertaTotal / demandaDiaria);
}

/**
 * Carga máxima para un número fijo de días objetivo.
 * Carga (cabezas) = oferta_total / (días × peso_prom × consumo_pct / 100)
 */
function _calcularCargaMaxima(ofertaTotal, rodeo, dias) {
  if (dias <= 0) return 0;
  const demandaPorCabeza = dias * rodeo.peso_prom_kg * (rodeo.consumo_pct_pv / 100);
  return Math.floor(ofertaTotal / demandaPorCabeza);
}

/**
 * Fechas de entrada, salida estimada y retorno previsto.
 */
function _calcularFechas(potrero, diasOcupacion) {
  const entrada = potrero.fecha_disponible
    ? new Date(potrero.fecha_disponible)
    : new Date();

  const salida = new Date(entrada);
  salida.setDate(salida.getDate() + diasOcupacion);

  const retorno = new Date(salida);
  retorno.setDate(retorno.getDate() + potrero.periodo_descanso_dias);

  const fmt = (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");

  return {
    entrada:  fmt(entrada),
    salida:   fmt(salida),
    retorno:  fmt(retorno),
  };
}

// ─── ALERTAS AGRONÓMICAS ───────────────────────────────────────────────────

/**
 * Genera alertas basadas en reglas agronómicas.
 * Cada alerta tiene: { nivel: "info"|"warning"|"error", mensaje: string }
 */
function _generarAlertas(potrero, rodeo, diasPosibles, diasObj) {
  const alertas = [];

  // Potrero no disponible
  if (potrero.estado !== "disponible") {
    alertas.push({
      nivel: "warning",
      mensaje: `El potrero está en estado "${potrero.estado}". Verificar fecha de disponibilidad.`
    });
  }

  // Muy pocos días de pastoreo
  if (diasPosibles < 3) {
    alertas.push({
      nivel: "error",
      mensaje: `Solo ${diasPosibles} días de pastoreo disponibles. Considerar diferir el ingreso o reducir la carga.`
    });
  }

  // Días objetivo mayor a días posibles
  if (diasObj && diasObj > diasPosibles) {
    alertas.push({
      nivel: "warning",
      mensaje: `Los ${diasObj} días objetivo superan los ${diasPosibles} días posibles. Se ajustó automáticamente.`
    });
  }

  // Aprovechamiento muy alto (riesgo de sobrepastoreo)
  if (potrero.aprovec_pct > 80) {
    alertas.push({
      nivel: "warning",
      mensaje: `Aprovechamiento del ${potrero.aprovec_pct}% es elevado. Riesgo de sobrepastoreo y daño de plantas base.`
    });
  }

  // Rodeo muy grande para el potrero
  const cargaInstantanea = rodeo.cabezas / potrero.ha;
  if (cargaInstantanea > 5) {
    alertas.push({
      nivel: "info",
      mensaje: `Carga instantánea de ${cargaInstantanea.toFixed(1)} cab/ha. Alta presión de pastoreo — controlar altura de salida.`
    });
  }

  // Consumo por encima del típico para la categoría
  const ref = CONSUMO_POR_CATEGORIA[rodeo.categoria];
  if (ref && rodeo.consumo_pct_pv > ref.tipico * 1.15) {
    alertas.push({
      nivel: "info",
      mensaje: `Consumo del ${rodeo.consumo_pct_pv}% PV es mayor al típico para ${rodeo.categoria} (${ref.tipico}%). Revisar si es correcto.`
    });
  }

  if (alertas.length === 0) {
    alertas.push({ nivel: "info", mensaje: "Sin alertas. Los parámetros están dentro de rangos normales." });
  }

  return alertas;
}

// ─── LOOKUP DE DATOS EN SHEETS ─────────────────────────────────────────────

/**
 * Busca un potrero por ID en la pestaña Potreros.
 * Asume columnas: A=potrero_id, C=nombre, D=ha, E=tipo_recurso,
 *   F=estado, G=fecha_disponible, H=kg_ms_ha, I=aprovec_pct, J=periodo_descanso_dias
 */
function _getPotrero(potreroId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Potreros");
  if (!sheet) throw new Error("No se encontró la pestaña 'Potreros'.");

  const data = sheet.getDataRange().getValues();
  // Fila 0 = título, Fila 1 = descripción, Fila 2 = headers → datos desde fila 3 (índice 3)
  for (let i = 3; i < data.length; i++) {
    if (data[i][0] === potreroId) {
      return {
        id:                   data[i][0],
        est_id:               data[i][1],
        nombre:               data[i][2],
        ha:                   parseFloat(data[i][3]),
        tipo_recurso:         data[i][4],
        estado:               data[i][5],
        fecha_disponible:     data[i][6],
        kg_ms_ha:             parseFloat(data[i][7]),
        aprovec_pct:          parseFloat(data[i][8]),
        periodo_descanso_dias:parseFloat(data[i][9]),
      };
    }
  }
  throw new Error(`Potrero "${potreroId}" no encontrado en la pestaña Potreros.`);
}

/**
 * Busca un rodeo por ID en la pestaña Rodeos.
 * Asume columnas: A=rodeo_id, C=nombre, D=categoria, E=cabezas,
 *   F=peso_prom_kg, G=consumo_pct_pv
 */
function _getRodeo(rodeoId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Rodeos");
  if (!sheet) throw new Error("No se encontró la pestaña 'Rodeos'.");

  const data = sheet.getDataRange().getValues();
  for (let i = 3; i < data.length; i++) {
    if (data[i][0] === rodeoId) {
      return {
        id:             data[i][0],
        est_id:         data[i][1],
        nombre:         data[i][2],
        categoria:      data[i][3],
        cabezas:        parseInt(data[i][4]),
        peso_prom_kg:   parseFloat(data[i][5]),
        consumo_pct_pv: parseFloat(data[i][6]),
      };
    }
  }
  throw new Error(`Rodeo "${rodeoId}" no encontrado en la pestaña Rodeos.`);
}

// ─── TASA DE CRECIMIENTO ───────────────────────────────────────────────────

/**
 * Obtiene la tasa de crecimiento diaria (kg MS/ha/día) desde Tabla_forraje.
 * Lookup por tipo_recurso + zona + mes (número 1-12).
 *
 * @param {string} tipo_recurso
 * @param {string} zona
 * @param {number} mes  - 1 a 12
 * @returns {number} tasa en kg MS/ha/día, o null si no hay dato
 */
function obtenerTasaCrecimiento(tipo_recurso, zona, mes) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Tabla_forraje");
  if (!sheet) throw new Error("No se encontró la pestaña 'Tabla_forraje'.");

  const data = sheet.getDataRange().getValues();
  // Columnas: A=tipo_recurso, B=zona, C=mes(número), E=tasa
  for (let i = 3; i < data.length; i++) {
    if (
      data[i][0] === tipo_recurso &&
      data[i][1] === zona &&
      parseInt(data[i][2]) === mes
    ) {
      return parseFloat(data[i][4]);
    }
  }
  Logger.log(`Sin tasa de crecimiento para: ${tipo_recurso} / ${zona} / mes ${mes}`);
  return null;
}

// ─── ESCRITURA EN CALCULADORA ──────────────────────────────────────────────

/**
 * Escribe el resultado de calcularPastoreo() en la pestaña Calculadora.
 * Útil cuando se invoca desde el frontend o desde un trigger.
 *
 * @param {string} potreroId
 * @param {string} rodeoId
 * @param {number} [diasObjetivo]
 */
function actualizarCalculadora(potreroId, rodeoId, diasObjetivo) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Calculadora");
  if (!sheet) throw new Error("No se encontró la pestaña 'Calculadora'.");

  // Escribir las entradas — las fórmulas de la hoja hacen el resto
  sheet.getRange("B4").setValue(potreroId);
  sheet.getRange("B5").setValue(rodeoId);
  if (diasObjetivo) sheet.getRange("B6").setValue(diasObjetivo);

  SpreadsheetApp.flush();
}

// ─── VALIDACIONES ──────────────────────────────────────────────────────────

function _validarEntradas(potrero, rodeo) {
  if (!potrero.ha || potrero.ha <= 0)
    throw new Error(`Potrero "${potrero.id}": superficie inválida (${potrero.ha} ha).`);
  if (!potrero.kg_ms_ha || potrero.kg_ms_ha <= 0)
    throw new Error(`Potrero "${potrero.id}": kg MS/ha inválido.`);
  if (!potrero.aprovec_pct || potrero.aprovec_pct <= 0 || potrero.aprovec_pct > 100)
    throw new Error(`Potrero "${potrero.id}": % aprovechamiento inválido (${potrero.aprovec_pct}).`);
  if (!rodeo.cabezas || rodeo.cabezas <= 0)
    throw new Error(`Rodeo "${rodeo.id}": número de cabezas inválido.`);
  if (!rodeo.peso_prom_kg || rodeo.peso_prom_kg <= 0)
    throw new Error(`Rodeo "${rodeo.id}": peso promedio inválido.`);
  if (!rodeo.consumo_pct_pv || rodeo.consumo_pct_pv <= 0)
    throw new Error(`Rodeo "${rodeo.id}": consumo % PV inválido.`);
}

// ─── FUNCIÓN DE PRUEBA (ejecutar manualmente desde el editor) ──────────────

/**
 * Ejecutá esta función desde el editor de Apps Script para probar el módulo.
 * Menu: Ejecutar → testCalculadora
 */
function testCalculadora() {
  const resultado = calcularPastoreo("POT-001", "ROD-001", 7);
  Logger.log(JSON.stringify(resultado, null, 2));

  if (resultado.ok) {
    Logger.log("─── RESUMEN ───────────────────────────────");
    Logger.log(`Potrero: ${resultado.potrero.nombre} (${resultado.potrero.ha} ha)`);
    Logger.log(`Rodeo:   ${resultado.rodeo.nombre} (${resultado.rodeo.cabezas} cabezas)`);
    Logger.log(`Oferta:  ${resultado.calculo.ofertaTotalKgMs} kg MS`);
    Logger.log(`Demanda: ${resultado.calculo.demandaDiariaKgMs} kg MS/día`);
    Logger.log(`Días posibles: ${resultado.calculo.diasPastoreoPosibles}`);
    Logger.log(`Días finales:  ${resultado.calculo.diasPastoreoFinal}`);
    Logger.log(`Carga máx p/7 días: ${resultado.calculo.cargaMaximaCabezas} cabezas`);
    Logger.log(`Entrada: ${resultado.fechas.entrada} | Salida: ${resultado.fechas.salidaEst} | Retorno: ${resultado.fechas.retornoPrev}`);
    Logger.log("Alertas:");
    resultado.alertas.forEach(a => Logger.log(`  [${a.nivel.toUpperCase()}] ${a.mensaje}`));
  } else {
    Logger.log("ERROR: " + resultado.error);
  }
}
