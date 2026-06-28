/**
 * PastoreoGestion — Rotacion.gs
 * Módulo de scheduling de rotación y proyección de calendario.
 *
 * Responsabilidades:
 *   - calcularRotacion(rodeoId, diasProyeccion)
 *     Dado un rodeo, proyecta la secuencia óptima de potreros
 *     para los próximos N días ordenando por disponibilidad.
 *
 *   - calcularCalendarioEstablecimiento(diasProyeccion)
 *     Proyecta todos los rodeos activos del establecimiento
 *     y devuelve un calendario consolidado.
 *
 *   - _ordenarPotrerosPorPrioridad(potreros, fechaBase)
 *     Ordena potreros disponibles: primero los que están listos,
 *     luego los que estarán disponibles más pronto.
 *
 * @version 1.0.0
 */

// ─── FUNCIÓN PRINCIPAL — ROTACIÓN DE UN RODEO ──────────────────────────────

/**
 * Calcula la secuencia de rotación óptima para un rodeo.
 *
 * @param {string} rodeoId         - ID del rodeo a proyectar
 * @param {number} diasProyeccion  - Días a proyectar (default 90)
 * @returns {Object} secuencia de pastoreos proyectados + alertas
 */
function calcularRotacion(rodeoId, diasProyeccion) {
  diasProyeccion = diasProyeccion || 90;

  try {
    const rodeo    = _getRodeo(rodeoId);
    const potreros = _getPotrerosByEst(rodeo.est_id);

    if (!potreros.length) {
      return { ok: false, error: "No hay potreros cargados para este establecimiento." };
    }

    const secuencia = [];
    const alertas   = [];
    let fechaActual = new Date();
    let diasRestantes = diasProyeccion;

    // Estado mutable de los potreros durante la proyección
    const estadoPotreros = potreros.map(p => ({
      ...p,
      fechaDisp: p.fecha_disponible ? new Date(p.fecha_disponible) : new Date(),
    }));

    let iteraciones = 0;
    const maxIter   = 50; // seguridad contra loop infinito

    while (diasRestantes > 0 && iteraciones < maxIter) {
      iteraciones++;

      // Ordenar potreros por prioridad para la fecha actual
      const ordenados = _ordenarPotrerosPorPrioridad(estadoPotreros, fechaActual);

      if (!ordenados.length) {
        alertas.push({
          nivel: "error",
          mensaje: "No hay potreros suficientes para cubrir la rotación completa.",
        });
        break;
      }

      const potrero = ordenados[0];

      // Si el próximo potrero no está disponible aún, hay un gap
      const fechaEntrada = potrero.fechaDisp > fechaActual
        ? new Date(potrero.fechaDisp)
        : new Date(fechaActual);

      if (potrero.fechaDisp > fechaActual) {
        const diasEspera = Math.round((potrero.fechaDisp - fechaActual) / 86400000);
        if (diasEspera > 3) {
          alertas.push({
            nivel:   "warning",
            mensaje: `Gap de ${diasEspera} días sin potrero disponible antes de ingresar a ${potrero.nombre}.`,
          });
        }
      }

      // Calcular días posibles en este potrero
      const oferta        = potrero.ha * potrero.kg_ms_ha * (potrero.aprovec_pct / 100);
      const demandaDiaria = rodeo.cabezas * rodeo.peso_prom_kg * (rodeo.consumo_pct_pv / 100);
      const diasPosibles  = demandaDiaria > 0 ? Math.floor(oferta / demandaDiaria) : 0;

      if (diasPosibles < 3) {
        alertas.push({
          nivel:   "warning",
          mensaje: `${potrero.nombre} solo ofrece ${diasPosibles} días de pastoreo. Considerar aumentar kg MS/ha o reducir carga.`,
        });
      }

      const diasOcupacion = Math.min(diasPosibles, diasRestantes);
      const fechaSalida   = new Date(fechaEntrada);
      fechaSalida.setDate(fechaSalida.getDate() + diasOcupacion);

      const fechaRetorno = new Date(fechaSalida);
      fechaRetorno.setDate(fechaRetorno.getDate() + potrero.periodo_descanso_dias);

      secuencia.push({
        orden:        secuencia.length + 1,
        potrero_id:   potrero.potrero_id,
        nombre:       potrero.nombre,
        ha:           potrero.ha,
        tipo_recurso: potrero.tipo_recurso,
        fecha_entrada:  _fmtFecha(fechaEntrada),
        fecha_salida:   _fmtFecha(fechaSalida),
        fecha_retorno:  _fmtFecha(fechaRetorno),
        dias_ocupacion: diasOcupacion,
        oferta_kg_ms:   Math.round(oferta),
        demanda_diaria: Math.round(demandaDiaria),
      });

      // Actualizar estado del potrero usado
      const idx = estadoPotreros.findIndex(p => p.potrero_id === potrero.potrero_id);
      if (idx >= 0) estadoPotreros[idx].fechaDisp = new Date(fechaRetorno);

      // Avanzar en el tiempo
      fechaActual   = new Date(fechaSalida);
      diasRestantes -= diasOcupacion;

      // Agregar el día de espera al conteo si hubo gap
      if (potrero.fechaDisp > new Date()) {
        const gap = Math.round((potrero.fechaDisp - new Date()) / 86400000);
        diasRestantes -= Math.max(0, gap);
      }
    }

    // Alerta si no alcanza la cobertura completa
    if (diasRestantes > 5) {
      alertas.push({
        nivel:   "info",
        mensaje: `La proyección cubre ${diasProyeccion - diasRestantes} de ${diasProyeccion} días. Agregar más potreros para mayor cobertura.`,
      });
    }

    return {
      ok:              true,
      rodeo:           { id: rodeo.rodeo_id, nombre: rodeo.nombre, cabezas: rodeo.cabezas },
      dias_proyeccion: diasProyeccion,
      secuencia:       secuencia,
      alertas:         alertas,
      generado_en:     new Date().toISOString(),
    };

  } catch(e) {
    Logger.log("calcularRotacion ERROR: " + e.message);
    return { ok: false, error: e.message };
  }
}

// ─── CALENDARIO DEL ESTABLECIMIENTO ────────────────────────────────────────

/**
 * Proyecta todos los rodeos activos y devuelve un calendario consolidado.
 * Útil para la vista de cronograma en la app.
 *
 * @param {number} diasProyeccion - Días a proyectar (default 90)
 * @returns {Object} eventos del calendario ordenados por fecha
 */
function calcularCalendarioEstablecimiento(diasProyeccion) {
  diasProyeccion = diasProyeccion || 90;

  try {
    const config = _leerConfig();
    const estId  = config["est_id_activo"] || "EST-001";
    const rodeos = _getRodeosByEst(estId);

    if (!rodeos.length) {
      return { ok: false, error: "No hay rodeos cargados para este establecimiento." };
    }

    const eventos  = [];
    const alertas  = [];

    rodeos.forEach(rodeo => {
      const rotacion = calcularRotacion(rodeo.rodeo_id, diasProyeccion);
      if (!rotacion.ok) return;

      rotacion.secuencia.forEach(item => {
        eventos.push({
          rodeo_id:     rodeo.rodeo_id,
          rodeo_nombre: rodeo.nombre,
          ...item,
        });
      });

      // Agregar alertas del rodeo
      rotacion.alertas.forEach(a => {
        if (a.nivel !== "info") {
          alertas.push({ ...a, rodeo: rodeo.nombre });
        }
      });
    });

    // Ordenar eventos por fecha de entrada
    eventos.sort((a, b) => {
      const da = _parseFecha(a.fecha_entrada);
      const db = _parseFecha(b.fecha_entrada);
      return da - db;
    });

    // Detectar conflictos: mismo potrero ocupado por dos rodeos al mismo tiempo
    const conflictos = _detectarConflictos(eventos);
    conflictos.forEach(c => alertas.push(c));

    return {
      ok:              true,
      est_id:          estId,
      dias_proyeccion: diasProyeccion,
      eventos:         eventos,
      alertas:         alertas,
      total_eventos:   eventos.length,
      generado_en:     new Date().toISOString(),
    };

  } catch(e) {
    Logger.log("calcularCalendario ERROR: " + e.message);
    return { ok: false, error: e.message };
  }
}

// ─── HELPERS DE SCHEDULING ─────────────────────────────────────────────────

/**
 * Ordena los potreros por prioridad para una fecha dada.
 * Criterios:
 *   1. Potreros disponibles ahora (fechaDisp <= fechaBase)
 *   2. Entre los disponibles, los de mayor oferta primero
 *   3. Potreros que estarán disponibles pronto (ordenados por fecha)
 */
function _ordenarPotrerosPorPrioridad(potreros, fechaBase) {
  const disponiblesAhora = potreros
    .filter(p => p.fechaDisp <= fechaBase && p.kg_ms_ha > 0)
    .sort((a, b) => {
      const ofertaA = a.ha * a.kg_ms_ha * (a.aprovec_pct / 100);
      const ofertaB = b.ha * b.kg_ms_ha * (b.aprovec_pct / 100);
      return ofertaB - ofertaA; // mayor oferta primero
    });

  const proximamente = potreros
    .filter(p => p.fechaDisp > fechaBase && p.kg_ms_ha > 0)
    .sort((a, b) => a.fechaDisp - b.fechaDisp); // más pronto primero

  return [...disponiblesAhora, ...proximamente];
}

/**
 * Detecta conflictos de potreros asignados a más de un rodeo simultáneamente.
 */
function _detectarConflictos(eventos) {
  const conflictos = [];

  for (let i = 0; i < eventos.length; i++) {
    for (let j = i + 1; j < eventos.length; j++) {
      const a = eventos[i];
      const b = eventos[j];

      if (a.potrero_id !== b.potrero_id) continue;
      if (a.rodeo_id   === b.rodeo_id)   continue;

      const inicioA = _parseFecha(a.fecha_entrada);
      const finA    = _parseFecha(a.fecha_salida);
      const inicioB = _parseFecha(b.fecha_entrada);
      const finB    = _parseFecha(b.fecha_salida);

      // Solapamiento
      if (inicioA < finB && inicioB < finA) {
        conflictos.push({
          nivel:   "error",
          mensaje: `Conflicto: ${a.nombre} asignado a "${a.rodeo_nombre}" y "${b.rodeo_nombre}" al mismo tiempo.`,
        });
      }
    }
  }

  return conflictos;
}

// ─── LECTURA DE DATOS ───────────────────────────────────────────────────────

function _getPotrerosByEst(estId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Potreros");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 3; i < data.length; i++) {
    const f = data[i];
    if (!f[0]) continue;
    if (estId && f[1] !== estId) continue;
    result.push({
      potrero_id:            f[0],
      est_id:                f[1],
      nombre:                f[2],
      ha:                    parseFloat(f[3]) || 0,
      tipo_recurso:          f[4],
      estado:                f[5],
      fecha_disponible:      f[6] ? _formatFechaRot(f[6]) : null,
      kg_ms_ha:              parseFloat(f[7]) || 0,
      aprovec_pct:           parseFloat(f[8]) || 70,
      periodo_descanso_dias: parseInt(f[9])   || 45,
    });
  }
  return result;
}

function _getRodeosByEst(estId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Rodeos");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 3; i < data.length; i++) {
    const f = data[i];
    if (!f[0]) continue;
    if (estId && f[1] !== estId) continue;
    result.push({
      rodeo_id:       f[0],
      est_id:         f[1],
      nombre:         f[2],
      categoria:      f[3],
      cabezas:        parseInt(f[4])   || 0,
      peso_prom_kg:   parseFloat(f[5]) || 0,
      consumo_pct_pv: parseFloat(f[6]) || 2.5,
    });
  }
  return result;
}

// ─── HELPERS DE FECHA ───────────────────────────────────────────────────────

function _fmtFecha(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function _formatFechaRot(valor) {
  if (!valor) return null;
  if (valor instanceof Date)
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(valor);
}

function _parseFecha(str) {
  if (!str) return new Date();
  // Soporta dd/MM/yyyy y yyyy-MM-dd
  if (str.includes("/")) {
    const [d, m, y] = str.split("/");
    return new Date(y, m - 1, d);
  }
  return new Date(str);
}

// ─── FUNCIÓN DE PRUEBA ──────────────────────────────────────────────────────

/**
 * Ejecutá desde el editor para probar el módulo.
 * Menu: Ejecutar → testRotacion
 */
function testRotacion() {
  const resultado = calcularRotacion("ROD-001", 90);
  Logger.log(JSON.stringify(resultado, null, 2));

  if (resultado.ok) {
    Logger.log("─── SECUENCIA ─────────────────────────────");
    resultado.secuencia.forEach(s => {
      Logger.log(`${s.orden}. ${s.nombre} | ${s.fecha_entrada} → ${s.fecha_salida} (${s.dias_ocupacion} días)`);
    });
    Logger.log("─── ALERTAS ───────────────────────────────");
    resultado.alertas.forEach(a => Logger.log(`[${a.nivel.toUpperCase()}] ${a.mensaje}`));
  }

  const calendario = calcularCalendarioEstablecimiento(90);
  Logger.log("─── CALENDARIO ────────────────────────────");
  Logger.log(`Total eventos: ${calendario.total_eventos}`);
}
