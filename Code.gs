/**
 * ============================================================================
 *  Reserva de Salas – OGPP UNMSM  ·  Aplicación web (Apps Script + Google Sheets)
 * ----------------------------------------------------------------------------
 *  La página (Index.html) llama a las funciones api* de este archivo con
 *  google.script.run. La identidad SIEMPRE se toma de la cuenta Google de quien
 *  abre la página (Session.getActiveUser), nunca de datos enviados por el navegador.
 *
 *  Primer uso: ejecutar setup() una vez desde el editor y publicar la web app.
 * ============================================================================
 */

const TZ = 'America/Lima';

const HOJAS = {
  RESERVAS: 'Reservas',
  USUARIOS: 'Usuarios',
  SALAS: 'Salas',
  CONFIG: 'Config',
  FERIADOS: 'Feriados',
  LOG: 'Log'
};

const ENCABEZADOS = {
  Reservas: ['ID_Reserva', 'Fecha_Registro', 'Sala', 'Fecha_Reserva', 'Hora_Inicio', 'Hora_Fin',
    'Oficina_Solicitante', 'Responsable', 'Correo_Contacto', 'Tema_Reunion', 'Asistentes',
    'Estado', 'Cancelado_Por', 'Fecha_Cancelacion', 'Motivo_Cancelacion'],
  Usuarios: ['Correo', 'Nombre', 'Oficina', 'Rol', 'Activo'],
  Salas: ['Sala', 'Capacidad', 'Ubicacion', 'Equipamiento', 'Correo_Responsable', 'Activa'],
  Config: ['Clave', 'Valor', 'Descripcion'],
  Feriados: ['Fecha', 'Descripcion'],
  Log: ['Fecha', 'Accion', 'Correo', 'Codigo', 'Detalle']
};

const CONFIG_DEFECTO = [
  ['HORA_APERTURA', '08:00', 'Hora mínima de inicio de una reserva (HH:mm).'],
  ['HORA_CIERRE', '17:00', 'Hora máxima de fin de una reserva (HH:mm).'],
  ['DIAS_HABILES', '1,2,3,4,5', 'Días permitidos: 1=lunes ... 7=domingo.'],
  ['DURACION_MIN', '30', 'Duración mínima en minutos.'],
  ['DURACION_MAX', '240', 'Duración máxima en minutos.'],
  ['ANTICIPACION_MAX_DIAS', '60', 'Cuántos días hacia adelante se puede reservar.'],
  ['PASO_MINUTOS', '30', 'Intervalo de las horas del formulario y de las sugerencias (15, 30 o 60).'],
  ['CORREOS_NOTIFICACION', '', 'Correos de los encargados que reciben aviso de TODA reserva o cancelación (separados por coma). Cada sala puede tener además su propio encargado en la pestaña Salas.'],
  ['ADMINS', '', 'Correos con acceso de administrador además de los usuarios con Rol=ADMIN (separados por coma).'],
  ['NOMBRE_SERVICIO', 'Reserva de Salas – OGPP UNMSM', 'Nombre que aparece en la página y en los correos.']
];

// Feriados nacionales restantes de 2026 (verificar y agregar días no laborables decretados).
const FERIADOS_DEFECTO = [
  ['2026-10-08', 'Combate de Angamos'],
  ['2026-11-01', 'Día de Todos los Santos'],
  ['2026-12-08', 'Inmaculada Concepción'],
  ['2026-12-09', 'Batalla de Ayacucho'],
  ['2026-12-25', 'Navidad']
];

const ESTADO = { CONFIRMADA: 'CONFIRMADA', CANCELADA: 'CANCELADA' };
const DIAS = ['', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo'];
const DIAS_CORTOS = ['', 'lun', 'mar', 'mié', 'jue', 'vie', 'sáb', 'dom'];

/* ============================================================================
 *  1. PUNTOS DE ENTRADA
 * ========================================================================== */

/** Sirve la página. Solo cuentas del dominio que estén en la hoja Usuarios (o en Config › ADMINS). */
function doGet() {
  let usuario = null;
  try { usuario = usuarioActual_(); } catch (e) { /* se trata como sin acceso */ }
  if (!usuario) {
    const correo = correoActivo_();
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif;padding:40px;max-width:560px;margin:auto">' +
      '<h2 style="color:#7a1f2b">Acceso restringido</h2>' +
      '<p>Tu cuenta no está autorizada para reservar salas de la OGPP.</p>' +
      '<p>Solicita acceso a la oficina indicando este correo: <b>' + escaparHtml_(correo || '(no se pudo detectar tu cuenta; inicia sesión con tu correo @unmsm.edu.pe)') + '</b></p></div>')
      .setTitle('Reserva de Salas OGPP')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Reserva de Salas OGPP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Menú en la hoja de cálculo. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Reservas OGPP')
    .addItem('Configurar / reparar hojas', 'setup')
    .addItem('Probar con mi cuenta', 'probar')
    .addToUi();
}

/* ============================================================================
 *  2. API PARA LA PÁGINA  (google.script.run)
 *  Todas devuelven { ok, codigo, mensaje, data } y nunca lanzan excepciones.
 * ========================================================================== */

/** Ejecuta fn(usuario) con la cuenta de quien llama. */
function api_(nombre, fn) {
  let usuario = null;
  let res;
  try {
    usuario = usuarioActual_();
    if (!usuario) {
      return respuesta_(false, 'NO_AUTORIZADO',
        'Tu cuenta no está autorizada para reservar salas de la OGPP. Comunícate con la oficina para solicitar acceso.');
    }
    res = fn(usuario);
  } catch (err) {
    res = respuesta_(false, 'ERROR', 'Ocurrió un error interno. Intenta nuevamente en unos minutos.',
      { detalle: String((err && err.message) || err) });
  }
  try { registrarLog_(nombre, usuario && usuario.Correo, res.codigo, res.ok ? '' : res.mensaje); } catch (ignorado) { /* el log nunca rompe la respuesta */ }
  return res;
}

/** Datos para armar la página: usuario, salas, reglas y feriados próximos. */
function apiInicio() {
  return api_('inicio', usuario => {
    const cfg = leerConfig_();
    const ahora = ahora_();
    const limite = sumarDias_(ahora.fecha, cfg.ANTICIPACION_MAX_DIAS);
    const feriados = leerFeriados_();
    return respuesta_(true, 'EXITO', 'Listo.', {
      usuario: { nombre: usuario.Nombre, oficina: usuario.Oficina, correo: usuario.Correo, rol: usuario.Rol },
      servicio: cfg.NOMBRE_SERVICIO,
      salas: leerSalas_().map(s => ({ sala: s.Sala, capacidad: entero_(s.Capacidad), ubicacion: s.Ubicacion, equipamiento: s.Equipamiento })),
      cfg: {
        apertura: cfg.HORA_APERTURA, cierre: cfg.HORA_CIERRE, dias_habiles: cfg.DIAS_HABILES,
        duracion_min: cfg.DURACION_MIN, duracion_max: cfg.DURACION_MAX,
        anticipacion_max_dias: cfg.ANTICIPACION_MAX_DIAS, paso: cfg.PASO_MINUTOS
      },
      hoy: ahora.fecha, ahora_min: ahora.min, limite: limite,
      feriados: Object.keys(feriados).filter(f => f >= ahora.fecha && f <= limite).sort()
        .map(f => ({ fecha: f, descripcion: feriados[f] }))
    });
  });
}

/** Verifica una solicitud SIN registrarla (la página la llama mientras el usuario llena el formulario). */
function apiVerificar(form) {
  return api_('verificar', () => {
    form = form || {};
    const cfg = leerConfig_();
    const salas = leerSalas_();
    const v = validar_(form, cfg, salas);
    if (v.errores.length) return datosInvalidos_(v.errores);
    return evaluarSolicitud_(v, form, cfg, salas, reservasDelDia_(v.fecha), false);
  });
}

/** Registra la reserva (revalida dentro de un bloqueo para evitar dobles reservas). */
function apiReservar(form) {
  return api_('reservar', usuario => {
    form = form || {};
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) {
      return respuesta_(false, 'ERROR', 'El sistema está atendiendo otra solicitud. Intenta de nuevo en unos segundos.');
    }
    let reserva;
    try {
      const cfg = leerConfig_();
      const salas = leerSalas_();
      const v = validar_(form, cfg, salas);
      if (v.errores.length) return datosInvalidos_(v.errores);

      const evaluacion = evaluarSolicitud_(v, form, cfg, salas, reservasDelDia_(v.fecha), true);
      if (evaluacion.codigo !== 'DISPONIBLE') return evaluacion;

      reserva = {
        ID_Reserva: generarId_(),
        Fecha_Registro: ahoraTexto_(),
        Sala: evaluacion.data.sala,
        Fecha_Reserva: v.fecha,
        Hora_Inicio: v.hi,
        Hora_Fin: v.hf,
        Oficina_Solicitante: usuario.Oficina,
        Responsable: usuario.Nombre,
        Correo_Contacto: usuario.Correo,
        Tema_Reunion: limpiarTexto_(form.tema) || 'Reunión de trabajo',
        Asistentes: entero_(form.asistentes) || '',
        Estado: ESTADO.CONFIRMADA,
        Cancelado_Por: '', Fecha_Cancelacion: '', Motivo_Cancelacion: ''
      };
      agregarFila_(HOJAS.RESERVAS, reserva);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    const envio = enviarNotificaciones_('CONFIRMACION', reserva);
    return respuesta_(true, 'EXITO', 'Reserva confirmada. Código: ' + reserva.ID_Reserva + '.',
      { reserva: reservaVista_(reserva, usuario), correo_usuario: envio.usuario, correo_encargados: envio.encargados });
  });
}

/** Próximas reservas y últimas reservas pasadas o canceladas del usuario. */
function apiMisReservas() {
  return api_('mis_reservas', usuario => {
    const ahora = ahora_();
    const mias = leerTabla_(HOJAS.RESERVAS).filas.filter(r => r.Correo_Contacto === usuario.Correo);
    const esProxima = r => r.Estado === ESTADO.CONFIRMADA &&
      (r.Fecha_Reserva > ahora.fecha || (r.Fecha_Reserva === ahora.fecha && aMin_(r.Hora_Fin) > ahora.min));
    const clave = r => r.Fecha_Reserva + ' ' + r.Hora_Inicio;
    const proximas = mias.filter(esProxima).sort((a, b) => clave(a).localeCompare(clave(b)));
    const historial = mias.filter(r => !esProxima(r)).sort((a, b) => clave(b).localeCompare(clave(a))).slice(0, 15);
    return respuesta_(true, 'EXITO', proximas.length ? '' : 'No tienes reservas próximas.', {
      proximas: proximas.map(r => reservaVista_(r, usuario)),
      historial: historial.map(r => reservaVista_(r, usuario))
    });
  });
}

/** Cancela una reserva propia (o cualquiera si el usuario es ADMIN). */
function apiCancelar(idReserva, motivo) {
  return api_('cancelar', usuario => cancelarReserva_(idReserva, usuario, motivo));
}

/** Ocupación de la semana (lunes a viernes, según días hábiles) para el calendario. */
function apiSemana(fechaReferencia) {
  return api_('semana', usuario => {
    const cfg = leerConfig_();
    const ref = normFecha_(fechaReferencia) || ahora_().fecha;
    const lunes = sumarDias_(ref, 1 - diaSemana_(ref));
    const domingo = sumarDias_(lunes, 6);
    const feriados = leerFeriados_();
    const dias = [];
    for (let i = 0; i < 7; i++) {
      const f = sumarDias_(lunes, i);
      if (cfg.DIAS_HABILES.indexOf(diaSemana_(f)) !== -1) {
        dias.push({ fecha: f, etiqueta: DIAS_CORTOS[diaSemana_(f)] + ' ' + fechaCorta_(f), feriado: feriados[f] || '' });
      }
    }
    const reservas = leerTabla_(HOJAS.RESERVAS).filas
      .filter(r => r.Estado === ESTADO.CONFIRMADA && r.Fecha_Reserva >= lunes && r.Fecha_Reserva <= domingo)
      .map(r => reservaVista_(r, usuario));
    return respuesta_(true, 'EXITO', '', {
      lunes: lunes, dias: dias, reservas: reservas,
      salas: leerSalas_().map(s => ({ sala: s.Sala, capacidad: entero_(s.Capacidad) })),
      apertura: cfg.HORA_APERTURA, cierre: cfg.HORA_CIERRE, paso: cfg.PASO_MINUTOS
    });
  });
}

/**
 * Listado para administración. Filtros: { desde, hasta, estado: 'TODAS'|'CONFIRMADA'|'CANCELADA', sala }.
 * Solo ADMIN. Devuelve las reservas del rango y estadísticas.
 */
function apiAdminLista(filtros) {
  return api_('admin_lista', usuario => {
    if (usuario.Rol !== 'ADMIN') return respuesta_(false, 'NO_PERMITIDO', 'Esta sección es solo para administradores.');
    filtros = filtros || {};
    const cfg = leerConfig_();
    const hoy = ahora_().fecha;
    const desde = normFecha_(filtros.desde) || hoy;
    let hasta = normFecha_(filtros.hasta) || sumarDias_(desde, 30);
    if (hasta < desde) hasta = desde;
    if (diasEntre_(desde, hasta) > 366) hasta = sumarDias_(desde, 366);

    const todas = leerTabla_(HOJAS.RESERVAS).filas.filter(r => r.Fecha_Reserva >= desde && r.Fecha_Reserva <= hasta);
    const clave = r => r.Fecha_Reserva + ' ' + r.Hora_Inicio;
    const sel = todas
      .filter(r => !filtros.estado || filtros.estado === 'TODAS' || r.Estado === filtros.estado)
      .filter(r => !filtros.sala || r.Sala === filtros.sala)
      .sort((a, b) => clave(a).localeCompare(clave(b)));

    // Estadísticas sobre las reservas confirmadas del rango (y sala, si se filtró)
    const salas = leerSalas_().filter(s => !filtros.sala || s.Sala === filtros.sala);
    const confirmadas = todas.filter(r => r.Estado === ESTADO.CONFIRMADA && (!filtros.sala || r.Sala === filtros.sala));
    const minutos = confirmadas.reduce((s, r) => s + aMin_(r.Hora_Fin) - aMin_(r.Hora_Inicio), 0);
    const feriados = leerFeriados_();
    let diasHabiles = 0;
    for (let f = desde; f <= hasta; f = sumarDias_(f, 1)) {
      if (cfg.DIAS_HABILES.indexOf(diaSemana_(f)) !== -1 && !feriados[f]) diasHabiles++;
    }
    const capacidadMin = diasHabiles * salas.length * (aMin_(cfg.HORA_CIERRE) - aMin_(cfg.HORA_APERTURA));
    const porOficina = {};
    confirmadas.forEach(r => { porOficina[r.Oficina_Solicitante] = (porOficina[r.Oficina_Solicitante] || 0) + (aMin_(r.Hora_Fin) - aMin_(r.Hora_Inicio)); });
    const top = Object.keys(porOficina).map(o => ({ oficina: o, horas: Math.round(porOficina[o] / 6) / 10 }))
      .sort((a, b) => b.horas - a.horas).slice(0, 5);

    return respuesta_(true, 'EXITO', '', {
      desde: desde, hasta: hasta, total: sel.length, truncado: sel.length > 500,
      reservas: sel.slice(0, 500).map(r => reservaVista_(r, usuario)),
      stats: {
        confirmadas: confirmadas.length,
        canceladas: todas.filter(r => r.Estado === ESTADO.CANCELADA && (!filtros.sala || r.Sala === filtros.sala)).length,
        horas_reservadas: Math.round(minutos / 6) / 10,
        ocupacion_pct: capacidadMin ? Math.round(1000 * minutos / capacidadMin) / 10 : 0,
        top_oficinas: top
      }
    });
  });
}

/* ============================================================================
 *  3. LÓGICA DE NEGOCIO
 * ========================================================================== */

/**
 * Valida formato y reglas. Devuelve { errores[], fecha, hi, hf, ini, fin, sala }.
 * `sala` es opcional en la solicitud; si viene, debe existir.
 */
function validar_(p, cfg, salas) {
  const errores = [];
  const fecha = normFecha_(p.fecha);
  const hi = normHora_(p.hora_inicio);
  const hf = normHora_(p.hora_fin);
  if (!fecha) errores.push('Selecciona una fecha válida.');
  if (!hi) errores.push('Selecciona la hora de inicio.');
  if (!hf) errores.push('Selecciona la hora de fin.');
  let sala = null;
  if (p.sala) {
    sala = resolverSala_(p.sala, salas);
    if (!sala) errores.push('La sala "' + p.sala + '" no existe o no está activa. Salas disponibles: ' + salas.map(s => s.Sala).join(', ') + '.');
  }
  if (errores.length) return { errores: errores };

  const ini = aMin_(hi), fin = aMin_(hf);
  const apertura = aMin_(cfg.HORA_APERTURA), cierre = aMin_(cfg.HORA_CIERRE);
  if (fin <= ini) {
    errores.push('La hora de fin debe ser posterior a la de inicio.');
  } else {
    const dur = fin - ini;
    if (dur < cfg.DURACION_MIN) errores.push('La duración mínima es de ' + duracionTexto_(cfg.DURACION_MIN) + '.');
    if (dur > cfg.DURACION_MAX) errores.push('La duración máxima es de ' + duracionTexto_(cfg.DURACION_MAX) + '.');
  }
  if (ini < apertura || fin > cierre) {
    errores.push('El horario de reservas es de ' + cfg.HORA_APERTURA + ' a ' + cfg.HORA_CIERRE + '.');
  }
  const noHabil = motivoDiaNoHabil_(fecha, cfg);
  if (noHabil) errores.push(noHabil);

  const ahora = ahora_();
  if (fecha < ahora.fecha || (fecha === ahora.fecha && ini <= ahora.min)) {
    errores.push('No se puede reservar en una fecha u hora que ya pasó.');
  } else if (diasEntre_(ahora.fecha, fecha) > cfg.ANTICIPACION_MAX_DIAS) {
    errores.push('Solo se puede reservar con hasta ' + cfg.ANTICIPACION_MAX_DIAS + ' días de anticipación.');
  }
  return { errores: errores, fecha: fecha, hi: hi, hf: hf, ini: ini, fin: fin, sala: sala };
}

/** Devuelve el motivo si la fecha no es hábil (fin de semana o feriado), o '' si es hábil. */
function motivoDiaNoHabil_(fecha, cfg) {
  const dow = diaSemana_(fecha);
  if (cfg.DIAS_HABILES.indexOf(dow) === -1) {
    const nombres = cfg.DIAS_HABILES.map(d => DIAS[d]);
    return 'El ' + DIAS[dow] + ' no es día hábil para reservas (días permitidos: ' + nombres.join(', ') + ').';
  }
  const feriado = leerFeriados_()[fecha];
  if (feriado) return 'El ' + fechaLarga_(fecha) + ' es feriado (' + feriado + ').';
  return '';
}

/**
 * Evalúa disponibilidad de la solicitud ya validada.
 * Códigos: DISPONIBLE | CONFLICTO_SALA | CONFLICTO_AMBAS
 */
function evaluarSolicitud_(v, req, cfg, salas, reservas, esRegistro) {
  const estado = salas.map(s => {
    const choques = conflictos_(reservas, s.Sala, v.ini, v.fin);
    return {
      sala: s.Sala, capacidad: entero_(s.Capacidad), libre: choques.length === 0,
      ocupada_por: choques.map(r => ({ oficina: r.Oficina_Solicitante, responsable: r.Responsable, hora_inicio: r.Hora_Inicio, hora_fin: r.Hora_Fin, tema: r.Tema_Reunion }))
    };
  });
  const libres = estado.filter(e => e.libre).map(e => e.sala);
  const base = { fecha: v.fecha, hora_inicio: v.hi, hora_fin: v.hf, estado_salas: estado, salas_libres: libres };
  const asistentes = entero_(req.asistentes);
  const horario = fechaLarga_(v.fecha) + ' de ' + v.hi + ' a ' + v.hf;

  // Caso 1: el usuario eligió sala
  if (v.sala) {
    const objetivo = estado.filter(e => e.sala === v.sala.Sala)[0];
    if (objetivo.libre) {
      base.sala = v.sala.Sala;
      base.advertencia = advertenciaCapacidad_(v.sala, asistentes);
      return respuesta_(true, 'DISPONIBLE',
        'La ' + v.sala.Sala + ' está libre el ' + horario + '.' + (base.advertencia ? ' ' + base.advertencia : ''), base);
    }
    const ocupante = objetivo.ocupada_por.map(o => o.hora_inicio + '–' + o.hora_fin + ' por ' + o.oficina).join('; ');
    const prefijo = (esRegistro ? 'Mientras tanto se registró otra reserva: l' : 'L') +
      'a ' + v.sala.Sala + ' está ocupada (' + ocupante + ').';
    if (libres.length) {
      base.sala_alternativa = libres[0];
      return respuesta_(false, 'CONFLICTO_SALA',
        prefijo + ' En ese mismo horario está libre: ' + libres.join(', ') + '.', base);
    }
    base.sugerencias = sugerencias_(v, cfg, salas, reservas);
    return respuesta_(false, 'CONFLICTO_AMBAS',
      prefijo + ' Ninguna otra sala está libre en ese horario.' + textoSugerencias_(base.sugerencias), base);
  }

  // Caso 2: "cualquier sala libre"
  if (libres.length) {
    base.sala = libres[0];
    const salaObj = salas.filter(s => s.Sala === libres[0])[0];
    base.advertencia = advertenciaCapacidad_(salaObj, asistentes);
    return respuesta_(true, 'DISPONIBLE',
      'Salas libres el ' + horario + ': ' + libres.join(', ') + '. Se asignará la ' + libres[0] + '.' + (base.advertencia ? ' ' + base.advertencia : ''), base);
  }
  base.sugerencias = sugerencias_(v, cfg, salas, reservas);
  return respuesta_(false, 'CONFLICTO_AMBAS',
    'Todas las salas están ocupadas el ' + horario + '.' + textoSugerencias_(base.sugerencias), base);
}

/** Reservas confirmadas de una sala que se traslapan con [ini, fin). */
function conflictos_(reservas, sala, ini, fin) {
  return reservas.filter(r => r.Sala === sala && hayTraslape_(ini, fin, aMin_(r.Hora_Inicio), aMin_(r.Hora_Fin)));
}

/** Conflicto = (InicioNueva < FinExistente) ∧ (FinNueva > InicioExistente). */
function hayTraslape_(ini1, fin1, ini2, fin2) {
  return ini1 < fin2 && fin1 > ini2;
}

/**
 * Horarios alternativos con la misma duración: primero el mismo día (más cercanos a la hora pedida),
 * y si no hay, los siguientes días hábiles. Máximo 4.
 */
function sugerencias_(v, cfg, salas, reservasDia) {
  const dur = v.fin - v.ini;
  const paso = cfg.PASO_MINUTOS;
  const apertura = aMin_(cfg.HORA_APERTURA), cierre = aMin_(cfg.HORA_CIERRE);
  const ahora = ahora_();
  const resultado = [];
  let fecha = v.fecha;
  for (let intento = 0; intento < 6 && resultado.length === 0; intento++) {
    if (intento > 0) {
      fecha = siguienteDiaHabil_(fecha, cfg);
      if (!fecha || diasEntre_(ahora.fecha, fecha) > cfg.ANTICIPACION_MAX_DIAS) break;
    }
    const reservas = intento === 0 ? reservasDia : reservasDelDia_(fecha);
    const candidatos = [];
    salas.forEach(s => {
      for (let t = apertura; t + dur <= cierre; t += paso) {
        if (fecha === ahora.fecha && t <= ahora.min) continue;
        if (conflictos_(reservas, s.Sala, t, t + dur).length === 0) {
          candidatos.push({ sala: s.Sala, fecha: fecha, hora_inicio: aHora_(t), hora_fin: aHora_(t + dur), _d: Math.abs(t - v.ini) });
        }
      }
    });
    candidatos.sort((a, b) => a._d - b._d || a.sala.localeCompare(b.sala));
    const vistos = {}; // evita repetir la misma hora en salas distintas
    candidatos.forEach(c => {
      if (resultado.length < 4 && !vistos[c.hora_inicio]) {
        vistos[c.hora_inicio] = true;
        delete c._d;
        resultado.push(c);
      }
    });
  }
  return resultado;
}

function textoSugerencias_(sug) {
  if (!sug || !sug.length) return ' No encontré horarios libres cercanos; prueba con otra fecha.';
  return ' Horarios libres de la misma duración: ' +
    sug.map(s => s.sala + ' ' + fechaCorta_(s.fecha) + ' ' + s.hora_inicio + '–' + s.hora_fin).join('; ') + '.';
}

function advertenciaCapacidad_(sala, asistentes) {
  const cap = entero_(sala && sala.Capacidad);
  if (asistentes && cap && asistentes > cap) {
    return 'Atención: la ' + sala.Sala + ' tiene capacidad para ' + cap + ' personas y se indicaron ' + asistentes + '.';
  }
  return '';
}

/** Cancela una reserva. actor = { Nombre, Correo, Rol }. */
function cancelarReserva_(idReserva, actor, motivo) {
  const id = String(idReserva || '').trim().toUpperCase();
  if (!id) return datosInvalidos_(['Indica el código de la reserva.']);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return respuesta_(false, 'ERROR', 'El sistema está ocupado. Intenta de nuevo en unos segundos.');
  let r;
  try {
    const t = leerTabla_(HOJAS.RESERVAS);
    r = t.filas.filter(x => String(x.ID_Reserva).toUpperCase() === id)[0];
    if (!r) return respuesta_(false, 'NO_ENCONTRADA', 'No encontré la reserva ' + id + '.');
    if (r.Estado === ESTADO.CANCELADA) return datosInvalidos_(['La reserva ' + id + ' ya estaba cancelada.']);
    const esDueno = !!actor.Correo && r.Correo_Contacto === actor.Correo;
    if (!esDueno && actor.Rol !== 'ADMIN') {
      return respuesta_(false, 'NO_PERMITIDO', 'Solo quien hizo la reserva o un administrador de la OGPP puede cancelarla.');
    }
    if (!puedeCancelar_(r)) {
      return datosInvalidos_(['La reserva ' + id + ' ya empezó o terminó; no se puede cancelar.']);
    }
    const col = nombre => t.enc.indexOf(nombre) + 1;
    const motivoLimpio = limpiarTexto_(motivo);
    t.hoja.getRange(r._fila, col('Estado')).setValue(ESTADO.CANCELADA);
    t.hoja.getRange(r._fila, col('Cancelado_Por')).setValue(actor.Nombre || actor.Correo || '');
    t.hoja.getRange(r._fila, col('Fecha_Cancelacion')).setValue(ahoraTexto_());
    t.hoja.getRange(r._fila, col('Motivo_Cancelacion')).setValue(motivoLimpio);
    SpreadsheetApp.flush();
    r.Estado = ESTADO.CANCELADA;
    r.Cancelado_Por = actor.Nombre || '';
    r.Motivo_Cancelacion = motivoLimpio;
  } finally {
    lock.releaseLock();
  }
  const envio = enviarNotificaciones_('CANCELACION', r);
  return respuesta_(true, 'CANCELADA', 'La reserva ' + r.ID_Reserva + ' fue cancelada.',
    { reserva: reservaVista_(r, actor), correo_usuario: envio.usuario, correo_encargados: envio.encargados });
}

/** Una reserva confirmada solo se cancela antes de su hora de inicio. */
function puedeCancelar_(r) {
  if (r.Estado !== ESTADO.CONFIRMADA) return false;
  const ahora = ahora_();
  return r.Fecha_Reserva > ahora.fecha || (r.Fecha_Reserva === ahora.fecha && aMin_(r.Hora_Inicio) > ahora.min);
}

/** ID correlativo por día de registro: RES-AAMMDD-NNN (se llama dentro del bloqueo). */
function generarId_() {
  const prefijo = 'RES-' + Utilities.formatDate(fechaActual_(), TZ, 'yyMMdd') + '-';
  let max = 0;
  leerTabla_(HOJAS.RESERVAS).filas.forEach(r => {
    const id = String(r.ID_Reserva);
    if (id.indexOf(prefijo) === 0) max = Math.max(max, parseInt(id.slice(prefijo.length), 10) || 0);
  });
  return prefijo + ('00' + (max + 1)).slice(-3);
}

/* ============================================================================
 *  4. IDENTIDAD
 * ========================================================================== */

function correoActivo_() {
  try { return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { return ''; }
}

/**
 * Usuario que está usando la página: lista blanca Usuarios (activo) o, si es administrador
 * por Config › ADMINS, un perfil de administración. null si no tiene acceso.
 */
function usuarioActual_() {
  const correo = correoActivo_();
  if (!correo) return null;
  const u = leerTabla_(HOJAS.USUARIOS).filas.filter(x => esSi_(x.Activo) && x.Correo === correo)[0];
  if (u) {
    return { Correo: correo, Nombre: u.Nombre || correo, Oficina: u.Oficina || '', Rol: String(u.Rol || 'USUARIO').toUpperCase() === 'ADMIN' ? 'ADMIN' : 'USUARIO' };
  }
  const extra = String(leerConfig_().ADMINS || '').toLowerCase().split(',').map(x => x.trim()).filter(Boolean);
  if (extra.indexOf(correo) !== -1) return { Correo: correo, Nombre: correo, Oficina: 'OGPP (administración)', Rol: 'ADMIN' };
  return null;
}

/* ============================================================================
 *  5. ACCESO A DATOS
 * ========================================================================== */

function libro_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function hoja_(nombre) {
  const h = libro_().getSheetByName(nombre);
  if (!h) throw new Error('Falta la pestaña "' + nombre + '". Ejecuta setup().');
  return h;
}

/** Lee una pestaña como objetos { encabezado: valor, _fila }. */
function leerTabla_(nombre) {
  const h = hoja_(nombre);
  const valores = h.getDataRange().getValues();
  const enc = (valores[0] || []).map(x => String(x).trim());
  const filas = [];
  for (let i = 1; i < valores.length; i++) {
    const fila = valores[i];
    if (fila.every(c => c === '' || c === null)) continue;
    const o = { _fila: i + 1 };
    enc.forEach((k, j) => { o[k] = normalizarCelda_(k, fila[j]); });
    filas.push(o);
  }
  return { hoja: h, enc: enc, filas: filas };
}

/** Convierte valores de Sheets a texto homogéneo según el tipo de columna. */
function normalizarCelda_(columna, v) {
  if (v === null || v === undefined) return '';
  const esHora = /^Hora_/.test(columna);
  const esFecha = columna === 'Fecha_Reserva' || columna === 'Fecha';
  if (v instanceof Date) {
    if (esHora) return Utilities.formatDate(v, TZ, 'HH:mm');
    if (esFecha) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  }
  if (esHora) return normHora_(v) || String(v).trim();
  if (columna === 'Fecha_Reserva') return normFecha_(v) || String(v).trim();
  if (columna === 'Correo' || columna === 'Correo_Contacto') return String(v).trim().toLowerCase();
  return typeof v === 'string' ? v.trim() : v;
}

function agregarFila_(nombre, obj) {
  const h = hoja_(nombre);
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(x => String(x).trim());
  h.appendRow(enc.map(k => (obj[k] === undefined ? '' : obj[k])));
}

function reservasDelDia_(fecha) {
  return leerTabla_(HOJAS.RESERVAS).filas.filter(r => r.Estado === ESTADO.CONFIRMADA && r.Fecha_Reserva === fecha);
}

function leerSalas_() {
  return leerTabla_(HOJAS.SALAS).filas.filter(s => s.Sala && esSi_(s.Activa));
}

function leerFeriados_() {
  const mapa = {};
  leerTabla_(HOJAS.FERIADOS).filas.forEach(f => {
    const fecha = normFecha_(f.Fecha);
    if (fecha) mapa[fecha] = f.Descripcion || 'feriado';
  });
  return mapa;
}

function leerConfig_() {
  const c = {};
  CONFIG_DEFECTO.forEach(r => { c[r[0]] = r[1]; });
  try {
    leerTabla_(HOJAS.CONFIG).filas.forEach(r => {
      if (r.Clave && r.Valor !== '') c[String(r.Clave).trim()] = r.Valor;
    });
  } catch (e) { /* usa valores por defecto */ }
  c.HORA_APERTURA = normHora_(c.HORA_APERTURA) || '08:00';
  c.HORA_CIERRE = normHora_(c.HORA_CIERRE) || '17:00';
  c.DIAS_HABILES = String(c.DIAS_HABILES).split(',').map(x => parseInt(x, 10)).filter(x => x >= 1 && x <= 7);
  ['DURACION_MIN', 'DURACION_MAX', 'ANTICIPACION_MAX_DIAS', 'PASO_MINUTOS'].forEach(k => { c[k] = parseInt(c[k], 10); });
  if (!(c.PASO_MINUTOS > 0)) c.PASO_MINUTOS = 30;
  return c;
}

function registrarLog_(accion, correo, codigo, detalle) {
  const h = libro_().getSheetByName(HOJAS.LOG);
  if (!h) return;
  h.appendRow([ahoraTexto_(), accion || '', correo || '', codigo || '', String(detalle || '').slice(0, 45000)]);
}

/* ============================================================================
 *  6. CORREOS (usuario + encargados, con invitación .ics)
 * ========================================================================== */

/**
 * Tras una reserva o cancelación avisa (1) a quien reservó y (2) a los encargados:
 * el Correo_Responsable de la sala y los correos de Config › CORREOS_NOTIFICACION.
 * Un fallo de correo nunca deshace la reserva.
 */
function enviarNotificaciones_(tipo, r) {
  const res = { usuario: false, encargados: false };
  let cfg, sala;
  try {
    cfg = leerConfig_();
    sala = leerSalas_().filter(s => s.Sala === r.Sala)[0] || {};
  } catch (e) { return res; }

  try {
    if (r.Correo_Contacto) {
      enviarCorreo_(armarCorreo_(tipo, r, sala, cfg, false), r.Correo_Contacto, '');
      res.usuario = true;
    }
  } catch (e) { logCorreoFallido_(tipo, r, 'usuario', e); }

  try {
    const encargados = encargados_(sala, cfg).filter(c => c !== r.Correo_Contacto);
    if (encargados.length) {
      enviarCorreo_(armarCorreo_(tipo, r, sala, cfg, true), encargados.join(','), '');
      res.encargados = true;
    }
  } catch (e) { logCorreoFallido_(tipo, r, 'encargados', e); }
  return res;
}

function encargados_(sala, cfg) {
  const lista = [];
  [sala && sala.Correo_Responsable, cfg.CORREOS_NOTIFICACION].forEach(s => {
    String(s || '').split(/[,;]/).map(x => x.trim().toLowerCase()).filter(Boolean).forEach(x => { if (lista.indexOf(x) === -1) lista.push(x); });
  });
  return lista;
}

function logCorreoFallido_(tipo, r, para, e) {
  try { registrarLog_('correo_' + tipo, r && r.Correo_Contacto, 'ERROR', para + ': ' + String(e)); } catch (x) { /* nada */ }
}

function armarCorreo_(tipo, r, sala, cfg, paraEncargados) {
  const titulos = {
    CONFIRMACION: { color: '#1a7f37', asunto: paraEncargados ? 'Nueva reserva' : 'Reserva confirmada', titulo: paraEncargados ? 'Nueva reserva registrada' : 'Reserva confirmada' },
    CANCELACION: { color: '#b42318', asunto: 'Reserva cancelada', titulo: 'Reserva cancelada' }
  };
  const t = titulos[tipo];
  const filas = [
    ['Código', r.ID_Reserva], ['Sala', r.Sala + (sala.Ubicacion ? ' – ' + sala.Ubicacion : '')],
    ['Fecha', fechaLarga_(r.Fecha_Reserva)], ['Horario', r.Hora_Inicio + ' – ' + r.Hora_Fin + ' h'],
    ['Oficina', r.Oficina_Solicitante], ['Responsable', r.Responsable + ' (' + r.Correo_Contacto + ')'], ['Tema', r.Tema_Reunion]
  ];
  if (r.Asistentes) filas.push(['Asistentes', r.Asistentes]);
  if (tipo === 'CANCELACION') {
    filas.push(['Cancelada por', r.Cancelado_Por || '—']);
    if (r.Motivo_Cancelacion) filas.push(['Motivo', r.Motivo_Cancelacion]);
  }
  const pie = tipo === 'CONFIRMACION' && !paraEncargados
    ? '<p style="font-size:13px;color:#555">Puedes cancelarla desde la página de reservas, en la pestaña “Mis reservas”.</p>' : '';
  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;margin:auto;border:1px solid #e5e5e5;border-radius:8px;overflow:hidden">' +
    '<div style="background:' + t.color + ';color:#fff;padding:16px 24px;font-size:18px;font-weight:bold">' + t.titulo + '</div>' +
    '<div style="padding:16px 24px"><table style="border-collapse:collapse;width:100%;font-size:14px">' +
    filas.map(f => '<tr><td style="padding:6px 0;color:#666;width:120px;vertical-align:top">' + f[0] + '</td><td style="padding:6px 0"><b>' +
      escaparHtml_(String(f[1] || '')) + '</b></td></tr>').join('') +
    '</table>' + pie + '</div><div style="background:#f6f6f6;padding:10px 24px;font-size:12px;color:#777">' +
    escaparHtml_(cfg.NOMBRE_SERVICIO) + ' · Mensaje automático</div></div>';
  return {
    subject: '[' + t.asunto + '] ' + r.Sala + ' · ' + fechaCorta_(r.Fecha_Reserva) + ' ' + r.Hora_Inicio + ' · ' + r.ID_Reserva,
    htmlBody: html,
    name: cfg.NOMBRE_SERVICIO,
    ics: generarIcs_(r, tipo === 'CANCELACION')
  };
}

function enviarCorreo_(c, para) {
  MailApp.sendEmail({
    to: para, subject: c.subject, htmlBody: c.htmlBody, name: c.name,
    attachments: [Utilities.newBlob(c.ics, 'text/calendar', 'reserva.ics')]
  });
}

function generarIcs_(r, cancelar) {
  const utc = (fecha, hora) => Utilities.formatDate(Utilities.parseDate(fecha + ' ' + hora, TZ, 'yyyy-MM-dd HH:mm'), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const esc = s => String(s || '').replace(/([,;\\])/g, '\\$1').replace(/\n/g, '\\n');
  return [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//OGPP UNMSM//Reservas//ES',
    'METHOD:' + (cancelar ? 'CANCEL' : 'PUBLISH'),
    'BEGIN:VEVENT',
    'UID:' + r.ID_Reserva + '@ogpp-unmsm',
    'SEQUENCE:' + (cancelar ? 1 : 0),
    'DTSTAMP:' + Utilities.formatDate(fechaActual_(), 'UTC', "yyyyMMdd'T'HHmmss'Z'"),
    'DTSTART:' + utc(r.Fecha_Reserva, r.Hora_Inicio),
    'DTEND:' + utc(r.Fecha_Reserva, r.Hora_Fin),
    'SUMMARY:' + esc(r.Tema_Reunion + ' (' + r.Sala + ')'),
    'LOCATION:' + esc(r.Sala),
    'DESCRIPTION:' + esc('Reserva ' + r.ID_Reserva + ' – ' + r.Oficina_Solicitante),
    'STATUS:' + (cancelar ? 'CANCELLED' : 'CONFIRMED'),
    'END:VEVENT', 'END:VCALENDAR'
  ].join('\r\n');
}

/* ============================================================================
 *  7. CONFIGURACIÓN INICIAL Y PRUEBAS
 * ========================================================================== */

/** Crea o repara todas las pestañas. Seguro de ejecutar varias veces. */
function setup() {
  const libro = libro_();
  libro.setSpreadsheetTimeZone(TZ);
  Object.keys(ENCABEZADOS).forEach(nombre => {
    let h = libro.getSheetByName(nombre);
    if (!h) h = libro.insertSheet(nombre);
    const enc = ENCABEZADOS[nombre];
    if (h.getMaxColumns() < enc.length) h.insertColumnsAfter(h.getMaxColumns(), enc.length - h.getMaxColumns());
    const actuales = h.getRange(1, 1, 1, enc.length).getValues()[0];
    if (actuales.join('') === '') h.getRange(1, 1, 1, enc.length).setValues([enc]);
    h.getRange(1, 1, 1, enc.length).setFontWeight('bold').setBackground('#e8eef7');
    h.setFrozenRows(1);
    // Todo como texto plano: Sheets no convierte fechas ni horas.
    h.getRange(1, 1, h.getMaxRows(), enc.length).setNumberFormat('@');
  });

  const cfgHoja = libro.getSheetByName(HOJAS.CONFIG);
  const claves = cfgHoja.getDataRange().getValues().map(r => String(r[0]));
  CONFIG_DEFECTO.forEach(r => { if (claves.indexOf(r[0]) === -1) cfgHoja.appendRow(r); });

  const fer = libro.getSheetByName(HOJAS.FERIADOS);
  if (fer.getLastRow() < 2) FERIADOS_DEFECTO.forEach(r => fer.appendRow(r));

  const salas = libro.getSheetByName(HOJAS.SALAS);
  if (salas.getLastRow() < 2) {
    salas.appendRow(['Sala 1', '12', 'Por definir', 'Proyector, TV', '', 'SI']);
    salas.appendRow(['Sala 2', '8', 'Por definir', 'TV', '', 'SI']);
  }

  const usuarios = libro.getSheetByName(HOJAS.USUARIOS);
  if (usuarios.getLastRow() < 2) {
    usuarios.appendRow([String(Session.getEffectiveUser().getEmail()).toLowerCase(), 'Administrador inicial', 'OGPP', 'ADMIN', 'SI']);
  }
  Logger.log('Setup completo. Completa las pestañas Usuarios, Salas y Config, y publica la aplicación web.');
}

/** Prueba rápida desde el editor con tu propia cuenta. Revisa el registro de ejecución. */
function probar() {
  const ini = apiInicio();
  Logger.log('inicio → ' + ini.codigo + ' · ' + ini.mensaje);
  if (!ini.ok) return;
  const cfg = leerConfig_();
  const fecha = siguienteDiaHabil_(ahora_().fecha, cfg);
  const ver = apiVerificar({ sala: '', fecha: fecha, hora_inicio: '10:00', hora_fin: '11:00' });
  Logger.log('verificar ' + fecha + ' 10:00–11:00 → ' + ver.codigo + ' · ' + ver.mensaje);
  Logger.log('mis reservas → ' + JSON.stringify(apiMisReservas().data));
}

/* ============================================================================
 *  8. UTILIDADES
 * ========================================================================== */

function respuesta_(ok, codigo, mensaje, data) {
  return { ok: ok, codigo: codigo, mensaje: mensaje, data: data || {} };
}

function datosInvalidos_(errores) {
  return respuesta_(false, 'DATOS_INVALIDOS', errores.join(' '), { errores: errores });
}

/**
 * Vista de una reserva para la página. Los datos personales (responsable, correo) y la
 * opción de cancelar solo se incluyen para su dueño y para administradores.
 */
function reservaVista_(r, usuario) {
  const o = {
    id_reserva: r.ID_Reserva, sala: r.Sala, fecha: r.Fecha_Reserva, hora_inicio: r.Hora_Inicio, hora_fin: r.Hora_Fin,
    oficina: r.Oficina_Solicitante, tema: r.Tema_Reunion, estado: r.Estado,
    propia: !!usuario && !!usuario.Correo && r.Correo_Contacto === usuario.Correo
  };
  if (o.propia || (usuario && usuario.Rol === 'ADMIN')) {
    o.responsable = r.Responsable;
    o.correo = r.Correo_Contacto;
    o.asistentes = r.Asistentes;
    o.puede_cancelar = puedeCancelar_(r);
    if (r.Estado === ESTADO.CANCELADA) { o.cancelado_por = r.Cancelado_Por; o.motivo_cancelacion = r.Motivo_Cancelacion; }
  }
  return o;
}

/** Punto único para obtener la fecha actual (facilita pruebas). */
function fechaActual_() {
  return new Date();
}

/** { fecha: 'yyyy-MM-dd', min: minutos desde 00:00 } en hora de Lima. */
function ahora_() {
  const s = Utilities.formatDate(fechaActual_(), TZ, 'yyyy-MM-dd HH:mm');
  return { fecha: s.slice(0, 10), min: aMin_(s.slice(11, 16)) };
}

function ahoraTexto_() {
  return Utilities.formatDate(fechaActual_(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function normFecha_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) {
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // dd/mm/aaaa
    if (m) m = [m[0], m[3], m[2], m[1]];
  }
  if (!m) return '';
  const y = +m[1], mo = +m[2], d = +m[3];
  const dt = new Date(Date.UTC(y, mo - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return '';
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
}

function normHora_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  const s = String(v || '').trim().toLowerCase().replace(/\s+/g, '')
    .replace(/(hrs?|horas)$/, '').replace(/^(\d{1,2})h(\d{2})/, '$1:$2');
  const m = s.match(/^(\d{1,2})(?::?(\d{2}))?(am|pm|a\.m\.|p\.m\.)?$/);
  if (!m) return '';
  let h = +m[1];
  const mi = m[2] ? +m[2] : 0;
  const suf = m[3] ? m[3].replace(/\./g, '') : '';
  if (suf === 'pm' && h < 12) h += 12;
  if (suf === 'am' && h === 12) h = 0;
  if (h > 23 || mi > 59) return '';
  return ('0' + h).slice(-2) + ':' + ('0' + mi).slice(-2);
}

function resolverSala_(entrada, salas) {
  const clave = s => String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, '');
  const k = clave(entrada);
  return salas.filter(s => clave(s.Sala) === k || clave(s.Sala) === 'sala' + k)[0] || null;
}

function aMin_(hhmm) {
  const p = String(hhmm).split(':');
  return (+p[0]) * 60 + (+p[1] || 0);
}

function aHora_(min) {
  return ('0' + Math.floor(min / 60)).slice(-2) + ':' + ('0' + (min % 60)).slice(-2);
}

/** 1 = lunes ... 7 = domingo */
function diaSemana_(fecha) {
  const d = new Date(fecha + 'T12:00:00Z').getUTCDay();
  return d === 0 ? 7 : d;
}

function sumarDias_(fecha, n) {
  const d = new Date(fecha + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

function diasEntre_(a, b) {
  return Math.round((new Date(b + 'T12:00:00Z') - new Date(a + 'T12:00:00Z')) / 86400000);
}

function siguienteDiaHabil_(fecha, cfg) {
  let f = fecha;
  for (let i = 0; i < 30; i++) {
    f = sumarDias_(f, 1);
    if (!motivoDiaNoHabil_(f, cfg)) return f;
  }
  return '';
}

function fechaLarga_(fecha) {
  const p = fecha.split('-');
  return DIAS[diaSemana_(fecha)] + ' ' + p[2] + '/' + p[1] + '/' + p[0];
}

function fechaCorta_(fecha) {
  const p = fecha.split('-');
  return p[2] + '/' + p[1];
}

function duracionTexto_(min) {
  const h = Math.floor(min / 60), m = min % 60;
  return (h ? h + ' h' : '') + (h && m ? ' ' : '') + (m ? m + ' min' : '');
}

function entero_(v) {
  const n = parseInt(v, 10);
  return isNaN(n) ? 0 : n;
}

function esSi_(v) {
  return ['si', 'sí', 'true', '1', 'x', 'activo', 'activa'].indexOf(String(v).trim().toLowerCase()) !== -1 || v === true;
}

function limpiarTexto_(v) {
  return String(v || '').replace(/\s+/g, ' ').trim().slice(0, 200);
}

function escaparHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
