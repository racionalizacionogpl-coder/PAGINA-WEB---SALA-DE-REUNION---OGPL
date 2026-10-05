/**
 * ============================================================================
 *  Reserva de Salas – OGPL UNMSM  ·  Aplicación web (Apps Script + Google Sheets)
 * ----------------------------------------------------------------------------
 *  Acceso con correo y contraseña propios de la aplicación (no depende de la
 *  cuenta Google del visitante). La página (Index.html) llama a las funciones
 *  api* con google.script.run, enviando el token de sesión que entrega apiLogin.
 *
 *  Primer uso: ejecutar setup() una vez desde el editor (muestra la contraseña
 *  temporal del administrador en el registro de ejecución) y publicar la web app.
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
  Usuarios: ['Correo', 'Nombre', 'Oficina', 'Rol', 'Activo', 'Hash_Contrasena', 'Sal', 'Debe_Cambiar', 'Fecha_Creacion', 'Ultimo_Acceso'],
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
  ['NOMBRE_SERVICIO', 'Reserva de Salas – OGPL UNMSM', 'Nombre que aparece en la página y como remitente de los correos.'],
  ['NOMBRE_OFICINA', 'Oficina General de Planificación – UNMSM', 'Texto del encabezado de los correos, junto al logo.'],
  ['MINUTOS_SESION', '120', 'Minutos sin actividad tras los cuales se cierra la sesión (máximo 360).'],
  ['PERMITIR_SOLICITUDES', 'SI', 'SI = en la pantalla de ingreso aparece "Solicitar una cuenta" (un administrador debe aprobarla). NO = solo los administradores crean usuarios.'],
  ['DOMINIOS_PERMITIDOS', '', 'Dominios de correo aceptados en las solicitudes de cuenta, separados por coma (p. ej. unmsm.edu.pe). Vacío = cualquiera.']
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
const ITERACIONES_HASH = 300;
const MAX_INTENTOS = 5;          // intentos fallidos antes de bloquear
const MINUTOS_BLOQUEO = 15;
const COLOR_MARCA = '#7a1f2b';

/* ============================================================================
 *  1. PUNTOS DE ENTRADA
 * ========================================================================== */

/** Sirve la página. El acceso se controla con usuario y contraseña dentro de la página. */
function doGet() {
  const cfg = leerConfig_();
  const t = HtmlService.createTemplateFromFile('Index');
  t.logo = 'data:image/png;base64,' + LOGO_PNG_BASE64;
  t.servicio = cfg.NOMBRE_SERVICIO;
  t.oficina = cfg.NOMBRE_OFICINA;
  t.permitirSolicitudes = esSi_(cfg.PERMITIR_SOLICITUDES) ? 'SI' : 'NO';
  return t.evaluate()
    .setTitle(cfg.NOMBRE_SERVICIO)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Menú en la hoja de cálculo. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Reservas OGPL')
    .addItem('Configurar / reparar hojas', 'setup')
    .addItem('Asignar contraseña a un usuario…', 'asignarContrasena')
    .addToUi();
}

/* ============================================================================
 *  2. SESIONES Y CONTRASEÑAS
 * ========================================================================== */

/** Inicia sesión. Devuelve un token que la página envía en cada llamada. */
function apiLogin(correo, contrasena) {
  const c = normCorreo_(correo);
  let res;
  try {
    const cache = CacheService.getScriptCache();
    const claveFallos = 'fallos_' + c;
    const fallos = entero_(cache.get(claveFallos));
    if (!c || !contrasena) {
      res = respuesta_(false, 'CREDENCIALES', 'Ingresa tu correo y tu contraseña.');
    } else if (fallos >= MAX_INTENTOS) {
      res = respuesta_(false, 'BLOQUEADO', 'Demasiados intentos fallidos. Espera ' + MINUTOS_BLOQUEO + ' minutos e intenta de nuevo.');
    } else {
      const u = buscarUsuario_(c);
      if (!u || !u.Hash_Contrasena || hashContrasena_(String(contrasena), u.Sal) !== u.Hash_Contrasena) {
        cache.put(claveFallos, String(fallos + 1), MINUTOS_BLOQUEO * 60);
        res = respuesta_(false, 'CREDENCIALES', 'Correo o contraseña incorrectos.');
      } else if (u.Estado === 'PENDIENTE') {
        res = respuesta_(false, 'PENDIENTE', 'Tu solicitud de cuenta aún no ha sido aprobada por un administrador.');
      } else if (u.Estado !== 'SI') {
        res = respuesta_(false, 'DESACTIVADO', 'Tu cuenta está desactivada. Comunícate con la administración.');
      } else {
        cache.remove(claveFallos);
        actualizarFila_(HOJAS.USUARIOS, u._fila, { Ultimo_Acceso: ahoraTexto_() });
        res = respuesta_(true, 'EXITO', 'Bienvenido(a), ' + u.Nombre + '.', {
          token: crearSesion_(u), debe_cambiar: u.DebeCambiar, usuario: perfil_(u)
        });
      }
    }
  } catch (err) {
    res = errorInterno_(err);
  }
  try { registrarLog_('login', c, res.codigo, res.ok ? '' : res.mensaje); } catch (ignorado) { /* nada */ }
  return res;
}

/** Cierra la sesión. */
function apiLogout(token) {
  try { CacheService.getScriptCache().remove('ses_' + String(token || '')); } catch (e) { /* nada */ }
  return respuesta_(true, 'EXITO', 'Sesión cerrada.');
}

/** Cambia la contraseña propia. Devuelve un token nuevo (las demás sesiones abiertas se cierran). */
function apiCambiarContrasena(token, actual, nueva) {
  return api_('cambiar_contrasena', token, usuario => {
    const u = buscarUsuario_(usuario.Correo);
    if (hashContrasena_(String(actual || ''), u.Sal) !== u.Hash_Contrasena) {
      return respuesta_(false, 'CREDENCIALES', 'La contraseña actual no es correcta.');
    }
    const error = validarContrasena_(nueva);
    if (error) return datosInvalidos_([error]);
    if (String(nueva) === String(actual)) return datosInvalidos_(['La nueva contraseña debe ser distinta de la actual.']);
    const sal = nuevaSal_();
    actualizarFila_(HOJAS.USUARIOS, u._fila, { Hash_Contrasena: hashContrasena_(String(nueva), sal), Sal: sal, Debe_Cambiar: 'NO' });
    CacheService.getScriptCache().remove('ses_' + token);
    u.Sal = sal;
    return respuesta_(true, 'EXITO', 'Tu contraseña fue actualizada.', { token: crearSesion_(u) });
  }, true);
}

/**
 * Solicitud pública de cuenta (pantalla de ingreso). Queda PENDIENTE hasta que un ADMIN la apruebe.
 * datos = { correo, nombre, oficina, contrasena }
 */
function apiSolicitarCuenta(datos) {
  datos = datos || {};
  const c = normCorreo_(datos.correo);
  let res;
  try {
    const cfg = leerConfig_();
    const cache = CacheService.getScriptCache();
    const enHora = entero_(cache.get('solicitudes_hora'));
    const errores = [];
    if (!esSi_(cfg.PERMITIR_SOLICITUDES)) {
      res = respuesta_(false, 'NO_PERMITIDO', 'Las solicitudes de cuenta están deshabilitadas. Pide a un administrador que te cree una cuenta.');
    } else if (enHora >= 20) {
      res = respuesta_(false, 'ERROR', 'Se recibieron demasiadas solicitudes. Intenta de nuevo en una hora.');
    } else {
      validarDatosUsuario_(datos, errores);
      const dominios = String(cfg.DOMINIOS_PERMITIDOS || '').toLowerCase().split(',').map(x => x.trim().replace(/^@/, '')).filter(Boolean);
      if (c && dominios.length && dominios.indexOf(c.split('@')[1]) === -1) {
        errores.push('Solo se aceptan correos de: ' + dominios.map(d => '@' + d).join(', ') + '.');
      }
      const errorPw = validarContrasena_(datos.contrasena);
      if (errorPw) errores.push(errorPw);
      if (errores.length) {
        res = datosInvalidos_(errores);
      } else {
        const lock = LockService.getScriptLock();
        if (!lock.tryLock(20000)) return respuesta_(false, 'ERROR', 'El sistema está ocupado. Intenta de nuevo.');
        try {
          const existente = buscarUsuario_(c);
          if (existente) {
            res = respuesta_(false, 'YA_EXISTE', existente.Estado === 'PENDIENTE'
              ? 'Ya existe una solicitud pendiente para este correo.'
              : 'Ya existe una cuenta con este correo. Si olvidaste tu contraseña, pide a un administrador que la restablezca.');
          } else {
            const sal = nuevaSal_();
            agregarFila_(HOJAS.USUARIOS, {
              Correo: c, Nombre: limpiarTexto_(datos.nombre), Oficina: limpiarTexto_(datos.oficina), Rol: 'USUARIO',
              Activo: 'PENDIENTE', Hash_Contrasena: hashContrasena_(String(datos.contrasena), sal), Sal: sal,
              Debe_Cambiar: 'NO', Fecha_Creacion: ahoraTexto_(), Ultimo_Acceso: ''
            });
            SpreadsheetApp.flush();
            cache.put('solicitudes_hora', String(enHora + 1), 3600);
            res = respuesta_(true, 'EXITO', 'Solicitud enviada. Podrás ingresar cuando un administrador la apruebe; te avisaremos por correo.');
          }
        } finally {
          lock.releaseLock();
        }
        if (res.ok) avisarSolicitudAdmins_({ Correo: c, Nombre: limpiarTexto_(datos.nombre), Oficina: limpiarTexto_(datos.oficina) });
      }
    }
  } catch (err) {
    res = errorInterno_(err);
  }
  try { registrarLog_('solicitar_cuenta', c, res.codigo, res.ok ? '' : res.mensaje); } catch (ignorado) { /* nada */ }
  return res;
}

function crearSesion_(u) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  CacheService.getScriptCache().put('ses_' + token, JSON.stringify({ c: u.Correo, s: u.Sal }), segundosSesion_());
  return token;
}

function segundosSesion_() {
  const m = entero_(leerConfig_().MINUTOS_SESION) || 120;
  return Math.min(21600, Math.max(300, m * 60));
}

/** Usuario dueño de un token válido (o null). Renueva la vigencia en cada uso. */
function usuarioDeToken_(token) {
  if (typeof token !== 'string' || token.length < 32) return null;
  const cache = CacheService.getScriptCache();
  const clave = 'ses_' + token;
  const valor = cache.get(clave);
  if (!valor) return null;
  let d;
  try { d = JSON.parse(valor); } catch (e) { return null; }
  const u = buscarUsuario_(d.c);
  // La sesión muere si la cuenta se desactiva o si cambió la contraseña (cambia la sal)
  if (!u || u.Estado !== 'SI' || u.Sal !== d.s) { cache.remove(clave); return null; }
  cache.put(clave, valor, segundosSesion_());
  return u;
}

function hashContrasena_(contrasena, sal) {
  let h = String(sal) + '|' + contrasena;
  for (let i = 0; i < ITERACIONES_HASH; i++) {
    h = bytesAHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + '|' + sal, Utilities.Charset.UTF_8));
  }
  return h;
}

function bytesAHex_(bytes) {
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function nuevaSal_() {
  return Utilities.getUuid().replace(/-/g, '');
}

/** Devuelve el motivo si la contraseña no cumple la política, o ''. */
function validarContrasena_(pw) {
  const s = String(pw || '');
  if (s.length < 8) return 'La contraseña debe tener al menos 8 caracteres.';
  if (!/[A-Za-zÁÉÍÓÚáéíóúÑñ]/.test(s) || !/\d/.test(s)) return 'La contraseña debe combinar letras y números.';
  if (s.length > 100) return 'La contraseña es demasiado larga.';
  return '';
}

/** Contraseña temporal legible (10 caracteres, letras y números, sin caracteres ambiguos). */
function contrasenaTemporal_() {
  const letras = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ', nums = '23456789';
  const fuente = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  let pw = '';
  for (let i = 0; i < 10; i++) {
    const n = parseInt(fuente.substr(i * 2, 2), 16);
    pw += (i === 3 || i === 7) ? nums.charAt(n % nums.length) : letras.charAt(n % letras.length);
  }
  return pw;
}

/* ============================================================================
 *  3. API PARA LA PÁGINA (requieren sesión)
 *  Todas devuelven { ok, codigo, mensaje, data } y nunca lanzan excepciones.
 * ========================================================================== */

/** Ejecuta fn(usuario) si el token es válido. permitirSiDebeCambiar: solo para cambiar contraseña. */
function api_(nombre, token, fn, permitirSiDebeCambiar) {
  let usuario = null;
  let res;
  try {
    usuario = usuarioDeToken_(token);
    if (!usuario) {
      res = respuesta_(false, 'SESION_EXPIRADA', 'Tu sesión expiró. Ingresa nuevamente.');
    } else if (usuario.DebeCambiar && !permitirSiDebeCambiar) {
      res = respuesta_(false, 'DEBE_CAMBIAR', 'Debes cambiar tu contraseña antes de continuar.');
    } else {
      res = fn(usuario);
    }
  } catch (err) {
    res = errorInterno_(err);
  }
  try { registrarLog_(nombre, usuario && usuario.Correo, res.codigo, res.ok ? '' : res.mensaje); } catch (ignorado) { /* el log nunca rompe la respuesta */ }
  return res;
}

function soloAdmin_(usuario) {
  return usuario.Rol === 'ADMIN' ? null : respuesta_(false, 'NO_PERMITIDO', 'Esta sección es solo para administradores.');
}

/** Datos para armar la página: usuario, salas, reglas y feriados próximos. */
function apiInicio(token) {
  return api_('inicio', token, usuario => {
    const cfg = leerConfig_();
    const ahora = ahora_();
    const limite = sumarDias_(ahora.fecha, cfg.ANTICIPACION_MAX_DIAS);
    const feriados = leerFeriados_();
    return respuesta_(true, 'EXITO', 'Listo.', {
      usuario: perfil_(usuario),
      servicio: cfg.NOMBRE_SERVICIO,
      salas: leerSalas_().map(s => ({ sala: s.Sala, capacidad: entero_(s.Capacidad), ubicacion: s.Ubicacion, equipamiento: s.Equipamiento })),
      cfg: {
        apertura: cfg.HORA_APERTURA, cierre: cfg.HORA_CIERRE, dias_habiles: cfg.DIAS_HABILES,
        duracion_min: cfg.DURACION_MIN, duracion_max: cfg.DURACION_MAX,
        anticipacion_max_dias: cfg.ANTICIPACION_MAX_DIAS, paso: cfg.PASO_MINUTOS
      },
      hoy: ahora.fecha, ahora_min: ahora.min, limite: limite,
      feriados: Object.keys(feriados).filter(f => f >= ahora.fecha && f <= limite).sort()
        .map(f => ({ fecha: f, descripcion: feriados[f] })),
      pendientes: usuario.Rol === 'ADMIN' ? leerUsuarios_().filter(u => u.Estado === 'PENDIENTE').length : 0
    });
  });
}

/** Verifica una solicitud SIN registrarla (la página la llama mientras el usuario llena el formulario). */
function apiVerificar(token, form) {
  return api_('verificar', token, () => {
    form = form || {};
    const cfg = leerConfig_();
    const salas = leerSalas_();
    const v = validar_(form, cfg, salas);
    if (v.errores.length) return datosInvalidos_(v.errores);
    return evaluarSolicitud_(v, form, cfg, salas, reservasDelDia_(v.fecha), false);
  });
}

/** Registra la reserva (revalida dentro de un bloqueo para evitar dobles reservas). */
function apiReservar(token, form) {
  return api_('reservar', token, usuario => {
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
function apiMisReservas(token) {
  return api_('mis_reservas', token, usuario => {
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
function apiCancelar(token, idReserva, motivo) {
  return api_('cancelar', token, usuario => cancelarReserva_(idReserva, usuario, motivo));
}

/** Ocupación de la semana (según días hábiles) para el calendario. */
function apiSemana(token, fechaReferencia) {
  return api_('semana', token, usuario => {
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
 * Listado de reservas para administración. Filtros: { desde, hasta, estado: 'TODAS'|'CONFIRMADA'|'CANCELADA', sala }.
 */
function apiAdminLista(token, filtros) {
  return api_('admin_lista', token, usuario => {
    const denegado = soloAdmin_(usuario);
    if (denegado) return denegado;
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

/* ---------------- Administración de usuarios ---------------- */

/** Lista de usuarios (pendientes primero). */
function apiAdminUsuarios(token) {
  return api_('admin_usuarios', token, usuario => {
    const denegado = soloAdmin_(usuario);
    if (denegado) return denegado;
    const orden = { PENDIENTE: 0, SI: 1, NO: 2 };
    const lista = leerUsuarios_()
      .sort((a, b) => (orden[a.Estado] - orden[b.Estado]) || String(a.Nombre).localeCompare(String(b.Nombre)))
      .map(u => ({
        correo: u.Correo, nombre: u.Nombre, oficina: u.Oficina, rol: u.Rol, estado: u.Estado,
        debe_cambiar: u.DebeCambiar, tiene_contrasena: !!u.Hash_Contrasena,
        fecha_creacion: u.Fecha_Creacion, ultimo_acceso: u.Ultimo_Acceso, soy_yo: u.Correo === usuario.Correo
      }));
    return respuesta_(true, 'EXITO', '', { usuarios: lista });
  });
}

/**
 * Crea un usuario activo con contraseña temporal (deberá cambiarla al ingresar).
 * d = { correo, nombre, oficina, rol, contrasena, enviar_correo }
 */
function apiAdminCrearUsuario(token, d) {
  return api_('admin_crear_usuario', token, usuario => {
    const denegado = soloAdmin_(usuario);
    if (denegado) return denegado;
    d = d || {};
    const errores = [];
    validarDatosUsuario_(d, errores);
    const errorPw = validarContrasena_(d.contrasena);
    if (errorPw) errores.push(errorPw);
    const rol = String(d.rol || 'USUARIO').toUpperCase() === 'ADMIN' ? 'ADMIN' : 'USUARIO';
    if (errores.length) return datosInvalidos_(errores);
    const c = normCorreo_(d.correo);

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return respuesta_(false, 'ERROR', 'El sistema está ocupado. Intenta de nuevo.');
    try {
      if (buscarUsuario_(c)) return respuesta_(false, 'YA_EXISTE', 'Ya existe un usuario con el correo ' + c + '.');
      const sal = nuevaSal_();
      agregarFila_(HOJAS.USUARIOS, {
        Correo: c, Nombre: limpiarTexto_(d.nombre), Oficina: limpiarTexto_(d.oficina), Rol: rol, Activo: 'SI',
        Hash_Contrasena: hashContrasena_(String(d.contrasena), sal), Sal: sal, Debe_Cambiar: 'SI',
        Fecha_Creacion: ahoraTexto_(), Ultimo_Acceso: ''
      });
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    let enviado = false;
    if (d.enviar_correo) enviado = enviarCredenciales_({ Correo: c, Nombre: limpiarTexto_(d.nombre) }, String(d.contrasena), false);
    return respuesta_(true, 'EXITO', 'Usuario ' + c + ' creado.' + (d.enviar_correo ? (enviado ? ' Se le enviaron sus datos de acceso por correo.' : ' No se pudo enviar el correo; comunícale la contraseña temporal.') : ''),
      { correo_enviado: enviado });
  });
}

/**
 * Modifica nombre, oficina, rol o estado (SI / NO / PENDIENTE) de un usuario.
 * Aprobar una solicitud = estado SI (se avisa al usuario por correo).
 */
function apiAdminActualizarUsuario(token, correo, cambios) {
  return api_('admin_actualizar_usuario', token, usuario => {
    const denegado = soloAdmin_(usuario);
    if (denegado) return denegado;
    cambios = cambios || {};
    const u = buscarUsuario_(correo);
    if (!u) return respuesta_(false, 'NO_ENCONTRADA', 'No existe el usuario ' + correo + '.');
    const campos = {};
    const errores = [];
    if (cambios.nombre !== undefined) {
      if (limpiarTexto_(cambios.nombre).length < 3) errores.push('Escribe el nombre completo.');
      else campos.Nombre = limpiarTexto_(cambios.nombre);
    }
    if (cambios.oficina !== undefined) {
      if (limpiarTexto_(cambios.oficina).length < 2) errores.push('Escribe la oficina.');
      else campos.Oficina = limpiarTexto_(cambios.oficina);
    }
    if (cambios.rol !== undefined) campos.Rol = String(cambios.rol).toUpperCase() === 'ADMIN' ? 'ADMIN' : 'USUARIO';
    if (cambios.estado !== undefined) {
      const e = String(cambios.estado).toUpperCase();
      campos.Activo = e === 'SI' || e === 'PENDIENTE' ? e : 'NO';
    }
    if (u.Correo === usuario.Correo && ((campos.Rol && campos.Rol !== 'ADMIN') || (campos.Activo && campos.Activo !== 'SI'))) {
      errores.push('No puedes quitarte el rol de administrador ni desactivar tu propia cuenta.');
    }
    if (errores.length) return datosInvalidos_(errores);
    actualizarFila_(HOJAS.USUARIOS, u._fila, campos);
    SpreadsheetApp.flush();
    const aprobado = u.Estado === 'PENDIENTE' && campos.Activo === 'SI';
    if (aprobado) enviarAprobacion_(Object.assign({}, u, { Nombre: campos.Nombre || u.Nombre }));
    return respuesta_(true, 'EXITO', aprobado ? 'Cuenta de ' + u.Correo + ' aprobada; se le avisó por correo.' : 'Cambios guardados.');
  });
}

/** Asigna una contraseña temporal a otro usuario (deberá cambiarla al ingresar). */
function apiAdminRestablecerContrasena(token, correo, nueva, enviarCorreo) {
  return api_('admin_restablecer', token, usuario => {
    const denegado = soloAdmin_(usuario);
    if (denegado) return denegado;
    const u = buscarUsuario_(correo);
    if (!u) return respuesta_(false, 'NO_ENCONTRADA', 'No existe el usuario ' + correo + '.');
    if (u.Correo === usuario.Correo) return datosInvalidos_(['Para tu propia cuenta usa "Cambiar contraseña".']);
    const error = validarContrasena_(nueva);
    if (error) return datosInvalidos_([error]);
    const sal = nuevaSal_();
    actualizarFila_(HOJAS.USUARIOS, u._fila, { Hash_Contrasena: hashContrasena_(String(nueva), sal), Sal: sal, Debe_Cambiar: 'SI' });
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove('fallos_' + u.Correo);
    let enviado = false;
    if (enviarCorreo) enviado = enviarCredenciales_(u, String(nueva), true);
    return respuesta_(true, 'EXITO', 'Contraseña restablecida para ' + u.Correo + '.' + (enviarCorreo ? (enviado ? ' Se le envió por correo.' : ' No se pudo enviar el correo; comunícasela.') : ''),
      { correo_enviado: enviado });
  });
}

function validarDatosUsuario_(d, errores) {
  const c = normCorreo_(d.correo);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(c)) errores.push('Escribe un correo válido.');
  if (limpiarTexto_(d.nombre).length < 3) errores.push('Escribe el nombre completo.');
  if (limpiarTexto_(d.oficina).length < 2) errores.push('Escribe la unidad u oficina.');
}

/* ============================================================================
 *  4. LÓGICA DE RESERVAS
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

/** Horarios alternativos de la misma duración: mismo día primero, luego días hábiles siguientes. Máximo 4. */
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
    const vistos = {};
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

/** Cancela una reserva. actor = usuario { Nombre, Correo, Rol }. */
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
      return respuesta_(false, 'NO_PERMITIDO', 'Solo quien hizo la reserva o un administrador puede cancelarla.');
    }
    if (!puedeCancelar_(r)) {
      return datosInvalidos_(['La reserva ' + id + ' ya empezó o terminó; no se puede cancelar.']);
    }
    const motivoLimpio = limpiarTexto_(motivo);
    actualizarFila_(HOJAS.RESERVAS, r._fila, {
      Estado: ESTADO.CANCELADA, Cancelado_Por: actor.Nombre || actor.Correo || '',
      Fecha_Cancelacion: ahoraTexto_(), Motivo_Cancelacion: motivoLimpio
    });
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
  if (columna === 'Correo' || columna === 'Correo_Contacto') return normCorreo_(v);
  return typeof v === 'string' ? v.trim() : v;
}

function agregarFila_(nombre, obj) {
  const h = hoja_(nombre);
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(x => String(x).trim());
  h.appendRow(enc.map(k => (obj[k] === undefined ? '' : obj[k])));
}

/** Escribe los campos indicados en una fila, ubicando cada columna por su encabezado. */
function actualizarFila_(nombre, fila, campos) {
  const h = hoja_(nombre);
  const enc = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(x => String(x).trim());
  Object.keys(campos).forEach(k => {
    const col = enc.indexOf(k) + 1;
    if (col < 1) throw new Error('Falta la columna "' + k + '" en ' + nombre + '. Ejecuta setup().');
    h.getRange(fila, col).setValue(campos[k]);
  });
}

function reservasDelDia_(fecha) {
  return leerTabla_(HOJAS.RESERVAS).filas.filter(r => r.Estado === ESTADO.CONFIRMADA && r.Fecha_Reserva === fecha);
}

function leerSalas_() {
  return leerTabla_(HOJAS.SALAS).filas.filter(s => s.Sala && esSi_(s.Activa));
}

/** Usuarios con campos normalizados: Estado (SI/NO/PENDIENTE), Rol, DebeCambiar. */
function leerUsuarios_() {
  return leerTabla_(HOJAS.USUARIOS).filas.filter(u => u.Correo).map(u => {
    const a = String(u.Activo || '').trim().toUpperCase();
    u.Estado = esSi_(u.Activo) ? 'SI' : (a.indexOf('PEND') === 0 ? 'PENDIENTE' : 'NO');
    u.Rol = String(u.Rol || '').toUpperCase() === 'ADMIN' ? 'ADMIN' : 'USUARIO';
    u.DebeCambiar = esSi_(u.Debe_Cambiar);
    u.Nombre = u.Nombre || u.Correo;
    u.Sal = String(u.Sal || '');
    u.Hash_Contrasena = String(u.Hash_Contrasena || '');
    return u;
  });
}

function buscarUsuario_(correo) {
  const c = normCorreo_(correo);
  if (!c) return null;
  return leerUsuarios_().filter(u => u.Correo === c)[0] || null;
}

function perfil_(u) {
  return { nombre: u.Nombre, oficina: u.Oficina, correo: u.Correo, rol: u.Rol };
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
 *  6. CORREOS (encabezado institucional con logo)
 * ========================================================================== */

/** Imagen del logo para incrustar en los correos (cid:logoUnmsm). */
function logoBlob_() {
  return Utilities.newBlob(Utilities.base64Decode(LOGO_PNG_BASE64), 'image/png', 'logo-unmsm.png');
}

/**
 * HTML completo de un correo: franja de color con el logo y el nombre de la oficina,
 * el estado debajo (p. ej. "Reserva confirmada") y el cuerpo.
 */
function htmlCorreo_(cfg, color, estado, cuerpoHtml) {
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;margin:auto;border:1px solid #e5e5e5;border-radius:8px;overflow:hidden">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:' + color + ';border-collapse:collapse"><tr>' +
    '<td width="72" style="padding:14px 0 14px 20px;vertical-align:middle;width:52px">' +
    '<img src="cid:logoUnmsm" width="52" height="52" alt="UNMSM" style="display:block;border:0;width:52px;height:52px"></td>' +
    '<td style="padding:14px 20px 14px 14px;vertical-align:middle;color:#ffffff;font-family:Arial,Helvetica,sans-serif">' +
    '<div style="font-size:16px;font-weight:bold;line-height:1.3;color:#ffffff">' + escaparHtml_(cfg.NOMBRE_OFICINA) + '</div>' +
    '<div style="font-size:13px;line-height:1.4;margin-top:3px;color:#ffffff;opacity:0.92">' + escaparHtml_(estado) + '</div>' +
    '</td></tr></table>' +
    '<div style="padding:16px 24px">' + cuerpoHtml + '</div>' +
    '<div style="background:#f6f6f6;padding:10px 24px;font-size:12px;color:#777">' + escaparHtml_(cfg.NOMBRE_SERVICIO) + ' · Mensaje automático</div></div>';
}

function tablaCorreo_(filas) {
  return '<table style="border-collapse:collapse;width:100%;font-size:14px">' +
    filas.map(f => '<tr><td style="padding:6px 0;color:#666;width:120px;vertical-align:top">' + f[0] + '</td><td style="padding:6px 0"><b>' +
      escaparHtml_(String(f[1] || '')) + '</b></td></tr>').join('') + '</table>';
}

function botonCorreo_(texto, url) {
  if (!url) return '';
  return '<p style="margin:18px 0 6px"><a href="' + escaparHtml_(url) + '" style="background:' + COLOR_MARCA +
    ';color:#ffffff;text-decoration:none;padding:10px 18px;border-radius:6px;font-weight:bold;display:inline-block">' + escaparHtml_(texto) + '</a></p>';
}

function urlAplicacion_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/** Envía un correo con el logo incrustado. */
function enviarCorreoHtml_(para, asunto, html, adjuntos, cfg) {
  const opciones = { to: para, subject: asunto, htmlBody: html, name: cfg.NOMBRE_SERVICIO, inlineImages: { logoUnmsm: logoBlob_() } };
  if (adjuntos && adjuntos.length) opciones.attachments = adjuntos;
  MailApp.sendEmail(opciones);
}

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
      const c = armarCorreoReserva_(tipo, r, sala, cfg, false);
      enviarCorreoHtml_(r.Correo_Contacto, c.asunto, c.html, [c.ics], cfg);
      res.usuario = true;
    }
  } catch (e) { logCorreoFallido_(tipo, r && r.Correo_Contacto, 'usuario', e); }

  try {
    const encargados = encargados_(sala, cfg).filter(x => x !== r.Correo_Contacto);
    if (encargados.length) {
      const c = armarCorreoReserva_(tipo, r, sala, cfg, true);
      enviarCorreoHtml_(encargados.join(','), c.asunto, c.html, [c.ics], cfg);
      res.encargados = true;
    }
  } catch (e) { logCorreoFallido_(tipo, r && r.Correo_Contacto, 'encargados', e); }
  return res;
}

function encargados_(sala, cfg) {
  const lista = [];
  [sala && sala.Correo_Responsable, cfg.CORREOS_NOTIFICACION].forEach(s => {
    String(s || '').split(/[,;]/).map(x => x.trim().toLowerCase()).filter(Boolean).forEach(x => { if (lista.indexOf(x) === -1) lista.push(x); });
  });
  return lista;
}

function logCorreoFallido_(tipo, correo, para, e) {
  try { registrarLog_('correo_' + tipo, correo, 'ERROR', para + ': ' + String(e)); } catch (x) { /* nada */ }
}

function armarCorreoReserva_(tipo, r, sala, cfg, paraEncargados) {
  const t = {
    CONFIRMACION: { color: '#1a7f37', asunto: paraEncargados ? 'Nueva reserva' : 'Reserva confirmada', estado: paraEncargados ? 'Nueva reserva registrada' : 'Reserva confirmada' },
    CANCELACION: { color: '#b42318', asunto: 'Reserva cancelada', estado: 'Reserva cancelada' }
  }[tipo];
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
  return {
    asunto: '[' + t.asunto + '] ' + r.Sala + ' · ' + fechaCorta_(r.Fecha_Reserva) + ' ' + r.Hora_Inicio + ' · ' + r.ID_Reserva,
    html: htmlCorreo_(cfg, t.color, t.estado, tablaCorreo_(filas) + pie),
    ics: Utilities.newBlob(generarIcs_(r, tipo === 'CANCELACION'), 'text/calendar', 'reserva.ics')
  };
}

/** Datos de acceso (cuenta nueva o contraseña restablecida por un administrador). */
function enviarCredenciales_(u, contrasena, esRestablecimiento) {
  try {
    const cfg = leerConfig_();
    const estado = esRestablecimiento ? 'Tu contraseña fue restablecida' : 'Tu cuenta de acceso';
    const cuerpo = '<p style="font-size:14px">Hola ' + escaparHtml_(u.Nombre) + ',</p>' +
      '<p style="font-size:14px">' + (esRestablecimiento ? 'Un administrador restableció tu contraseña.' : 'Se creó tu cuenta para reservar salas.') +
      ' Estos son tus datos de acceso:</p>' +
      tablaCorreo_([['Correo', u.Correo], ['Contraseña temporal', contrasena]]) +
      '<p style="font-size:13px;color:#555">Al ingresar, el sistema te pedirá crear una contraseña nueva.</p>' +
      botonCorreo_('Ingresar al sistema', urlAplicacion_());
    enviarCorreoHtml_(u.Correo, '[' + estado + '] ' + cfg.NOMBRE_SERVICIO, htmlCorreo_(cfg, COLOR_MARCA, estado, cuerpo), [], cfg);
    return true;
  } catch (e) {
    logCorreoFallido_('credenciales', u && u.Correo, 'usuario', e);
    return false;
  }
}

/** Aviso al usuario cuando un administrador aprueba su solicitud. */
function enviarAprobacion_(u) {
  try {
    const cfg = leerConfig_();
    const cuerpo = '<p style="font-size:14px">Hola ' + escaparHtml_(u.Nombre) + ',</p>' +
      '<p style="font-size:14px">Tu solicitud fue aprobada. Ya puedes ingresar con tu correo y la contraseña que registraste.</p>' +
      botonCorreo_('Ingresar al sistema', urlAplicacion_());
    enviarCorreoHtml_(u.Correo, '[Cuenta aprobada] ' + cfg.NOMBRE_SERVICIO, htmlCorreo_(cfg, '#1a7f37', 'Cuenta aprobada', cuerpo), [], cfg);
    return true;
  } catch (e) {
    logCorreoFallido_('aprobacion', u && u.Correo, 'usuario', e);
    return false;
  }
}

/** Aviso a los administradores activos cuando llega una solicitud de cuenta. */
function avisarSolicitudAdmins_(s) {
  try {
    const cfg = leerConfig_();
    const admins = leerUsuarios_().filter(u => u.Rol === 'ADMIN' && u.Estado === 'SI').map(u => u.Correo);
    if (!admins.length) return false;
    const cuerpo = '<p style="font-size:14px">Una persona solicitó acceso al sistema de reservas:</p>' +
      tablaCorreo_([['Nombre', s.Nombre], ['Oficina', s.Oficina], ['Correo', s.Correo]]) +
      '<p style="font-size:13px;color:#555">Apruébala o recházala en <b>Administración › Usuarios</b>.</p>' +
      botonCorreo_('Abrir el sistema', urlAplicacion_());
    enviarCorreoHtml_(admins.join(','), '[Solicitud de acceso] ' + s.Nombre + ' – ' + s.Oficina,
      htmlCorreo_(cfg, COLOR_MARCA, 'Nueva solicitud de acceso', cuerpo), [], cfg);
    return true;
  } catch (e) {
    logCorreoFallido_('solicitud', s && s.Correo, 'admins', e);
    return false;
  }
}

function generarIcs_(r, cancelar) {
  const utc = (fecha, hora) => Utilities.formatDate(Utilities.parseDate(fecha + ' ' + hora, TZ, 'yyyy-MM-dd HH:mm'), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const esc = s => String(s || '').replace(/([,;\\])/g, '\\$1').replace(/\n/g, '\\n');
  return [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//OGPL UNMSM//Reservas//ES',
    'METHOD:' + (cancelar ? 'CANCEL' : 'PUBLISH'),
    'BEGIN:VEVENT',
    'UID:' + r.ID_Reserva + '@ogpl-unmsm',
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
 *  7. CONFIGURACIÓN INICIAL (solo desde el editor o la hoja)
 * ========================================================================== */

/**
 * Las funciones de mantenimiento son públicas para poder usarlas desde el menú, pero
 * la página también podría invocarlas. Esta verificación las limita al propietario.
 */
function soloPropietario_() {
  let activo = '', dueno = '';
  try { activo = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) { /* vacío */ }
  try { dueno = String(Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) { /* vacío */ }
  if (!activo || activo !== dueno) throw new Error('Esta función solo puede ejecutarla el propietario desde el editor o la hoja.');
}

function avisar_(texto) {
  Logger.log(texto);
  try { SpreadsheetApp.getUi().alert(texto); } catch (e) { /* ejecutado desde el editor: queda en el registro */ }
}

/** Crea o repara todas las pestañas y columnas. Seguro de ejecutar varias veces. */
function setup() {
  soloPropietario_();
  const libro = libro_();
  libro.setSpreadsheetTimeZone(TZ);
  Object.keys(ENCABEZADOS).forEach(nombre => {
    let h = libro.getSheetByName(nombre);
    if (!h) h = libro.insertSheet(nombre);
    const requeridos = ENCABEZADOS[nombre];
    const ultima = Math.max(h.getLastColumn(), 1);
    const actuales = h.getRange(1, 1, 1, ultima).getValues()[0].map(x => String(x).trim()).filter(Boolean);
    if (!actuales.length) {
      h.getRange(1, 1, 1, requeridos.length).setValues([requeridos]);
    } else {
      // Migración: agrega al final las columnas que falten (no mueve datos existentes)
      const faltan = requeridos.filter(k => actuales.indexOf(k) === -1);
      if (faltan.length) {
        if (h.getMaxColumns() < actuales.length + faltan.length) h.insertColumnsAfter(h.getMaxColumns(), actuales.length + faltan.length - h.getMaxColumns());
        h.getRange(1, actuales.length + 1, 1, faltan.length).setValues([faltan]);
      }
    }
    const total = Math.max(h.getLastColumn(), requeridos.length);
    h.getRange(1, 1, 1, total).setFontWeight('bold').setBackground('#e8eef7');
    h.setFrozenRows(1);
    // Todo como texto plano: Sheets no convierte fechas ni horas.
    h.getRange(1, 1, h.getMaxRows(), total).setNumberFormat('@');
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

  // Administrador inicial: el propietario. Si ningún ADMIN tiene contraseña, se genera una temporal.
  const dueno = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  let usuarios = leerUsuarios_();
  if (!usuarios.length) {
    agregarFila_(HOJAS.USUARIOS, { Correo: dueno, Nombre: 'Administrador inicial', Oficina: 'OGPL', Rol: 'ADMIN', Activo: 'SI', Fecha_Creacion: ahoraTexto_() });
    SpreadsheetApp.flush();
    usuarios = leerUsuarios_();
  }
  const admins = usuarios.filter(u => u.Rol === 'ADMIN' && u.Estado === 'SI');
  let mensaje = 'Setup completo.';
  if (admins.length && !admins.some(u => u.Hash_Contrasena)) {
    const admin = admins.filter(u => u.Correo === dueno)[0] || admins[0];
    const temporal = contrasenaTemporal_();
    const sal = nuevaSal_();
    actualizarFila_(HOJAS.USUARIOS, admin._fila, { Hash_Contrasena: hashContrasena_(temporal, sal), Sal: sal, Debe_Cambiar: 'SI' });
    mensaje += '\n\nUsuario administrador: ' + admin.Correo + '\nContraseña temporal: ' + temporal +
      '\n\nIngresa a la página con estos datos; el sistema te pedirá cambiarla.';
  }
  avisar_(mensaje);
}

/** Menú: asigna una contraseña temporal a cualquier usuario (útil si un administrador la olvida). */
function asignarContrasena() {
  soloPropietario_();
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Asignar contraseña', 'Correo del usuario:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const u = buscarUsuario_(r1.getResponseText());
  if (!u) { ui.alert('No existe un usuario con ese correo en la pestaña Usuarios.'); return; }
  const temporal = contrasenaTemporal_();
  const sal = nuevaSal_();
  actualizarFila_(HOJAS.USUARIOS, u._fila, { Hash_Contrasena: hashContrasena_(temporal, sal), Sal: sal, Debe_Cambiar: 'SI' });
  CacheService.getScriptCache().remove('fallos_' + u.Correo);
  ui.alert('Contraseña temporal para ' + u.Correo + ':\n\n' + temporal + '\n\nAl ingresar deberá cambiarla.' +
    (u.Estado !== 'SI' ? '\n\nAtención: la cuenta no está activa (Activo = ' + u.Estado + ').' : ''));
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

function errorInterno_(err) {
  return respuesta_(false, 'ERROR', 'Ocurrió un error interno. Intenta nuevamente en unos minutos.',
    { detalle: String((err && err.message) || err) });
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

function normCorreo_(v) {
  return String(v === undefined || v === null ? '' : v).trim().toLowerCase();
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
