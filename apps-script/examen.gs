/**
 * Backend de la Autoevaluación CALEA (examen.html) y del registro de
 * cursos de género obligatorios (cursos.html).
 *
 * Este script vive DENTRO de un Google Sheet (Extensiones > Apps Script),
 * no se despliega desde GitHub. Se guarda aquí solo como referencia.
 *
 * El Sheet debe tener estas pestañas:
 *
 * "Agentes"    -> columnas: id | cuip | pin | nombre | apellidos
 *                 (la llena manualmente el administrador; solo quien
 *                 conozca el CUIP Y el PIN puede acceder al examen.
 *                 El PIN es un código corto que el administrador entrega
 *                 a cada agente por un canal privado, NO es público.)
 * "Preguntas"  -> columnas: id | pregunta | opcionA | opcionB | opcionC | opcionD | correcta
 * "Resultados" -> se llena sola; Apps Script agrega una fila por cada envío.
 *
 * Estas dos las crea el script solo si no existen:
 *
 * "Cursos"         -> una fila por constancia registrada (agente + curso).
 *                     El equipo de acreditación puede cambiar a mano la
 *                     columna "estado" (Validado / Pendiente de revisión /
 *                     Rechazado) y escribir el motivo en "nota".
 * "Resumen cursos" -> se regenera sola: una fila por agente con el estado
 *                     de sus 3 cursos y si ya cumplió con todos.
 * "Ideas"          -> reconocimientos e ideas de mejora de Voz DSPM (ideas.html).
 *                     Los anónimos no guardan CUIP ni nombre. El equipo puede
 *                     usar "estado" (Nueva / En análisis / Implementada) y
 *                     "notas_del_equipo" para darles seguimiento.
 * "Ciudadanía"     -> comentarios de Tu Voz Cuenta (ciudadania.html), formulario
 *                     público sin CUIP. Nombre y contacto son opcionales.
 *
 * Las constancias (PDF o foto) se guardan en la carpeta de Drive
 * CONSTANCIAS_FOLDER_NAME, en el Drive de la cuenta dueña del script.
 */

const AGENTES_TAB = "Agentes";
const PREGUNTAS_TAB = "Preguntas";
const RESULTADOS_TAB = "Resultados";
const CURSOS_TAB = "Cursos";
const RESUMEN_TAB = "Resumen cursos";
const CONSTANCIAS_FOLDER_NAME = "Constancias cursos de género CALEA";

const ESTADO_VALIDADO = "Validado";
const ESTADO_PENDIENTE = "Pendiente de revisión";
const ESTADO_RECHAZADO = "Rechazado";
const ESTADO_SIN_REGISTRO = "Sin registrar";

const MAX_ARCHIVO_BYTES = 5 * 1024 * 1024;

// Límite de intentos de acceso fallidos por CUIP.
const MAX_INTENTOS = 5;
const BLOQUEO_SEGUNDOS = 15 * 60;
const MENSAJE_BLOQUEO = "Demasiados intentos fallidos. Espera 15 minutos e intenta de nuevo.";

// "Voz DSPM": reconocimientos e ideas de mejora del personal (uso interno).
const IDEAS_TAB = "Ideas";
const IDEAS_CORREO = "caleadspmensenada@gmail.com";
const IDEAS_MAX_CARACTERES = 2000;
const IDEAS_MAX_POR_PERIODO = 3; // envíos por CUIP cada 6 horas (límite de CacheService)
const IDEAS_PERIODO_SEGUNDOS = 6 * 60 * 60;

// "Tu Voz Cuenta": comentarios de la ciudadanía (formulario público, sin CUIP).
const CIUDADANIA_TAB = "Ciudadanía";
const CIUDADANIA_MAX_CARACTERES = 2000;
const CIUDADANIA_MAX_CORTO = 200;
const CIUDADANIA_MIN_SEGUNDOS = 4; // menos que esto llenando el formulario = robot
// Límite global (no hay forma de identificar a cada visitante): protege la hoja
// y la cuota diaria de correos que comparte con Voz DSPM.
const CIUDADANIA_MAX_POR_10_MIN = 20;
const CIUDADANIA_MAX_CORREOS_6H = 20;

// Cursos obligatorios. "clave" es un fragmento del nombre del curso tal como
// aparece en la página de verificación del emisor (sin acentos, en mayúsculas).
const CURSOS = [
  {
    id: "yo-se-de-genero",
    nombre: "Yo sé de género 1-2-3",
    emisor: "ONU Mujeres",
    clave: "YO SE DE GENERO",
    plataforma: "onu",
  },
  {
    id: "igualdad-sustantiva",
    nombre: "Igualdad sustantiva y derechos de las mujeres para el cambio cultural",
    emisor: "Secretaría de las Mujeres",
    clave: "IGUALDAD SUSTANTIVA",
    plataforma: "campus",
  },
  {
    id: "hostigamiento-acoso",
    nombre: "Atención a casos de Hostigamiento Sexual y Acoso Sexual en la APF",
    emisor: "Secretaría de las Mujeres",
    clave: "HOSTIGAMIENTO SEXUAL",
    plataforma: "campus",
  },
];

const CURSOS_HEADERS = [
  "registrado",
  "cuip",
  "nombre",
  "apellidos",
  "curso_id",
  "curso",
  "folio",
  "nombre_en_constancia",
  "fecha_constancia",
  "estado",
  "nota",
  "archivo",
  "url_verificacion",
];

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

function buscarAgente(cuip, pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(AGENTES_TAB);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const cuipIdx = headers.indexOf("cuip");
  const pinIdx = headers.indexOf("pin");
  const nombreIdx = headers.indexOf("nombre");
  const apellidosIdx = headers.indexOf("apellidos");

  const cuipBuscado = String(cuip || "").trim();
  const pinBuscado = String(pin || "").trim();
  if (!cuipBuscado || !pinBuscado) return null;

  const fila = data.find(
    (row) =>
      String(row[cuipIdx]).trim() === cuipBuscado &&
      String(row[pinIdx]).trim() === pinBuscado
  );
  if (!fila) return null;

  return { cuip: cuipBuscado, nombre: fila[nombreIdx], apellidos: fila[apellidosIdx] };
}

/**
 * Verifica CUIP + PIN con límite de intentos: tras MAX_INTENTOS fallidos,
 * ese CUIP queda bloqueado BLOQUEO_SEGUNDOS (contados desde el último fallo).
 * Cada intento "reserva" su lugar en el contador antes de consultar la hoja,
 * así que lanzar muchas peticiones a la vez no permite saltarse el límite.
 * Devuelve { agente } si es válido, { bloqueado: true } si está bloqueado, o
 * { restantes } con los intentos que le quedan antes del bloqueo.
 */
function autenticar(cuip, pin) {
  const cuipLimpio = String(cuip || "").trim();
  if (!cuipLimpio || !String(pin || "").trim()) return {};

  const cache = CacheService.getScriptCache();
  const clave = "fallos_" + cuipLimpio;

  let intento;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const fallos = Number(cache.get(clave) || 0);
    if (fallos >= MAX_INTENTOS) return { bloqueado: true };
    intento = fallos + 1;
    cache.put(clave, String(intento), BLOQUEO_SEGUNDOS);
  } finally {
    lock.releaseLock();
  }

  const agente = buscarAgente(cuipLimpio, pin);
  if (agente) {
    cache.remove(clave);
    return { agente };
  }
  return { restantes: MAX_INTENTOS - intento };
}

function errorDeAcceso(auth) {
  if (auth.bloqueado) {
    return { autorizado: false, bloqueado: true, error: MENSAJE_BLOQUEO };
  }
  if (auth.restantes === 0) {
    return {
      autorizado: false,
      bloqueado: true,
      intentosRestantes: 0,
      error: "CUIP o PIN no autorizados. Ese era tu último intento: tu acceso quedó bloqueado 15 minutos.",
    };
  }
  const resp = { autorizado: false, error: "CUIP o PIN no autorizado" };
  if (auth.restantes !== undefined) {
    resp.intentosRestantes = auth.restantes;
    resp.error = `CUIP o PIN no autorizados. Te ${auth.restantes === 1 ? "queda 1 intento" : `quedan ${auth.restantes} intentos`} antes de que tu acceso se bloquee 15 minutos.`;
  }
  return resp;
}

function obtenerPreguntas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PREGUNTAS_TAB);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idIdx = headers.indexOf("id");

  return data
    .filter((row) => row[idIdx] !== "")
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => {
        if (h !== "correcta") obj[h] = row[i];
      });
      return obj;
    });
}

// Acceso inicial: devuelve las preguntas del examen o el estado de los cursos.
function responderAcceso(cuip, pin, action) {
  const auth = autenticar(cuip, pin);
  if (!auth.agente) return errorDeAcceso(auth);
  const agente = auth.agente;

  if (action === "cursos") {
    return {
      autorizado: true,
      nombre: agente.nombre,
      apellidos: agente.apellidos,
      cursos: estadoCursosDeAgente(agente.cuip),
    };
  }

  if (action === "perfil") {
    return { autorizado: true, nombre: agente.nombre, apellidos: agente.apellidos };
  }

  return {
    autorizado: true,
    nombre: agente.nombre,
    apellidos: agente.apellidos,
    preguntas: obtenerPreguntas(),
  };
}

// Compatibilidad con versiones anteriores de las páginas, que mandaban el PIN
// en la URL. Las páginas actuales usan POST (action "acceso" / "cursos").
function doGet(e) {
  const cuip = e.parameter.cuip;
  const pin = e.parameter.pin;

  if (!cuip || !pin) {
    return jsonOut({ error: "Falta el parámetro cuip o pin" });
  }
  return jsonOut(responderAcceso(cuip, pin, e.parameter.action));
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  if (body.action === "acceso" || body.action === "cursos" || body.action === "perfil") {
    return jsonOut(responderAcceso(body.cuip, body.pin, body.action));
  }

  if (body.action === "enviarIdea") {
    return jsonOut(enviarIdea(body));
  }

  if (body.action === "enviarCiudadano") {
    return jsonOut(enviarCiudadano(body));
  }

  if (body.action === "registrarCurso") {
    return jsonOut(registrarCurso(body));
  }

  // Envío de respuestas del examen
  const auth = autenticar(body.cuip, body.pin);
  if (!auth.agente) {
    return jsonOut(errorDeAcceso(auth));
  }
  const agente = auth.agente;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preguntasSheet = ss.getSheetByName(PREGUNTAS_TAB);
  const data = preguntasSheet.getDataRange().getValues();
  const headers = data.shift();
  const idIdx = headers.indexOf("id");
  const correctaIdx = headers.indexOf("correcta");

  let puntaje = 0;
  const total = data.filter((row) => row[idIdx] !== "").length;

  data.forEach((row) => {
    if (row[idIdx] === "") return;
    const id = String(row[idIdx]);
    const correcta = String(row[correctaIdx]).trim().toUpperCase();
    const respuesta = body.respuestas && body.respuestas[id];
    if (respuesta && respuesta.toUpperCase() === correcta) {
      puntaje++;
    }
  });

  const resultadosSheet = ss.getSheetByName(RESULTADOS_TAB);
  resultadosSheet.appendRow([
    new Date(),
    body.cuip,
    agente.nombre,
    agente.apellidos,
    puntaje,
    total,
    JSON.stringify(body.respuestas || {}),
  ]);

  return jsonOut({
    puntaje,
    total,
    nombre: agente.nombre,
    apellidos: agente.apellidos,
  });
}

/* ==========================================================================
   Cursos de género obligatorios
   ========================================================================== */

function hojaCursos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CURSOS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(CURSOS_TAB);
    sheet.appendRow(CURSOS_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, CURSOS_HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}

// Devuelve las filas de "Cursos" como objetos, con su número de fila real.
function leerCursos() {
  const sheet = hojaCursos();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map((row, i) => {
    const obj = { _fila: i + 2 };
    headers.forEach((h, j) => (obj[h] = row[j]));
    return obj;
  });
}

function estadoCursosDeAgente(cuip) {
  const registros = leerCursos().filter((r) => String(r.cuip).trim() === cuip);
  return CURSOS.map((curso) => {
    const r = registros.find((x) => x.curso_id === curso.id);
    return {
      id: curso.id,
      nombre: curso.nombre,
      emisor: curso.emisor,
      estado: r ? r.estado : ESTADO_SIN_REGISTRO,
      folio: r ? r.folio : "",
      nota: r ? r.nota : "",
    };
  });
}

function normalizar(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .trim();
}

// Todas las palabras del nombre registrado deben aparecer en la constancia.
function nombresCoinciden(nombreAgente, nombreConstancia) {
  const enConstancia = new Set(normalizar(nombreConstancia).split(" "));
  const palabras = normalizar(nombreAgente)
    .split(" ")
    .filter((p) => p.length > 1);
  return palabras.length > 0 && palabras.every((p) => enConstancia.has(p));
}

// Evita que un texto escrito por el usuario (o traído de otro sitio) se ejecute
// como fórmula en la hoja, p. ej. "=IMPORTDATA(...)": se guarda como texto.
function comoTexto(valor) {
  const s = String(valor == null ? "" : valor);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function decodificarHtml(texto) {
  return String(texto || "")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#0?39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&nbsp;/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function urlVerificacion(curso, folio) {
  const f = encodeURIComponent(folio);
  if (curso.plataforma === "onu") {
    return "https://trainingcentre.unwomen.org/verify/index.php?q=" + f;
  }
  return "https://campusgenero.mujeres.gob.mx/mod/customcert/verify_certificate.php?code=" + f + "&qrcode=1";
}

/**
 * Consulta la página oficial de verificación del emisor.
 * resultado: "verificado" | "invalido" | "no_disponible"
 *  - "invalido" solo cuando el emisor dice expresamente que el folio no existe.
 *  - "no_disponible" cuando no se pudo confirmar en línea (sitio caído, o el
 *    curso exige iniciar sesión para verificar, como Igualdad sustantiva).
 */
function verificarEnLinea(curso, folio) {
  const url = urlVerificacion(curso, folio);
  let html;
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) return { resultado: "no_disponible", url };
    html = res.getContentText("UTF-8");
  } catch (err) {
    return { resultado: "no_disponible", url };
  }

  if (curso.plataforma === "onu") {
    if (/Invalid Certificate Code/i.test(html)) return { resultado: "invalido", url };
    if (!/acknowledges the authenticity/i.test(html)) return { resultado: "no_disponible", url };
    const campo = (etiqueta) => {
      const m = html.match(new RegExp(etiqueta + ":</span>([^<]+)", "i"));
      return m ? decodificarHtml(m[1]) : "";
    };
    return {
      resultado: "verificado",
      url,
      nombre: campo("Participant's Full Name"),
      curso: campo("Course Name"),
      fecha: campo("Date Issued"),
    };
  }

  // Campus Género (Moodle, plugin customcert)
  if (!/alert-success[^>]*>\s*Verificado/i.test(html)) return { resultado: "no_disponible", url };
  const campo = (etiqueta) => {
    const m = html.match(new RegExp(etiqueta + ":</strong>\\s*(?:<a[^>]*>)?([^<]+)", "i"));
    return m ? decodificarHtml(m[1]) : "";
  };
  return {
    resultado: "verificado",
    url,
    nombre: campo("Nombre completo"),
    curso: campo("Curso"),
    fecha: campo("Otorgado en"),
  };
}

function carpetaConstancias() {
  const carpetas = DriveApp.getFoldersByName(CONSTANCIAS_FOLDER_NAME);
  return carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(CONSTANCIAS_FOLDER_NAME);
}

function guardarArchivo(archivo, agente, curso, folio) {
  const bytes = Utilities.base64Decode(archivo.base64);
  if (bytes.length > MAX_ARCHIVO_BYTES) throw new Error("archivo_grande");
  const mime = /^(application\/pdf|image\/(png|jpeg))$/.test(archivo.mime) ? archivo.mime : "application/pdf";
  const ext = mime === "application/pdf" ? "pdf" : mime === "image/png" ? "png" : "jpg";
  const nombre = `${agente.cuip} - ${agente.nombre} ${agente.apellidos} - ${curso.nombre} - ${folio}.${ext}`;
  const blob = Utilities.newBlob(bytes, mime, nombre);
  return carpetaConstancias().createFile(blob).getUrl();
}

// Manda a la papelera de Drive la constancia que se está reemplazando
// (sigue recuperable desde la papelera durante 30 días).
function enviarAPapelera(urlArchivo) {
  const m = String(urlArchivo || "").match(/\/d\/([-\w]+)/);
  if (!m) return;
  try {
    DriveApp.getFileById(m[1]).setTrashed(true);
  } catch (err) {
    console.error("No se pudo mandar a la papelera la constancia anterior: " + err);
  }
}

function registrarCurso(body) {
  const auth = autenticar(body.cuip, body.pin);
  if (!auth.agente) return errorDeAcceso(auth);
  const agente = auth.agente;

  const curso = CURSOS.find((c) => c.id === body.curso);
  if (!curso) return { error: "Curso no reconocido" };

  const folio = String(body.folio || "").trim();
  if (!/^[A-Za-z0-9]{6,20}$/.test(folio)) {
    return { error: "El folio no es válido. Revisa que lo hayas escrito igual que en tu constancia." };
  }
  if (!body.archivo || !body.archivo.base64) {
    return { error: "Adjunta tu constancia en PDF o como foto." };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const registros = leerCursos();

    const folioAjeno = registros.find(
      (r) => String(r.folio).trim() === folio && String(r.cuip).trim() !== agente.cuip
    );
    if (folioAjeno) {
      return { error: "Este folio ya fue registrado por otra persona. Si es un error, contacta al equipo de acreditación." };
    }

    const previo = registros.find(
      (r) => String(r.cuip).trim() === agente.cuip && r.curso_id === curso.id
    );
    if (previo && previo.estado === ESTADO_VALIDADO) {
      return { error: "Este curso ya está validado. No necesitas registrarlo de nuevo." };
    }

    const v = verificarEnLinea(curso, folio);
    const nombreAgente = `${agente.nombre} ${agente.apellidos}`;
    let estado;
    let nota;

    if (v.resultado === "invalido") {
      return { error: `${curso.emisor} indica que el folio ${folio} no existe. Revisa tu constancia.` };
    } else if (v.resultado === "verificado") {
      const cursoOk = normalizar(v.curso).indexOf(curso.clave) !== -1;
      const nombreOk = nombresCoinciden(nombreAgente, v.nombre);
      if (cursoOk && nombreOk) {
        estado = ESTADO_VALIDADO;
        nota = `Verificado automáticamente con ${curso.emisor}.`;
      } else {
        estado = ESTADO_PENDIENTE;
        nota = !cursoOk
          ? `El folio pertenece a otro curso ("${v.curso}").`
          : `El nombre en la constancia ("${v.nombre}") no coincide con el registrado.`;
      }
    } else {
      estado = ESTADO_PENDIENTE;
      nota = `No se pudo verificar en línea con ${curso.emisor}; el equipo de acreditación revisará la constancia.`;
    }

    let archivoUrl;
    try {
      archivoUrl = guardarArchivo(body.archivo, agente, curso, folio);
    } catch (err) {
      if (err.message === "archivo_grande") {
        return { error: "El archivo es demasiado grande (máximo 5 MB)." };
      }
      console.error("No se pudo guardar la constancia en Drive: " + err);
      return { error: "No se pudo guardar la constancia. Intenta más tarde o avisa al equipo de acreditación." };
    }

    const fila = [
      new Date(),
      agente.cuip,
      agente.nombre,
      agente.apellidos,
      curso.id,
      curso.nombre,
      folio,
      comoTexto(v.nombre),
      comoTexto(v.fecha),
      estado,
      comoTexto(nota),
      archivoUrl,
      v.url,
    ];

    const sheet = hojaCursos();
    if (previo) {
      enviarAPapelera(previo.archivo);
      sheet.getRange(previo._fila, 1, 1, fila.length).setValues([fila]);
    } else {
      sheet.appendRow(fila);
    }

    reconstruirResumen();
    return { ok: true, estado, nota, cursos: estadoCursosDeAgente(agente.cuip) };
  } finally {
    lock.releaseLock();
  }
}

// Una fila por agente con el estado de cada curso y si ya cumplió con los 3.
function reconstruirResumen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(RESUMEN_TAB);
  if (!sheet) sheet = ss.insertSheet(RESUMEN_TAB);

  const agentesData = ss.getSheetByName(AGENTES_TAB).getDataRange().getValues();
  const h = agentesData.shift();
  const cuipIdx = h.indexOf("cuip");
  const nombreIdx = h.indexOf("nombre");
  const apellidosIdx = h.indexOf("apellidos");
  const registros = leerCursos();

  const headers = ["cuip", "nombre", "apellidos"]
    .concat(CURSOS.map((c) => c.nombre))
    .concat(["cursos validados", "cumple"]);

  const filas = agentesData
    .filter((row) => String(row[cuipIdx]).trim() !== "")
    .map((row) => {
      const cuip = String(row[cuipIdx]).trim();
      const estados = CURSOS.map((c) => {
        const r = registros.find((x) => String(x.cuip).trim() === cuip && x.curso_id === c.id);
        return r ? r.estado : ESTADO_SIN_REGISTRO;
      });
      const validados = estados.filter((e) => e === ESTADO_VALIDADO).length;
      return [cuip, row[nombreIdx], row[apellidosIdx]]
        .concat(estados)
        .concat([`${validados} de ${CURSOS.length}`, validados === CURSOS.length ? "Sí" : "No"]);
    });

  sheet.clear();
  const todas = [headers].concat(filas);
  sheet.getRange(1, 1, todas.length, headers.length).setValues(todas);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#0a1730").setFontColor("#ffffff");
  sheet.setFrozenRows(1);

  if (filas.length) {
    const colores = { [ESTADO_VALIDADO]: "#d9f2df", [ESTADO_PENDIENTE]: "#fff2cc", [ESTADO_RECHAZADO]: "#f8d7da", Sí: "#d9f2df", No: "#f3f3f3" };
    const fondos = filas.map((f) => f.map((celda) => colores[celda] || null));
    sheet.getRange(2, 1, filas.length, headers.length).setBackgrounds(fondos);
  }
  sheet.autoResizeColumns(1, headers.length);
}

// Si el equipo de acreditación cambia un estado a mano en "Cursos", se actualiza el resumen.
function onEdit(e) {
  if (e && e.range && e.range.getSheet().getName() === CURSOS_TAB) {
    reconstruirResumen();
  }
}

/* ==========================================================================
   Voz DSPM — reconocimientos e ideas de mejora (uso interno)
   ========================================================================== */

function hojaIdeas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(IDEAS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(IDEAS_TAB);
    const headers = ["fecha", "tipo", "cuip", "nombre", "lo_que_valora", "idea_de_mejora", "estado", "notas_del_equipo"];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  return sheet;
}

/**
 * Recibe un reconocimiento y/o una idea de mejora. El CUIP + PIN solo sirve
 * para confirmar que quien escribe es personal de la DSPM: si el agente marca
 * "anónimo", su nombre y CUIP NO se guardan en la hoja ni se envían por correo.
 * (El límite de envíos usa un contador temporal por CUIP que no guarda el
 * contenido y desaparece solo a las 6 horas.)
 */
function enviarIdea(body) {
  const auth = autenticar(body.cuip, body.pin);
  if (!auth.agente) return errorDeAcceso(auth);
  const agente = auth.agente;

  const limpiar = (t) => String(t || "").replace(/\r\n?/g, "\n").trim();
  const valora = limpiar(body.valora);
  const mejora = limpiar(body.mejora);
  const anonimo = body.anonimo === true;

  if (!valora && !mejora) {
    return { error: "Escribe al menos una de las dos respuestas." };
  }
  if (valora.length > IDEAS_MAX_CARACTERES || mejora.length > IDEAS_MAX_CARACTERES) {
    return { error: `Cada respuesta puede tener como máximo ${IDEAS_MAX_CARACTERES} caracteres.` };
  }

  const cache = CacheService.getScriptCache();
  const clave = "ideas_" + agente.cuip;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const enviados = Number(cache.get(clave) || 0);
    if (enviados >= IDEAS_MAX_POR_PERIODO) {
      return { error: "Ya enviaste varias ideas hoy, ¡gracias! Podrás enviar más en unas horas." };
    }
    cache.put(clave, String(enviados + 1), IDEAS_PERIODO_SEGUNDOS);
  } finally {
    lock.releaseLock();
  }

  const nombre = `${agente.nombre} ${agente.apellidos}`.trim();
  const tipo = anonimo ? "Anónimo" : "Identificado";

  hojaIdeas().appendRow([
    new Date(),
    tipo,
    anonimo ? "" : agente.cuip,
    anonimo ? "" : nombre,
    comoTexto(valora),
    comoTexto(mejora),
    "Nueva",
    "",
  ]);

  const cuerpo = [
    "Nuevo mensaje en Voz DSPM (sitio de acreditación CALEA)",
    "",
    "Enviado por: " + (anonimo ? "Anónimo (personal verificado de la DSPM)" : `${nombre} — CUIP ${agente.cuip}`),
    "",
    "Lo que más valora de la Policía Municipal de Ensenada:",
    valora || "(sin respuesta)",
    "",
    "Idea para mejorar:",
    mejora || "(sin respuesta)",
    "",
    "—",
    `Registro completo en la pestaña "${IDEAS_TAB}" del Sheet del examen.`,
  ].join("\n");

  try {
    MailApp.sendEmail({
      to: IDEAS_CORREO,
      subject: `Voz DSPM — nuevo mensaje (${tipo.toLowerCase()})`,
      body: cuerpo,
      name: "Voz DSPM",
    });
  } catch (err) {
    // El mensaje ya quedó guardado en la hoja; el correo es solo un aviso.
    console.error("No se pudo enviar el aviso por correo: " + err);
  }

  return { ok: true };
}

/* ==========================================================================
   Tu Voz Cuenta — comentarios de la ciudadanía (público, sin CUIP)
   ========================================================================== */

function hojaCiudadania() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CIUDADANIA_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(CIUDADANIA_TAB);
    const headers = [
      "fecha",
      "lo_que_destaca",
      "sugerencia",
      "reconocimiento_a_elemento",
      "colonia_o_zona",
      "nombre",
      "contacto",
      "estado",
      "notas_del_equipo",
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  return sheet;
}

// Incrementa un contador temporal; devuelve false si ya llegó al máximo.
function dentroDeLimite(clave, maximo, segundos) {
  const cache = CacheService.getScriptCache();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const n = Number(cache.get(clave) || 0);
    if (n >= maximo) return false;
    cache.put(clave, String(n + 1), segundos);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function enviarCiudadano(body) {
  // Robots: el campo trampa "sitio" es invisible para personas, y nadie llena
  // el formulario en menos de unos segundos. Se responde "ok" para no darles pistas.
  if (String(body.sitio || "").trim() !== "") return { ok: true };
  if (Number(body.segundos || 0) < CIUDADANIA_MIN_SEGUNDOS) return { ok: true };

  const limpiar = (t, max) => String(t || "").replace(/\r\n?/g, "\n").trim().slice(0, max + 1);
  const destaca = limpiar(body.destaca, CIUDADANIA_MAX_CARACTERES);
  const sugerencia = limpiar(body.sugerencia, CIUDADANIA_MAX_CARACTERES);
  const reconocimiento = limpiar(body.reconocimiento, CIUDADANIA_MAX_CARACTERES);
  const zona = limpiar(body.zona, CIUDADANIA_MAX_CORTO);
  const nombre = limpiar(body.nombre, CIUDADANIA_MAX_CORTO);
  const contacto = limpiar(body.contacto, CIUDADANIA_MAX_CORTO);

  if (!destaca && !sugerencia && !reconocimiento) {
    return { error: "Escribe al menos un comentario." };
  }
  if ([destaca, sugerencia, reconocimiento].some((t) => t.length > CIUDADANIA_MAX_CARACTERES)) {
    return { error: `Cada respuesta puede tener como máximo ${CIUDADANIA_MAX_CARACTERES} caracteres.` };
  }
  if ([zona, nombre, contacto].some((t) => t.length > CIUDADANIA_MAX_CORTO)) {
    return { error: "Alguno de los datos opcionales es demasiado largo." };
  }

  if (!dentroDeLimite("ciudadania_10min", CIUDADANIA_MAX_POR_10_MIN, 10 * 60)) {
    return { error: "En este momento estamos recibiendo muchos comentarios. Intenta de nuevo en unos minutos." };
  }

  hojaCiudadania().appendRow([
    new Date(),
    comoTexto(destaca),
    comoTexto(sugerencia),
    comoTexto(reconocimiento),
    comoTexto(zona),
    comoTexto(nombre),
    comoTexto(contacto),
    "Nuevo",
    "",
  ]);

  if (dentroDeLimite("ciudadania_correos_6h", CIUDADANIA_MAX_CORREOS_6H, 6 * 60 * 60)) {
    const cuerpo = [
      "Nuevo comentario ciudadano en Tu Voz Cuenta (sitio de acreditación CALEA)",
      "",
      "Lo que destaca o agradece:",
      destaca || "(sin respuesta)",
      "",
      "Sugerencia para servir mejor:",
      sugerencia || "(sin respuesta)",
      "",
      "Reconocimiento a un elemento:",
      reconocimiento || "(sin respuesta)",
      "",
      "Colonia o zona: " + (zona || "(no indicó)"),
      "Nombre: " + (nombre || "(anónimo)"),
      "Contacto: " + (contacto || "(no dejó datos de contacto)"),
      "",
      "—",
      `Registro completo en la pestaña "${CIUDADANIA_TAB}" del Sheet del examen.`,
    ].join("\n");
    try {
      MailApp.sendEmail({
        to: IDEAS_CORREO,
        subject: "Tu Voz Cuenta — nuevo comentario ciudadano",
        body: cuerpo,
        name: "Tu Voz Cuenta",
      });
    } catch (err) {
      console.error("No se pudo enviar el aviso por correo: " + err);
    }
  }

  return { ok: true };
}

/**
 * Ejecútala una vez desde el editor para autorizar el permiso de enviar
 * correos (necesario para Voz DSPM). Envía un correo de prueba a IDEAS_CORREO.
 */
function probarCorreo() {
  MailApp.sendEmail({
    to: IDEAS_CORREO,
    subject: "Voz DSPM — correo de prueba",
    body: "Si recibiste este correo, los avisos de Voz DSPM ya funcionan.",
    name: "Voz DSPM",
  });
  Logger.log("Correo de prueba enviado a " + IDEAS_CORREO + ". Correos restantes hoy: " + MailApp.getRemainingDailyQuota());
}

/**
 * Prueba manual: selecciónala en la barra superior del editor de Apps Script
 * y pulsa "Ejecutar". La primera vez Google pide autorizar los permisos de
 * Drive y de conexión a sitios externos (necesarios para cursos.html).
 * Revisa el resultado en "Registro de ejecución".
 */
function probarVerificacion() {
  // 0) Respuesta directa de cada sitio (código HTTP 200 = Google sí puede consultarlo).
  [urlVerificacion(CURSOS[0], "XXXXXXXXXX"), urlVerificacion(CURSOS[2], "XXXXXXXXXX")].forEach((url) => {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    Logger.log("HTTP " + res.getResponseCode() + " — " + url);
  });

  // 1) Consulta a ONU Mujeres con un folio inventado: debe decir resultado "invalido".
  Logger.log("ONU Mujeres (folio inventado): " + JSON.stringify(verificarEnLinea(CURSOS[0], "XXXXXXXXXX")));

  // 2) Consulta al Campus Género con un folio inventado: debe decir resultado "no_disponible".
  Logger.log("Campus Género (folio inventado): " + JSON.stringify(verificarEnLinea(CURSOS[2], "XXXXXXXXXX")));

  // 3) Acceso a Drive: crea (si no existe) la carpeta de constancias.
  Logger.log("Carpeta de constancias: " + carpetaConstancias().getUrl());
}
