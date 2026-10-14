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
 * Devuelve { agente } si es válido, { bloqueado: true } si está bloqueado, o {}.
 */
function autenticar(cuip, pin) {
  const cuipLimpio = String(cuip || "").trim();
  if (!cuipLimpio || !String(pin || "").trim()) return {};

  const cache = CacheService.getScriptCache();
  const clave = "fallos_" + cuipLimpio;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const fallos = Number(cache.get(clave) || 0);
    if (fallos >= MAX_INTENTOS) return { bloqueado: true };
    cache.put(clave, String(fallos + 1), BLOQUEO_SEGUNDOS);
  } finally {
    lock.releaseLock();
  }

  const agente = buscarAgente(cuipLimpio, pin);
  if (agente) {
    cache.remove(clave);
    return { agente };
  }
  return {};
}

function errorDeAcceso(auth) {
  return auth.bloqueado
    ? { autorizado: false, bloqueado: true, error: MENSAJE_BLOQUEO }
    : { autorizado: false, error: "CUIP o PIN no autorizado" };
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

  if (body.action === "acceso" || body.action === "cursos") {
    return jsonOut(responderAcceso(body.cuip, body.pin, body.action));
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
      v.nombre || "",
      v.fecha || "",
      estado,
      nota,
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
