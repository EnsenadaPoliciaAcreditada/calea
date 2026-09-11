/**
 * Backend de la Autoevaluación CALEA (examen.html).
 *
 * Este script vive DENTRO de un Google Sheet (Extensiones > Apps Script),
 * no se despliega desde GitHub. Se guarda aquí solo como referencia.
 *
 * El Sheet debe tener tres pestañas:
 *
 * "Agentes"    -> columnas: id | cuip | nombre | apellidos
 *                 (la llena manualmente el administrador; solo los CUIP
 *                 que aparezcan aquí pueden acceder al examen)
 * "Preguntas"  -> columnas: id | pregunta | opcionA | opcionB | opcionC | opcionD | correcta
 * "Resultados" -> se llena sola; Apps Script agrega una fila por cada envío.
 */

const AGENTES_TAB = "Agentes";
const PREGUNTAS_TAB = "Preguntas";
const RESULTADOS_TAB = "Resultados";

function buscarAgente(cuip) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(AGENTES_TAB);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const cuipIdx = headers.indexOf("cuip");
  const nombreIdx = headers.indexOf("nombre");
  const apellidosIdx = headers.indexOf("apellidos");

  const cuipBuscado = String(cuip || "").trim();
  if (!cuipBuscado) return null;

  const fila = data.find((row) => String(row[cuipIdx]).trim() === cuipBuscado);
  if (!fila) return null;

  return { nombre: fila[nombreIdx], apellidos: fila[apellidosIdx] };
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

function doGet(e) {
  const cuip = e.parameter.cuip;

  if (!cuip) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: "Falta el parámetro cuip" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const agente = buscarAgente(cuip);
  if (!agente) {
    return ContentService.createTextOutput(
      JSON.stringify({ autorizado: false })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(
    JSON.stringify({
      autorizado: true,
      nombre: agente.nombre,
      apellidos: agente.apellidos,
      preguntas: obtenerPreguntas(),
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);

  const agente = buscarAgente(body.cuip);
  if (!agente) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: "CUIP no autorizado" })
    ).setMimeType(ContentService.MimeType.JSON);
  }

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

  return ContentService.createTextOutput(
    JSON.stringify({
      puntaje,
      total,
      nombre: agente.nombre,
      apellidos: agente.apellidos,
    })
  ).setMimeType(ContentService.MimeType.JSON);
}
