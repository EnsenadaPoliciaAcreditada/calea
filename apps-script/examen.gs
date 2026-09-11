/**
 * Backend de la Autoevaluación CALEA (examen.html).
 *
 * Este script vive DENTRO de un Google Sheet (Extensiones > Apps Script),
 * no se despliega desde GitHub. Se guarda aquí solo como referencia.
 *
 * El Sheet debe tener dos pestañas:
 *
 * "Preguntas"  -> columnas: id | pregunta | opcionA | opcionB | opcionC | opcionD | correcta
 * "Resultados" -> se llena sola; Apps Script agrega una fila por cada envío.
 */

const PREGUNTAS_TAB = "Preguntas";
const RESULTADOS_TAB = "Resultados";

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PREGUNTAS_TAB);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const preguntas = data
    .filter((row) => row[headers.indexOf("id")] !== "")
    .map((row) => {
      const obj = {};
      headers.forEach((h, i) => {
        if (h !== "correcta") obj[h] = row[i];
      });
      return obj;
    });

  return ContentService.createTextOutput(JSON.stringify({ preguntas })).setMimeType(
    ContentService.MimeType.JSON
  );
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
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
    body.nombre || "",
    body.area || "",
    puntaje,
    total,
    JSON.stringify(body.respuestas || {}),
  ]);

  return ContentService.createTextOutput(JSON.stringify({ puntaje, total })).setMimeType(
    ContentService.MimeType.JSON
  );
}
