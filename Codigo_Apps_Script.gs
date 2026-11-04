// Envío de certificados "Construir futuro en La Pampa", Catriló, 20 de noviembre.
// Recibe nombre, apellido, DNI, correo y el PDF ya armado en el celular,
// manda el correo con el PDF adjunto y anota cada pedido en la hoja de cálculo.

const ASUNTO = 'Tu certificado: Construir futuro en La Pampa';
const NOMBRE_REMITENTE = 'Construir futuro en La Pampa';

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const correo = String(d.correo || '').trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(correo)) return respuesta({ ok: false, error: 'correo' });
    if (!d.pdf || String(d.pdf).length > 4000000) return respuesta({ ok: false, error: 'pdf' });
    if (MailApp.getRemainingDailyQuota() < 1) {
      anotar(d, correo, 'sin cupo diario');
      return respuesta({ ok: false, error: 'cupo' });
    }
    const adjunto = Utilities.newBlob(Utilities.base64Decode(d.pdf), 'application/pdf', String(d.archivo || 'Certificado.pdf'));
    MailApp.sendEmail({
      to: correo,
      subject: ASUNTO,
      name: NOMBRE_REMITENTE,
      body: 'Hola ' + d.nombre + ':\n\nTe mandamos adjunto el certificado de participación en la jornada "Construir futuro en La Pampa: jóvenes, trabajo y producción", realizada el 20 de noviembre en Catriló.\n\nGracias por venir.',
      attachments: [adjunto]
    });
    anotar(d, correo, 'enviado');
    return respuesta({ ok: true });
  } catch (err) {
    return respuesta({ ok: false, error: String(err) });
  }
}

function anotar(d, correo, estado) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (hoja.getLastRow() === 0) hoja.appendRow(['Fecha', 'Nombre', 'Apellido', 'DNI', 'Correo', 'Estado']);
  hoja.appendRow([new Date(), d.nombre, d.apellido, "'" + d.dni, correo, estado]);
}

function respuesta(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
