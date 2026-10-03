function doPost(e) {
  const response = (result, message) => ContentService
    .createTextOutput(JSON.stringify({ result, message }))
    .setMimeType(ContentService.MimeType.JSON);

  try {
    const data = e.parameter && e.parameter.email
      ? e.parameter
      : JSON.parse(e.postData.contents);
    const email = String(data.email || '').trim();
    const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

    if (!emailPattern.test(email)) {
      return response('error', 'Adresse email invalide');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const timestamp = new Date().toLocaleString('fr-FR', {
      timeZone: 'Africa/Brazzaville'
    });
    const safeEmail = /^[=+\-@]/.test(email) ? "'" + email : email;

    sheet.appendRow([timestamp, safeEmail]);
    return response('success', 'Adresse enregistrée');
  } catch (error) {
    console.error(error);
    return response('error', 'Impossible d’enregistrer la demande');
  }
}
