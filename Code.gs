// Google Apps Script - Tre Stelle Menu Publisher
// Incolla questo codice in: script.google.com → Nuovo progetto

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Legge Config (Data, Prezzi)
  const config = ss.getSheetByName('Config');
  const configData = config.getRange('A2:C2').getValues()[0];

  // Legge Menu (Sezione, Nome, Ingredienti, FotoURL)
  const menu = ss.getSheetByName('Menu');
  const rows = menu.getDataRange().getValues().slice(1); // salta intestazione

  const sezioni = {};
  rows.forEach(([sezione, nome, ingredienti, foto]) => {
    if (!nome) return; // riga vuota
    if (!sezioni[sezione]) sezioni[sezione] = [];
    sezioni[sezione].push({ nome, ingredienti, foto });
  });

  const output = {
    data: configData[0],
    prezzoCompleto: configData[1],
    prezzoMezzo: configData[2],
    sezioni
  };

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}
