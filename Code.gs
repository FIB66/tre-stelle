// Google Apps Script - Tre Stelle Menu Publisher
// Questo file riflette il codice REALMENTE distribuito (deployment "FINALE", @7+).
// Distribuzione gestita via clasp dal PC di Giorgio:
//   cd <cartella progetto> && clasp push --force
//   clasp redeploy AKfycbxczilq2ynSUZv5cS9BjdpZTIQcDL07ZhOkCgi0QKJuzyOf5cw1EDKRe4WJ67AYynk0Yw -d "FINALE"
//
// NB: il foglio "Menu" (menu del giorno) NON ha riga di intestazione: la riga 1 è
// già un piatto, quindi si leggono TUTTE le righe (niente slice). Le schede "Carta"
// e "Dolci" invece hanno l'intestazione (Nome | Prezzo) in riga 1 → si salta con slice(1).

function formatPrezzo_(v) {
  // Sheets a volte interpreta "12,50" come orario 12:50 → la cella diventa una data
  if (v instanceof Date) {
    return (v.getHours() + v.getMinutes() / 100).toFixed(2).replace('.', ',');
  }
  if (typeof v === 'number') {
    return v % 1 === 0 ? String(v) : v.toFixed(2).replace('.', ',');
  }
  return String(v).trim();
}

// Scheda listino: riga 1 = intestazioni, dati dalla riga 2.
// Due formati supportati (riconosciuti dall'intestazione della colonna A):
//   Nome | Prezzo                → sezione vuota (lista unica)
//   Sezione | Nome | Prezzo     → piatti raggruppati per sezione su carta.html
// Restituisce [] se la scheda non esiste ancora.
function leggiListino_(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const haSezione = String(data[0][0]).trim().toLowerCase().indexOf('sezion') === 0;
  const out = [];
  data.slice(1).forEach(function (r) {
    const sezione = haSezione ? String(r[0] || '').trim() : '';
    const nome = haSezione ? r[1] : r[0];
    const prezzo = haSezione ? r[2] : r[1];
    if (!nome) return; // riga vuota
    out.push({ sezione: sezione, nome: String(nome).trim(), prezzo: formatPrezzo_(prezzo) });
  });
  return out;
}

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Foglio Config: riga 1 = header, riga 2 = etichette, riga 3 = valori reali
  const config = ss.getSheetByName('Config');
  const configData = config.getRange('A3:C3').getValues()[0];

  // Foglio Menu del giorno: nessuna intestazione → si leggono tutte le righe
  const menu = ss.getSheetByName('Menu');
  const rows = menu.getDataRange().getValues();

  const sezioni = {};
  rows.forEach(([sezione, nome, ingredienti, foto]) => {
    if (!nome) return; // riga vuota
    const key = String(sezione).trim().toLowerCase();
    if (!sezioni[key]) sezioni[key] = [];
    sezioni[key].push({ nome: String(nome).trim(), ingredienti, foto });
  });

  // Nuove schede: menu alla carta e dolci (Nome | Prezzo)
  const carta = leggiListino_(ss.getSheetByName('Carta'));
  const dolci = leggiListino_(ss.getSheetByName('Dolci'));

  const output = {
    data: configData[0],
    prezzoCompleto: configData[1],
    prezzoMezzo: configData[2],
    sezioni,
    carta,
    dolci
  };

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}
