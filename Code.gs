// Google Apps Script - Tre Stelle Menu Publisher
// Incolla questo codice in: script.google.com → progetto del menu
// Dopo ogni modifica: Distribuisci → Gestisci deployment → aggiorna "FINALE" a nuova versione

const GIORNI = ['Domenica', 'Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato'];
const MESI = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno',
              'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];

function formatData_(v) {
  if (v instanceof Date) {
    return GIORNI[v.getDay()] + ' ' + v.getDate() + ' ' + MESI[v.getMonth()] + ' ' + v.getFullYear();
  }
  return String(v).trim();
}

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

// Legge una scheda "listino" semplice: riga 1 = intestazioni (Nome | Prezzo),
// dati dalla riga 2. Restituisce [] se la scheda non esiste ancora.
function leggiListino_(sheet) {
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues().slice(1); // salta intestazione
  const out = [];
  rows.forEach(([nome, prezzo]) => {
    if (!nome) return; // riga vuota
    out.push({ nome: String(nome).trim(), prezzo: formatPrezzo_(prezzo) });
  });
  return out;
}

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Foglio Config: riga 1 = header, riga 2 = etichette, riga 3 = valori reali
  const config = ss.getSheetByName('Config');
  const configData = config.getRange('A3:C3').getValues()[0];

  // Foglio Menu: Sezione, Nome, Ingredienti, FotoURL (dati dalla riga 2)
  const menu = ss.getSheetByName('Menu');
  const rows = menu.getDataRange().getValues().slice(1); // salta intestazione

  const sezioni = {};
  rows.forEach(([sezione, nome, ingredienti, foto]) => {
    if (!nome) return; // riga vuota
    const key = String(sezione).trim().toLowerCase();
    if (!sezioni[key]) sezioni[key] = [];
    sezioni[key].push({ nome: String(nome).trim(), ingredienti, foto });
  });

  // Schede listino (Nome | Prezzo): menu alla carta e dolci
  const carta = leggiListino_(ss.getSheetByName('Carta'));
  const dolci = leggiListino_(ss.getSheetByName('Dolci'));

  const output = {
    data: formatData_(configData[0]),
    prezzoCompleto: formatPrezzo_(configData[1]),
    prezzoMezzo: formatPrezzo_(configData[2]),
    sezioni,
    carta,
    dolci
  };

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}
