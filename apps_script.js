/**
 * Google Apps Script - API per Dashboard Arbitri Beach Volley
 * 
 * ISTRUZIONI:
 * 1. Apri il Google Sheet: https://docs.google.com/spreadsheets/d/11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo
 * 2. Menu → Estensioni → Apps Script
 * 3. Cancella tutto il contenuto e incolla questo codice
 * 4. Salva (Ctrl+S)
 * 5. Deploy → Nuova distribuzione → Tipo: App web
 *    - Esegui come: Me (dadecresce@gmail.com)
 *    - Chi ha accesso: Chiunque
 * 6. Copia l'URL generato e incollalo nella dashboard (variabile API_URL)
 */

const SHEET_ID = '11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo';
const TAB_ORGANICO = 'Organico';      // Adatta al nome esatto del tab
const TAB_ANALISI = 'Analisi stagioni'; // Adatta al nome esatto del tab

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    const organico = sheetToJSON(ss.getSheetByName(TAB_ORGANICO));
    const analisi = sheetToJSON(ss.getSheetByName(TAB_ANALISI));
    
    const result = {
      success: true,
      organico: organico,
      analisi: analisi,
      timestamp: new Date().toISOString()
    };
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    // Supporta singolo update o batch
    const updates = payload.updates || [payload];
    const results = [];
    
    updates.forEach(function(upd) {
      const tabName = upd.tab || TAB_ORGANICO;
      const sheet = ss.getSheetByName(tabName);
      
      if (!sheet) {
        results.push({ success: false, error: 'Tab non trovato: ' + tabName });
        return;
      }
      
      // row e col sono 1-indexed
      const row = parseInt(upd.row);
      const col = parseInt(upd.col);
      const value = upd.value;
      
      if (!row || !col) {
        results.push({ success: false, error: 'row e col richiesti' });
        return;
      }
      
      sheet.getRange(row, col).setValue(value);
      results.push({ success: true, row: row, col: col, value: value, tab: tabName });
    });
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, results: results }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sheetToJSON(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  // Trova la riga header (prima riga con contenuto significativo)
  let headerIdx = 0;
  for (let i = 0; i < data.length; i++) {
    if (data[i].filter(c => c !== '').length >= 3) {
      headerIdx = i;
      break;
    }
  }
  
  const headers = data[headerIdx].map(h => String(h).trim());
  const rows = [];
  
  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue; // salta righe vuote
    
    const obj = { _rowIndex: i + 1 }; // 1-indexed per update
    headers.forEach(function(h, j) {
      if (h) {
        let val = row[j];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        }
        obj[h] = val !== undefined && val !== null ? String(val) : '';
      }
    });
    rows.push(obj);
  }
  
  return rows;
}

// Test: esegui questa funzione nell'editor per verificare
function test() {
  const result = doGet({});
  Logger.log(result.getContent());
}
