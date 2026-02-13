// Apps Script V2 - Aggiunge lettura tab Designazioni
// Deploy come Web App (chiunque, anche anonimi)

const SHEET_ID = '11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo';

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    // Tab Organico
    const orgSheet = ss.getSheetByName('Organico');
    const orgData = orgSheet.getDataRange().getValues();
    let hIdx = orgData.findIndex(r => r[0] && String(r[0]).includes('Tipo Tecnico'));
    if (hIdx < 0) hIdx = 3;
    const orgHeaders = orgData[hIdx].map(String);
    const organico = [];
    for (let i = hIdx + 1; i < orgData.length; i++) {
      if (!orgData[i][0]) continue;
      const obj = { _rowIndex: i + 1 };
      orgHeaders.forEach((h, ci) => obj[h] = String(orgData[i][ci] || ''));
      organico.push(obj);
    }
    
    // Tab Analisi stagioni
    const anaSheet = ss.getSheetByName('Analisi stagioni');
    const anaData = anaSheet.getDataRange().getValues();
    const anaHeaders = anaData[0].map(String);
    const analisi = [];
    for (let i = 1; i < anaData.length; i++) {
      if (!anaData[i][0]) continue;
      const obj = { _rowIndex: i + 1 };
      anaHeaders.forEach((h, ci) => obj[h] = String(anaData[i][ci] || ''));
      analisi.push(obj);
    }
    
    // Tab Designazioni (gid=21948265)
    const desSheet = ss.getSheetByName('Designazioni');
    const desData = desSheet.getDataRange().getValues();
    // Riga 0: tipologia, Riga 1: data, Riga 2: luogo, Righe 3+: arbitri
    const numCols = desData[0].length;
    const tornei = [];
    for (let c = 1; c < numCols; c++) {
      tornei.push({
        colIndex: c,
        tipo: String(desData[0][c] || '').trim(),
        data: String(desData[1][c] || '').trim(),
        luogo: String(desData[2][c] || '').trim()
      });
    }
    
    const arbitriDesignazioni = [];
    for (let r = 3; r < desData.length; r++) {
      const cognome = String(desData[r][0] || '').trim();
      if (!cognome) continue;
      const designati = [];
      for (let c = 1; c < numCols; c++) {
        const val = String(desData[r][c] || '').trim().toLowerCase();
        if (val === 'x') {
          designati.push(c); // column index matching tornei array
        }
      }
      arbitriDesignazioni.push({ cognome, designati });
    }
    
    const designazioni = { tornei, arbitri: arbitriDesignazioni };
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      organico,
      analisi,
      designazioni
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(body.tab);
    if (!sheet) throw new Error('Tab not found: ' + body.tab);
    sheet.getRange(body.row, body.col).setValue(body.value);
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
