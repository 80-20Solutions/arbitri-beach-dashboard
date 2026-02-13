/**
 * IMPORT DESIGNAZIONI ESTATE 2025
 * 
 * ISTRUZIONI:
 * 1. Apri il foglio principale: https://docs.google.com/spreadsheets/d/11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo
 * 2. Menu > Estensioni > Apps Script
 * 3. Cancella tutto il codice esistente e incolla QUESTO file
 * 4. Salva (Ctrl+S)
 * 5. Esegui la funzione "importAll" (selezionala dal dropdown e premi ▶)
 * 6. Autorizza quando richiesto
 * 7. Controlla il log (Visualizza > Log) per vedere cosa ha fatto
 *
 * COSA FA:
 * - TASK 1: Aggiunge colonne tornei "Arbitro Beach" alla tab "Designazioni"
 * - TASK 2: Crea tab "Designazioni Segnapunti" con tornei da segnapunti
 */

const SOURCE_IDS = [
  '1wWqcc8Qa6DRGCG767ZbRViTjYm5cwkrIih6hAMdYTuA', // luglio
  '1X3wbqD5cUeWyNCtJ-NBghM2A2XuSYgV1F6SHsf0mP9I', // agosto
  '1xdXXqzK3pbaHAJOPxLT7ME9BWr-dRlYAGhnBYMYVJ00', // settembre
];

const MAIN_ID = '11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo';

function importAll() {
  const allRows = loadAllCSV_();
  Logger.log('Righe totali caricate: ' + allRows.length);

  // TASK 1: Arbitro Beach → tab Designazioni
  const beachRows = allRows.filter(r => r.stato === 'gara accettata' && r.funzione === 'Arbitro Beach');
  Logger.log('Righe Arbitro Beach accettate: ' + beachRows.length);
  importToDesignazioni_(beachRows, 'Designazioni', false);

  // TASK 2: Segnapunti → nuova tab
  const segnaRows = allRows.filter(r => r.stato === 'gara accettata' && r.funzione === 'Segnapunti');
  Logger.log('Righe Segnapunti accettate: ' + segnaRows.length);
  importToDesignazioni_(segnaRows, 'Designazioni Segnapunti', true);

  Logger.log('✅ FATTO!');
}

function loadAllCSV_() {
  const rows = [];
  for (const id of SOURCE_IDS) {
    const ss = SpreadsheetApp.openById(id);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      rows.push({
        garaId: r[0],
        dataOra: r[1],
        campionato: r[2],
        impiantoNome: r[3],
        impiantoIndirizzo: r[4],
        impiantoLocalita: r[5],
        impiantoComune: r[6],
        impiantoProvincia: r[7],
        garaDescrizione: r[8],
        cognome: String(r[9]).trim().toUpperCase(),
        nome: String(r[10]).trim(),
        funzione: String(r[11]).trim(),
        stato: String(r[12]).trim().toLowerCase(),
      });
    }
  }
  return rows;
}

function importToDesignazioni_(rows, tabName, createNew) {
  const ss = SpreadsheetApp.openById(MAIN_ID);

  // Raggruppa tornei per chiave = garaDescrizione + impiantoNome
  const torneiMap = {};
  for (const r of rows) {
    const key = (r.garaDescrizione || '') + '|||' + (r.impiantoNome || '');
    if (!torneiMap[key]) {
      torneiMap[key] = {
        descrizione: r.garaDescrizione || '',
        impiantoNome: r.impiantoNome || '',
        localita: r.impiantoLocalita || r.impiantoNome || '',
        primaData: r.dataOra,
        arbitri: new Set(),
      };
    }
    // Track earliest date
    if (r.dataOra && r.dataOra < torneiMap[key].primaData) {
      torneiMap[key].primaData = r.dataOra;
    }
    torneiMap[key].arbitri.add(r.cognome);
  }

  const tornei = Object.values(torneiMap);
  // Sort by date
  tornei.sort((a, b) => (a.primaData > b.primaData ? 1 : -1));
  Logger.log(tabName + ': ' + tornei.length + ' tornei trovati');

  if (createNew) {
    // Crea nuova tab o svuota esistente
    let sheet = ss.getSheetByName(tabName);
    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet(tabName);
    }
    createFullTab_(sheet, tornei);
  } else {
    // Aggiungi colonne alla tab esistente
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      Logger.log('❌ Tab "' + tabName + '" non trovata!');
      return;
    }
    appendColumns_(sheet, tornei);
  }
}

function appendColumns_(sheet, tornei) {
  // Leggi arbitri esistenti dalla colonna A (righe 4+)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  // Leggi nomi arbitri esistenti
  const arbitriRange = sheet.getRange(4, 1, Math.max(1, lastRow - 3), 1).getValues();
  const arbitriList = arbitriRange.map(r => String(r[0]).trim().toUpperCase()).filter(a => a);
  
  // Raccogli tutti gli arbitri nuovi non presenti
  const arbitriSet = new Set(arbitriList);
  const nuoviArbitri = [];
  for (const t of tornei) {
    for (const a of t.arbitri) {
      if (a && !arbitriSet.has(a)) {
        arbitriSet.add(a);
        nuoviArbitri.push(a);
      }
    }
  }
  nuoviArbitri.sort();
  
  // Aggiungi nuovi arbitri in fondo
  if (nuoviArbitri.length > 0) {
    const startRow = lastRow + 1;
    const vals = nuoviArbitri.map(a => [a]);
    sheet.getRange(startRow, 1, nuoviArbitri.length, 1).setValues(vals);
    Logger.log('Aggiunti ' + nuoviArbitri.length + ' nuovi arbitri: ' + nuoviArbitri.join(', '));
  }
  
  // Rebuild full arbitri list
  const allArbitri = arbitriList.concat(nuoviArbitri);
  
  // Check existing tournaments to avoid duplicates
  let existingTornei = new Set();
  if (lastCol >= 2) {
    const row1 = sheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];
    const row2 = sheet.getRange(2, 2, 1, lastCol - 1).getValues()[0];
    const row3 = sheet.getRange(3, 2, 1, lastCol - 1).getValues()[0];
    for (let c = 0; c < row1.length; c++) {
      existingTornei.add(String(row1[c]) + '|||' + String(row3[c]));
    }
  }
  
  // Filter out already-existing tournaments
  const newTornei = tornei.filter(t => !existingTornei.has(t.descrizione + '|||' + t.localita));
  Logger.log('Tornei nuovi da aggiungere: ' + newTornei.length + ' (di ' + tornei.length + ' totali)');
  
  if (newTornei.length === 0) return;
  
  // Write new columns
  const startCol = lastCol + 1;
  const numRows = 3 + allArbitri.length;
  const data = [];
  for (let r = 0; r < numRows; r++) data.push([]);
  
  for (const t of newTornei) {
    // Row 1: descrizione
    data[0].push(t.descrizione);
    // Row 2: data
    data[1].push(formatDate_(t.primaData));
    // Row 3: luogo  
    data[2].push(t.localita);
    // Rows 4+: x per arbitri
    for (let a = 0; a < allArbitri.length; a++) {
      data[3 + a].push(t.arbitri.has(allArbitri[a]) ? 'x' : '');
    }
  }
  
  sheet.getRange(1, startCol, numRows, newTornei.length).setValues(data);
  
  for (const t of newTornei) {
    Logger.log('  + ' + t.descrizione + ' @ ' + t.localita + ' (' + formatDate_(t.primaData) + ') - ' + t.arbitri.size + ' arbitri');
  }
}

function createFullTab_(sheet, tornei) {
  if (tornei.length === 0) {
    Logger.log('Nessun torneo per questa tab');
    sheet.getRange(1, 1).setValue('Nessun torneo trovato');
    return;
  }
  
  // Collect all unique arbitri
  const arbitriSet = new Set();
  for (const t of tornei) {
    for (const a of t.arbitri) {
      if (a) arbitriSet.add(a);
    }
  }
  const allArbitri = Array.from(arbitriSet).sort();
  
  const numRows = 3 + allArbitri.length;
  const numCols = 1 + tornei.length;
  const data = [];
  
  // Row 1: header + descrizioni
  data.push(['Tipologia'].concat(tornei.map(t => t.descrizione)));
  // Row 2: data
  data.push(['Data'].concat(tornei.map(t => formatDate_(t.primaData))));
  // Row 3: luogo
  data.push(['Luogo'].concat(tornei.map(t => t.localita)));
  // Rows 4+: arbitri
  for (const arb of allArbitri) {
    const row = [arb];
    for (const t of tornei) {
      row.push(t.arbitri.has(arb) ? 'x' : '');
    }
    data.push(row);
  }
  
  sheet.getRange(1, 1, data.length, numCols).setValues(data);
  Logger.log('Tab creata con ' + tornei.length + ' tornei e ' + allArbitri.length + ' arbitri/segnapunti');
}

function formatDate_(d) {
  if (!d) return '';
  if (d instanceof Date) {
    return Utilities.formatDate(d, 'Europe/Rome', 'dd/MM/yyyy');
  }
  // Try parsing string
  const dt = new Date(d);
  if (isNaN(dt)) return String(d);
  return Utilities.formatDate(dt, 'Europe/Rome', 'dd/MM/yyyy');
}
