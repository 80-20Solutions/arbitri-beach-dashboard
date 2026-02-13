# 🏐 Dashboard Arbitri Beach Volley V7

Dashboard interattiva con editing bidirezionale collegata a Google Sheets.

## Funzionalità

- **KPI in tempo reale**: arbitri, supervisori, età media, tornei totali
- **Grafici interattivi**: distribuzione per età, anzianità, comitato, stagione
- **Analisi critica**: arbitri inattivi, in calo, candidati promozione
- **Editing inline**: clicca su una cella per modificarla direttamente (salva su Google Sheets)
- **Profili arbitro**: click sul nome per vedere il profilo completo
- **Filtri avanzati**: per nome, comitato, stato, fascia d'età
- **Dark theme**

## Setup

### 1. Deploy Google Apps Script

1. Apri il [Google Sheet](https://docs.google.com/spreadsheets/d/11yFEnHfKU2mu8ym5JqBn11Wu1PQYV1Xj8nYfy576Zoo)
2. Menu → **Estensioni** → **Apps Script**
3. Cancella tutto e incolla il contenuto di `apps_script.js`
4. **Salva** (Ctrl+S)
5. **Deploy** → **Nuova distribuzione**
   - Tipo: **App web**
   - Esegui come: **Me**
   - Chi ha accesso: **Chiunque**
6. Copia l'URL generato

### 2. Configura la Dashboard

Apri `index.html` e sostituisci il placeholder in cima allo script:

```javascript
const API_URL = 'INSERISCI_URL_WEB_APP_QUI';
```

con l'URL copiato al passo precedente.

### 3. GitHub Pages

1. Crea un repo su GitHub
2. Pusha `index.html` e i file nella root
3. Settings → Pages → Source: main branch
4. La dashboard sarà disponibile su `https://tuousername.github.io/nomerepo/`

## Nota

Senza API configurata, la dashboard funziona in **modalità sola lettura** via CSV export (come V6). L'editing richiede il deploy dell'Apps Script.

## Tech Stack

- HTML/CSS/JS vanilla
- Chart.js 4
- Google Apps Script (backend)
