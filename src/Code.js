// This function runs when the spreadsheet is opened and adds a custom menu.
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('🤖 Automazioni')
      .addItem('📉 Genera Messaggio Pagamenti', 'generaMessaggioPagamenti')
      .addItem('🚀 Genera Distinta', 'showSidebarDistinta')
      .addItem('➕ Aggiungi Tesserato', 'showSidebarNuovoTesserato')
      .addToUi();
}

/**
 * QUESTA È LA FUNZIONE PRINCIPALE DELLA TUA WEB APP
 * Viene eseguita ogni volta che carichi l'URL della app.
 */
function doGet(e) {
  // e.parameter.page ci dice quale pagina l'utente vuole vedere
  // (es. .../exec?page=distinta)
  const page = e.parameter.page;

  let htmlOutput;

  if (page === 'distinta') {
    // Mostra la pagina per generare la distinta
    htmlOutput = HtmlService.createHtmlOutputFromFile('SidebarDistinta')
    .setTitle('🚀 Genera Distinta')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // Ottimizza per mobile

  } else if (page === 'pagamenti') {
    // Genera il messaggio e mostralo nel popup
    const messaggio = ottieniMessaggioPagamenti(); // Usa la nuova funzione (ora CAPPED)
    const htmlTemplate = HtmlService.createTemplateFromFile('PopupChiodi');
    htmlTemplate.message = messaggio;
    
    // Calcola l'altezza come facevi prima
    const lineCount = messaggio.split('\n').length;
    const dialogHeight = 125 + (lineCount * 18);

    htmlOutput = htmlTemplate.evaluate()
        .setTitle('📉 Genera Messaggio Pagamenti')
        .setWidth(450) // Nota: questo verrà ignorato su mobile, è ok
        .setHeight(dialogHeight)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // Ottimizza per mobile
  
  } else if (page == 'nuovo_tesserato') {
    // Mostra la pagina per aggiungere un tesserato
    htmlOutput = HtmlService.createHtmlOutputFromFile('SidebarNuovoTesserato')
        .setTitle('➕ Aggiungi Tesserato')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  } else if (page === 'gestisci_pagamenti') {
    // Mostra la pagina per gestire i pagamenti
    htmlOutput = HtmlService.createHtmlOutputFromFile('SidebarPagamenti')
        .setTitle('💰 Aggiungi Pagamento')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  
  } else if (page === 'gestisci_debiti') {
    htmlOutput = HtmlService.createHtmlOutputFromFile('SidebarDebiti')
        .setTitle('💰 Vedi Debiti')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } else {
    // Se non viene specificata nessuna pagina mostra il menù principale
    const saldoData = getSaldoTotale(); // Usa la funzione (ora CAPPED)
    const htmlTemplate = HtmlService.createTemplateFromFile('Index.html');

    // 3. Passiamo i dati al template
    htmlTemplate.saldo = saldoData.saldo;
    htmlTemplate.saldoCell = saldoData.cell;

    htmlOutput = htmlTemplate.evaluate()
        .setTitle('🤖 Automazioni')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // Ottimizza per mobile
  }
  return htmlOutput;
}

/**
 * ================================================================================================
 * ======================= DATA ACCESS & CACHING LAYER (GETTERS) ==================================
 * ================================================================================================
 *
 * Tutte le funzioni che LEGGONO dati dal foglio sono qui.
 * Usano la cache per essere veloci.
 */

/**
 * FUNZIONE PER PRENDERE IL SALDO TOTALE
 * @returns {object} Oggetto con {saldo: number|string, cell: string}
 */
function getSaldoTotale() {
  const CACHE_KEY = 'saldoTotale';
  const CACHE_TIME = 300;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);
  
  if (cached != null) {
    Logger.log('CACHE HIT: ' + CACHE_KEY);
    return JSON.parse(cached); // I dati in cache sono stringhe
  }
  
  Logger.log('CACHE MISS: ' + CACHE_KEY);
  const CELLA_SALDO = 'N34';
  let dataToCache;

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pagamenti');
    if (!sheet) throw new Error('Foglio "pagamenti" non trovato');
    
    const saldo = sheet.getRange(CELLA_SALDO).getValue();
    
    if (typeof saldo !== 'number') {
      Logger.log(`Valore in ${CELLA_SALDO} non è un numero, è ${typeof saldo}`);
      dataToCache = { saldo: 'N/D', cell: CELLA_SALDO };
    } else {
      dataToCache = { saldo: saldo, cell: CELLA_SALDO };
    }

  } catch (e) {
    Logger.log(e);
    dataToCache = { saldo: 'Errore', cell: CELLA_SALDO };
  }

  // Salva in cache e restituisci
  cache.put(CACHE_KEY, JSON.stringify(dataToCache), CACHE_TIME);
  return dataToCache;
}

/**
 * Finds the next upcoming match in the 'calendario' sheet.
 * CACHED FUNCTION.
 */
function getNextMatch() {
  const CACHE_KEY = 'nextMatch';
  const CACHE_TIME = 300;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);
  
  if (cached != null) {
    Logger.log('CACHE HIT: ' + CACHE_KEY);
    return JSON.parse(cached);
  }
  
  Logger.log('CACHE MISS: ' + CACHE_KEY);
  let resultData = null;

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('calendario');
    const values = sheet.getRange('A2:F23').getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Normalize to midnight

    for (const row of values) {
      const matchDate = row[0]; // Assuming date is in column A
      if (!matchDate) continue;
      
      if (matchDate >= today) {
        const formattedDate = Utilities.formatDate(matchDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');
        resultData = {
          date: formattedDate,
          hour: row[1],
          homeAway: row[2], // C/T in column C
          opponent: row[3], // Opponent in column D
          at : row[4],
          matchId: row[5]   // Match ID in column F
        };
        break; // Trovata, esci dal ciclo
      }
    }
  } catch (e) {
    Logger.log(e);
    resultData = null; // Assicura null in caso di errore
  }
  
  // Metti in cache il risultato (anche se è null) e restituiscilo
  cache.put(CACHE_KEY, JSON.stringify(resultData), CACHE_TIME);
  return resultData;
}

/**
 * Gets a list of people from the 'tesserati' sheet.
 * CACHED FUNCTION.
 */
function getTesseratiData() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'tesseratiData';
  const CACHE_DURATION_SECONDS = 1800; // 30 minuti

  // 1. Prova a leggere dalla cache
  const cachedData = cache.get(CACHE_KEY);

  if (cachedData) {
    Logger.log("CACHE HIT: " + CACHE_KEY);
    return JSON.parse(cachedData);
  } 
  
  Logger.log("CACHE MISS: " + CACHE_KEY);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tesserati');
    const data = sheet.getRange('A2:R21').getValues();
    const headers = sheet.getRange('A1:R1').getValues()[0];
    const headerMap = {};
    headers.forEach((header, i) => { headerMap[header] = i; });
    
    const results = [];
    data.forEach(row => {
      if (!row[headerMap['id']]) {
        return;
      }
      const personData = {};
      for (const header in headerMap) {
        let value = row[headerMap[header]];
        if (value instanceof Date) {
          value = value.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' });
        }
        personData[header] = value;
      }
      results.push(personData);
    });
    
    // 2. Salva i dati letti nella cache per la prossima volta
    cache.put(CACHE_KEY, JSON.stringify(results), CACHE_DURATION_SECONDS);
    return results;
    
  } catch (e) {
    Logger.log(e);
    return []; // Ritorna array vuoto in caso di errore
  }
}

/**
 * Legge gli header dal foglio 'tesserati' e li passa alla sidebar.
 * CACHED FUNCTION.
 */
function getTesseratiHeaders() {
  const CACHE_KEY = 'tesseratiHeaders';
  const CACHE_TIME = 300;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);

  if (cached != null) {
    Logger.log('CACHE HIT: ' + CACHE_KEY);
    return JSON.parse(cached);
  }
  Logger.log('CACHE MISS: ' + CACHE_KEY);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tesserati');
    if (!sheet) throw new Error("Foglio 'tesserati' non trovato.");
    
    const headers = sheet.getRange('A1:R1').getValues()[0];
    const resultData = headers.slice(1); // Rimuove "id"
    
    cache.put(CACHE_KEY, JSON.stringify(resultData), CACHE_TIME);
    return resultData;
    
  } catch (e) {
    Logger.log(e);
    return { error: e.toString() }; // Per la gestione errori lato client
  }
}

/**
 * Legge tutti i debiti (<0) dal foglio pagamenti.
 * CACHED FUNCTION.
 */
function getDebitiData() {
  const CACHE_KEY = 'debitiData';
  const CACHE_TIME = 300;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);

  if (cached != null) {
    Logger.log('CACHE HIT: ' + CACHE_KEY);
    return JSON.parse(cached);
  }
  Logger.log('CACHE MISS: ' + CACHE_KEY);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pagamenti');
    if (!sheet) throw new Error("Foglio 'pagamenti' non trovato.");

    const headers = sheet.getRange('C1:AT1').getValues()[0];
    const nomiRange = sheet.getRange('B3:B22').getValues();
    const dataRange = sheet.getRange('C3:AT22').getValues();
    const startRow = 3;
    const startCol = 3;
    const listaDebitori = [];

    for (let row = 0; row < nomiRange.length; row++) {
      const cognome = nomiRange[row][0];
      if (!cognome) continue;
      const debitiDettaglio = [];
      
      for (let j = 0; j < headers.length; j += 2) {
        const spuntaColIndex = j;
        const quotaColIndex = j + 1;
        if (quotaColIndex >= headers.length) continue; 
        const importo = dataRange[row][quotaColIndex];
        const isPagato = dataRange[row][spuntaColIndex];
        
        if (typeof importo === 'number' && importo < 0 && !isPagato) {
          const matchID = headers[j];
          const riga = startRow + row;
          const col = startCol + spuntaColIndex;
          const cellaSpuntaA1 = sheet.getRange(riga, col).getA1Notation();
          debitiDettaglio.push({
            match: matchID || `Partita ${j+1}`,
            importo: importo,
            cellaSpunta: cellaSpuntaA1
          });
        }
      }
      if (debitiDettaglio.length > 0) {
        listaDebitori.push({
          cognome: cognome,
          debiti: debitiDettaglio
        });
      }
    }
    
    const resultData = { success: true, data: listaDebitori };
    cache.put(CACHE_KEY, JSON.stringify(resultData), CACHE_TIME);
    return resultData;

  } catch (e) {
    Logger.log(e);
    return { success: false, message: e.toString() }; // Non mettere in cache l'errore
  }
}

/**
 * Estrae la logica di generazione del messaggio.
 * CACHED FUNCTION (salva in cache il messaggio finale).
 */
function ottieniMessaggioPagamenti() {
  const CACHE_KEY = 'messaggioPagamenti';
  const CACHE_TIME = 300;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);

  if (cached != null) {
    Logger.log('CACHE HIT: ' + CACHE_KEY);
    return cached; // È una stringa, non serve JSON.parse
  }
  Logger.log('CACHE MISS: ' + CACHE_KEY);
  
  let messaggio;
  let debitoreTrovato = false;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('pagamenti');
    if (!sheet) {
      throw new Error('Errore: Foglio "pagamenti" non trovato!');
    }

    const nomiRange = sheet.getRange('B3:B22').getValues(); 
    const pagamentiRange = sheet.getRange('AU3:AU22').getValues();
    
    messaggio = '🚨 Ciao a tutti amorevoli sacchi di merda, questi sono i vostri chiodi: 🚨\n\n';

    for (let i = 0; i < pagamentiRange.length; i++) {
      const importo = pagamentiRange[i][0];
      const nome = nomiRange[i][0];

      if (typeof importo === 'number' && importo < 0) {
        if (nome === 'guidi') {
          continue;
        }
        const importoFormattato = (-importo).toLocaleString('it-IT', { style: 'currency', currency: 'EUR' });
        messaggio += `- *${nome}*: _${importoFormattato}_\n`;
        debitoreTrovato = true;
      }
    }
    
    messaggio += `\ncome al solito 👇🏼\n- *revolut*: revolut.me/masaccioo`;
    messaggio += `\n- *bonifico*: IT29A0366901600514286982529\n- *paypal*: https://www.paypal.me/tommasoguidi1998`;
    
    if (!debitoreTrovato) {
      messaggio = '🎉 Tutti i pagamenti sono in regola!';
    }

  } catch (e) {
    Logger.log(e);
    messaggio = e.toString();
  }
  
  // Salva il messaggio finale in cache e restituiscilo
  cache.put(CACHE_KEY, messaggio, CACHE_TIME);
  return messaggio;
}

/**
 * Gets all necessary initial data for the sidebar: next match, players, and staff.
 * This is more efficient than calling three separate functions from the client.
 * QUESTA FUNZIONE È STATA CORRETTA.
 */
function getSidebarDistintaData() {
  const nextMatch = getNextMatch(); // Usa la funzione CAPPED
  const allTesserati = getTesseratiData(); // Usa la funzione CAPPED

  // --- FIX: La tua funzione originale era buggata ---
  // Chiamava getTesseratiData() due volte e aveva un errore di variabile.
  // Ora chiamo getTesseratiData() UNA volta e filtro i risultati.
  //
  // !! ATTENZIONE !!
  // Assumo che tu abbia una colonna 'TIPO' nel foglio 'tesserati'.
  // 'G' = Giocatore, 'D' = Dirigente/Allenatore
  // Se la colonna ha un nome diverso (es. 'RUOLO') o i valori sono diversi,
  // DEVI AGGIORNARE LE DUE RIGHE SEGUENTI.
  
  const players = allTesserati
    .sort((a, b) => a.COGNOME.localeCompare(b.COGNOME));
  
  const staff = allTesserati
    .filter(p => p.TIPO_TESSERA === 'D')
    .sort((a, b) => a.COGNOME.localeCompare(b.COGNOME));

  return {
    nextMatch: nextMatch,
    players: players,
    staff: staff
  };
}

/**
 * ================================================================================================
 * ================================ CACHE MANAGEMENT ==============================================
 * ================================================================================================
 */

/**
 * Pulisce TUTTE le cache usate dalle funzioni GET.
 * NOTA: Corretto per includere tutte le chiavi e sistemare i typo.
 */
function clearCache() {
  const cache = CacheService.getScriptCache();
  const keys = [
    'saldoTotale',
    'nextMatch',
    'tesseratiData',
    'tesseratiHeaders',
    'debitiData',
    'messaggioPagamenti'
  ];
  cache.removeAll(keys);
  Logger.log('CACHE INVALIDATA: ' + keys.join(', '));
}


/**
 * ================================================================================================
 * =============================LOGICA DI GENERAZIONE MESSAGGIO CHIODI=============================
 * ================================================================================================
 */

function generaMessaggioPagamenti() {
  const messaggio = ottieniMessaggioPagamenti(); // Usa la funzione CAPPED
  showPopup(messaggio);
}

// Shows the custom popup per i chiodi.
function showPopup(messaggio) {
  const lineCount = messaggio.split('\n').length;
  const dialogHeight = 125 + (lineCount * 18);
  const htmlTemplate = HtmlService.createTemplateFromFile('PopupChiodi');
  htmlTemplate.message = messaggio; // Pass the 'messaggio' variable to the HTML file

  const htmlOutput = htmlTemplate.evaluate()
      .setWidth(450)
      .setHeight(dialogHeight);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Chiodi 📉');
}

/**
 * ================================================================================================
 * ==============================LOGICA DI GENERAZIONE DELLA DISTINTA==============================
 * ================================================================================================
 */

// Shows the custom sidebar per la distinta.
function showSidebarDistinta() {
  const html = HtmlService.createHtmlOutputFromFile('SidebarDistinta')
      .setTitle('Generatore Distinta 🥼🐂')
      .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- MAIN LOGIC ---

/**
 * The main function called by the sidebar to generate everything.
 * @param {object} formData - The data submitted from the sidebar form.
 */
function processDistintaGeneration(formData) {
  try {
    // 1. Update the 'convocati' sheet
    const matchData = formData.matchData;
    updateConvocatiSheet(matchData.matchId, formData.players);

    // 2. Create the new 'distinta' file by copying the template
    const newFile = createDistintaFile(matchData.matchId, matchData.opponent);
    const newSheet = SpreadsheetApp.openById(newFile.getId()).getSheets()[0];
    
    // 3. Prepare all the data for filling the template
    const allTesserati = getTesseratiData(); // Usa la funzione CAPPED
    const playersData = allTesserati.filter(p => formData.players.includes(p.id));
    playersData.sort((a, b) => a.COGNOME.localeCompare(b.COGNOME));
    const coachData = allTesserati.find(p => p.id === formData.coachId);
    const directorData = allTesserati.find(p => p.id === formData.directorId);

    // 4. Build the batch update request to fill the template
    const updateRequests = buildUpdateRequest(matchData, playersData, coachData, directorData);
    
    // 5. Execute the update
    updateRequests.forEach(req => {
      newSheet.getRange(req.range).setValues(req.values);
    });
    
    // IMPORTANTE: pulisci la cache dei tesserati e della partita
    // perché potresti aver aggiornato i convocati (anche se updateConvocatiSheet
    // non influenza nessuna funzione GET... ma è buona prassi)
    // In realtà, nessuna cache viene invalidata qui.
    // clearCache(); // Decommenta se necessario
    
    return { success: true, url: newFile.getUrl() };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: e.toString() };
  }
}

/**
 * Updates the 'convocati' sheet by checking the boxes for selected players.
 * (Questa è una funzione di SCRITTURA, non va messa in cache)
 */
function updateConvocatiSheet(matchId, playerIds) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('convocati');
  const headers = sheet.getRange('A1:AV1').getValues()[0];
  const playerIdsInSheet = sheet.getRange('A2:A21').getValues().flat();
  
  const matchColumnHeader = 'conv' + matchId;
  const matchColumnIndex = headers.indexOf(matchColumnHeader);
  if (matchColumnIndex === -1) {
    throw new Error(`Colonna '${matchColumnHeader}' non trovata nel foglio 'convocati'.`);
  }

  const allFalseColumn = Array(playerIdsInSheet.length).fill([false]);
  sheet.getRange(2, matchColumnIndex + 1, playerIdsInSheet.length, 1).setValues(allFalseColumn);
  
  playerIds.forEach(playerId => {
    const rowIndex = playerIdsInSheet.indexOf(playerId);
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex + 2, matchColumnIndex + 1).setValue(true);
    }
  });
}

/**
 * Creates a new 'distinta' spreadsheet file.
 * (Questa è una funzione di SCRITTURA)
 */
function createDistintaFile(matchId, opponent) {
    const templateFile = DriveApp.getFileById("1Z6aN_VS59ZiNDC2uwnGq-HWJzjhWO-wrkGks0l3tbBI");
    const destinationFolder = DriveApp.getFolderById("1PhxbuPBbJ5xNK6BHFDDNhvgPfb38kWlQ");
    const newFileName = `${matchId} vs ${opponent}`;
    return templateFile.makeCopy(newFileName, destinationFolder);
}

/**
 * Builds the data structure for the batchUpdate call.
 * (Questa è una funzione di LOGICA, non di accesso ai dati)
 */
function buildUpdateRequest(matchData, playersData, coachData, directorData) {
    const requests = [];

    // --- Match Info ---
    requests.push({
        range: 'D4',
        values: [[matchData.homeAway === 'C' ? 'DOCTORI' : matchData.opponent]]
    });
    requests.push({
        range: 'G4',
        values: [[matchData.homeAway === 'T' ? 'DOCTORI' : matchData.opponent]]
    });
    requests.push({
        range: 'E5',
        values: [[matchData.date]]
    });

    // --- Player Info ---
    let playerStartRow = 9;
    const playersValues = playersData.map(p => ([
        p.N_MAGLIA || '',
        p.COGNOME,
        p.NOME,
        '',
        p.DATA_DI_NASCITA,
        p.TIPO_TESSERA,
        p.N_TESSERA,
        p.DOCUMENTO,
        p.N_DOCUMENTO
    ]));
    if (playersValues.length > 0) {
        requests.push({
            range: `C${playerStartRow}:K${playerStartRow + playersValues.length - 1}`,
            values: playersValues
        });
    }

    // --- Staff Info ---
    if (coachData) {
        requests.push({
            range: 'D25:K25',
            values: [[
                `${coachData.COGNOME} ${coachData.NOME}`, '', '', '', coachData.TIPO_TESSERA,
                coachData.N_TESSERA, coachData.DOCUMENTO, coachData.N_DOCUMENTO
            ]]
        });
    }
    if (directorData) {
        requests.push({
            range: 'E26:K26',
            values: [[
                `${directorData.COGNOME} ${directorData.NOME}`, '', '', directorData.TIPO_TESSERA,
                directorData.N_TESSERA, directorData.DOCUMENTO, directorData.N_DOCUMENTO
            ]]
        });
    }

    return requests;
}

/**
 * ================================================================================================
 * ===============================LOGICA DI AGGIUNTA NUOVO TESSERATO===============================
 * ================================================================================================
 */

/**
 * Mostra la sidebar per aggiungere un nuovo tesserato.
 */
function showSidebarNuovoTesserato() {
  const html = HtmlService.createHtmlOutputFromFile('SidebarNuovoTesserato')
      .setTitle('Nuovo Tesserato ➕')
      .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Trova la prima riga vuota nel range B2:B21.
 * (Questa funzione LEGGE, ma non deve essere CAPPATA,
 * perché deve trovare lo stato *attuale* delle righe vuote)
 */
function findFirstEmptyRow(sheet, rangeStr) {
  const range = sheet.getRange(rangeStr);
  const startRow = range.getRow();
  const values = range.getValues().flat();

  const firstEmptyIndex = values.findIndex(cell => cell === "");
  
  if (firstEmptyIndex === -1) {
    return -1; 
  }
  
  return firstEmptyIndex+ startRow; 
}

/**
 * Riceve i dati del form dalla sidebar e li scrive sul foglio.
 * (Questa è una funzione di SCRITTURA)
 */
function aggiungiNuovoTesserato(formData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tesserati');
    if (!sheet) throw new Error("Foglio 'tesserati' non trovato.");

    const targetRow = findFirstEmptyRow(sheet, 'B2:B21');
    if (targetRow === -1) {
      throw new Error("Nessuna riga libera trovata nel range B2:B21.");
    }
    
    // Leggiamo gli header (NON dalla cache, per sicurezza)
    const headers = sheet.getRange('B1:R1').getValues()[0];
    
    const dataRow = headers.map(header => {
      return formData[header] || "";
    });

    sheet.getRange(targetRow, 2, 1, headers.length).setValues([dataRow]);

    // Adesso aggiungo anche il codice fiscale alle colonne degli id nei fogli 'convocati' e 'pagamenti'
    const sheetConvocati = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('convocati');
    if (!sheetConvocati) throw new Error("Foglio 'convocati' non trovato.");
    const targetRowConvocati = findFirstEmptyRow(sheetConvocati, 'A2:A21');
    if (targetRowConvocati === -1) throw new Error("Nessuna riga libera trovata nel range A2:A21.");
    sheetConvocati.getRange(targetRowConvocati, 1).setValue(formData['CF']);

    const sheetPagamenti = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pagamenti');
    if (!sheetPagamenti) throw new Error("Foglio 'pagamenti' non trovato.");
    const targetRowPagamenti = findFirstEmptyRow(sheetPagamenti, 'A3:A22');
    if (targetRowPagamenti === -1) throw new Error("Nessuna riga libera trovata nel range A3:A22.");
    sheetPagamenti.getRange(targetRowPagamenti, 1).setValue(formData['CF']);
    
    // IMPORTANTE: Un nuovo tesserato è stato aggiunto.
    // Pulisco la cache dei tesserati.
    clearCache(); // Pulisce tutta la cache
    
    return { success: true, message: `Tesserato aggiunto alla riga ${targetRow}!` };
    
  } catch (e) {
    Logger.log(e);
    return { success: false, message: e.toString() };
  }
}

/**
 * ================================================================================================
 * ==============================LOGICA DI GESTIONE PAGAMENTI PARTITE==============================
 * ================================================================================================
 */

/**
 * Mostra la sidebar per gestire i pagamenti.
 */
function showSidebarPagamenti() {
  const html = HtmlService.createHtmlOutputFromFile('SidebarPagamenti')
      .setTitle('Gestione Pagamenti 💰')
      .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Aggiorna il valore di una cella spunta (la imposta a TRUE).
 * (Questa è una funzione di SCRITTURA)
 * @param {string} cellaA1 La cella da aggiornare (es. "Y3")
 */
function setPagamentoSpuntato(cellaA1) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pagamenti');
    if (!sheet) throw new Error("Foglio 'pagamenti' non trovato.");
    
    sheet.getRange(cellaA1).setValue(true);
    SpreadsheetApp.flush(); 
    
    // IMPORTANTE: I pagamenti sono cambiati. Pulisco la cache.
    clearCache(); // Pulisce tutta la cache
    
    return { success: true };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: e.toString() };
  }
}

/**
 * Aggiorna il valore di una o più celle spunta (le imposta a TRUE).
 * (Questa è una funzione di SCRITTURA)
 * @param {string[]} celleA1 Un array di celle da aggiornare (es. ["C3", "E5"])
 */
function setPagamentiSpuntatiBatch(celleA1) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pagamenti');
    if (!sheet) throw new Error("Foglio 'pagamenti' non trovato.");
    
    if (!celleA1 || celleA1.length === 0) {
      return { success: true, message: "Nessuna cella da aggiornare." };
    }
    
    sheet.getRangeList(celleA1).setValue(true);
    SpreadsheetApp.flush(); 
    
    // IMPORTANTE: I pagamenti sono cambiati. Pulisco la cache.
    // Nota: questa funzione è chiamata da salvaPagamentiE_Aggiustamenti,
    // quindi la chiamata a clearCache() la mettiamo lì.
    
    return { success: true };
  } catch (e) {
    Logger.log(e);
    return { success: false, message: e.toString() };
  }
}


/**
 * Salva un valore di aggiustamento nella riga 25.
 * (Questa è una funzione di SCRITTURA)
 * @param {string} matchId L'identificativo della partita (es. "3A")
 * @param {number} importo Il valore numerico da scrivere
 */
function salvaAggiustamento(matchId, importo) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pagamenti');
    if (!sheet) throw new Error("Foglio 'pagamenti' non trovato.");
    
    const headers = sheet.getRange('C1:AT1').getValues()[0];
    const colIndex = headers.indexOf(matchId.toUpperCase());
    
    if (colIndex === -1) {
      throw new Error(`Header '${matchId.toUpperCase()}' non trovato nella riga 1.`);
    }
    
    const targetCol = colIndex + 4; // C (3) + 1 = D (4). Se colIndex è 0 (col C), target è 0+4=4 (col D)
    const targetRow = 25;
    
    sheet.getRange(targetRow, targetCol).setValue(importo);
    SpreadsheetApp.flush();
    
    // IMPORTANTE: I pagamenti sono cambiati. Pulisco la cache.
    // Nota: questa funzione è chiamata da salvaPagamentiE_Aggiustamenti,
    // quindi la chiamata a clearCache() la mettiamo lì.
    
    return { success: true, message: `Aggiustamento salvato per ${matchId}.` };
    
  } catch (e) {
    Logger.log(e);
    return { success: false, message: e.toString() };
  }
}

// /**
//  * Riceve sia le spunte che gli aggiustamenti e li salva.
//  * (Questa è una funzione di SCRITTURA)
//  * @param {object} payload Oggetto con { celle: string[], aggiustamento: object }
//  */
// function salvaPagamentiE_Aggiustamenti(payload) {
//   let spunteSuccess = true;
//   let aggiustamentoSuccess = true;
//   let spunteMessage = "";
//   let aggiustamentoMessage = "";
//   let spunteEffettuate = false;
//   let aggiustamentoEffettuato = false;

//   try {
//     // --- 1. Salva Spunte ---
//     if (payload.celle && payload.celle.length > 0) {
//       spunteEffettuate = true;
//       const spunteResponse = setPagamentiSpuntatiBatch(payload.celle);
//       if (!spunteResponse.success) {
//         spunteSuccess = false;
//         spunteMessage = spunteResponse.message;
//       }
//     }

//     // --- 2. Salva Aggiustamento ---
//     if (payload.aggiustamento) {
//       aggiustamentoEffettuato = true;
//       const aggResponse = salvaAggiustamento(
//         payload.aggiustamento.matchId, 
//         payload.aggiustamento.importo
//       );
//       if (!aggResponse.success) {
//         aggiustamentoSuccess = false;
//         aggiustamentoMessage = aggResponse.message;
//       }
//     }

//     // --- 3. Pulisci la cache se tutto è andato bene ---
//     if (spunteSuccess && aggiustamentoSuccess) {
//       // Pulisci la cache solo se almeno un'operazione è stata tentata
//       if (spunteEffettuate || aggiustamentoEffettuato) {
//         Logger.log("Pagamenti aggiornati, pulizia della cache...");
//         clearCache();
//       }
      
//       let msg = "";
//       if (spunteEffettuate) msg += "Spunte salvate. ";
//       if (aggiustamentoEffettuato) msg += "Aggiustamento salvato.";
//       return { success: true, message: msg.trim() || "Nessuna operazione eseguita." };

//     } else {
//       // Costruisci un messaggio di errore
//       let errorMessage = "";
//       if (!spunteSuccess) errorMessage += `Errore Spunte: ${spunteMessage} `;
//       if (!aggiustamentoSuccess) errorMessage += `Errore Aggiustamento: ${aggiustamentoMessage}`;
//       throw new Error(errorMessage.trim());
//     }

//   } catch (e) {
//     Logger.log(e);
//     return { success: false, message: e.toString() };
//   }
// }

/**
 * Riceve i dati di un nuovo pagamento e li scrive sul foglio.
 * Scrive in T:V (data, oggetto, importo), partendo dalla riga 33.
 * @param {object} pagamentoData Oggetto {data: string, oggetto: string, importo: number}
 */
function aggiungiNuovoPagamento(pagamentoData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pagamenti');
    if (!sheet) throw new Error("Foglio 'pagamenti' non trovato.");
    
    // Cerca la prima riga vuota controllando la colonna T (colonna 20)
    const targetRow = findFirstEmptyRow(sheet, 'T33:T50');
    
    if (targetRow === -1) {
      throw new Error(`Nessuna riga libera trovata nel range T33:T50.`);
    }

    // Prepara la data
    const dataParts = pagamentoData.data.split('/');
    const dataObj = new Date(dataParts[2], dataParts[1] - 1, dataParts[0]);

    // --- MODIFICA ---
    // Scrive i valori nelle colonne T, V, e AF
    
    // Colonna T (20)
    sheet.getRange(targetRow, 20).setValue(dataObj);
    
    // Colonna V (22)
    sheet.getRange(targetRow, 22).setValue(pagamentoData.oggetto);
    
    // Colonna AF (32)
    sheet.getRange(targetRow, 32).setValue(pagamentoData.importo);
    
    // --- FINE MODIFICA ---

    // Forza il ricalcolo e pulisce la cache (FONDAMENTALE per il saldo)
    SpreadsheetApp.flush();
    clearCache();
    
    return { success: true, message: `Pagamento aggiunto alla riga ${targetRow}!` };
    
  } catch (e) {
    Logger.log(e);
    return { success: false, message: e.toString() };
  }
}
