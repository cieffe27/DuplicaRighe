/**
 * CONFIGURAZIONE INIZIALE
 * Modifica questi valori per adattare lo script senza toccare il codice sotto.
 */
const CONFIG = {
  NOME_FOGLIO: "Invitati",
  COLONNA_COPIE: "L", // La colonna che indica quante copie AGGIUNTIVE fare
  NOME_MENU: "Duplicatore Righe"
};

/**
 * Funzione di utilità per ottenere il foglio e gestire l'errore se non esiste.
 */
function getTargetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.NOME_FOGLIO);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Errore: Il foglio "${CONFIG.NOME_FOGLIO}" non esiste.`);
    return null;
  }
  return sheet;
}

/**
 * Converte la lettera della colonna (es. "L") in un indice numerico (es. 11).
 */
function colLetterToIndex(letter) {
  let column = 0, length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column - 1;
}

/**
 * Crea un unico menu all'apertura del foglio.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(CONFIG.NOME_MENU)
    .addItem('Duplica Singola Riga (Manuale)', 'duplicaRigaManuale')
    .addItem(`Duplica tutto (basato su Colonna ${CONFIG.COLONNA_COPIE})`, 'duplicaRigheInBaseACampo')
    .addToUi();
}

/**
 * DUPLICAZIONE MANUALE
 * Chiede una riga e quante COPIE AGGIUNTIVE creare (1 originale + N copie).
 */
function duplicaRigaManuale() {
  const sheet = getTargetSheet();
  if (!sheet) return;

  const ui = SpreadsheetApp.getUi();

  // Chiede quale riga
  const promptRiga = ui.prompt(CONFIG.NOME_MENU, 'Inserisci il numero della riga da duplicare:', ui.ButtonSet.OK_CANCEL);
  if (promptRiga.getSelectedButton() !== ui.Button.OK) return;
  const numeroRiga = parseInt(promptRiga.getResponseText());

  if (isNaN(numeroRiga) || numeroRiga < 1 || numeroRiga > sheet.getLastRow()) {
    ui.alert('Numero di riga non valido.');
    return;
  }

  // Chiede quante copie AGGIUNTIVE
  const promptVolte = ui.prompt(
    CONFIG.NOME_MENU, 
    `Quante COPIE AGGIUNTIVE vuoi creare per la riga ${numeroRiga}?\n(Esempio: se scrivi 3, avrai l'originale + 3 copie = 4 righe totali)`, 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (promptVolte.getSelectedButton() !== ui.Button.OK) return;
  const numeroVolte = parseInt(promptVolte.getResponseText());

  if (isNaN(numeroVolte) || numeroVolte < 0) {
    ui.alert('Inserisci un numero intero (0 o superiore).');
    return;
  }

  if (numeroVolte === 0) {
    ui.alert('Hai inserito 0 copie aggiuntive. Nessuna operazione eseguita.');
    return;
  }

  // Esecuzione duplicazione
  const datiRiga = sheet.getRange(numeroRiga, 1, 1, sheet.getLastColumn()).getValues();
  const bloccoNuoveRighe = Array(numeroVolte).fill(datiRiga[0]); // Crea l'array delle copie
  
  sheet.insertRowsAfter(numeroRiga, numeroVolte);
  sheet.getRange(numeroRiga + 1, 1, numeroVolte, sheet.getLastColumn()).setValues(bloccoNuoveRighe);
  
  ui.alert(`Completato: Aggiunte ${numeroVolte} copie. Ora hai ${numeroVolte + 1} righe identiche.`);
}

/**
 * DUPLICAZIONE BATCH (TUTTO IL FOGLIO)
 * Legge il valore in colonna L e aggiunge N copie per ogni riga.
 */
function duplicaRigheInBaseACampo() {
  const sheet = getTargetSheet();
  if (!sheet) return;

  const ui = SpreadsheetApp.getUi();
  const datiOriginali = sheet.getDataRange().getValues();
  const indiceColonna = colLetterToIndex(CONFIG.COLONNA_COPIE);
  const nuoveRighe = [];

  // Gestione Intestazione (Prima riga)
  if (datiOriginali.length > 0) {
    nuoveRighe.push(datiOriginali[0]);
  }

  // Scorre i dati (dalla riga 2 in poi)
  for (let i = 1; i < datiOriginali.length; i++) {
    const rigaCorrente = datiOriginali[i];
    const numeroCopieAggiuntive = parseInt(rigaCorrente[indiceColonna]);
    
    // 1. Aggiunge sempre l'originale
    nuoveRighe.push(rigaCorrente);

    // 2. Aggiunge le N copie extra se il numero è valido
    if (!isNaN(numeroCopieAggiuntive) && numeroCopieAggiuntive > 0) {
      for (let j = 0; j < numeroCopieAggiuntive; j++) {
        nuoveRighe.push(rigaCorrente);
      }
    }
  }

  // Pulizia e scrittura finale (operazione atomica)
  sheet.clearContents();
  sheet.getRange(1, 1, nuoveRighe.length, nuoveRighe[0].length).setValues(nuoveRighe);
  
  ui.alert(`Operazione completata.\nIl foglio è stato rigenerato aggiungendo le copie indicate nella colonna ${CONFIG.COLONNA_COPIE}.`);
}
