class Sheet {
  private sheet: GoogleAppsScript.Spreadsheet.Sheet | null;

  constructor(sheetName: string) {
    try {
      this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      
      if (!this.sheet) {
        throw new Error("Il file specificato non esiste o non è accessibile.");
      }
    } catch (e) {
      throw new Error("Non è stato possibile accedere al file " + sheetName + ". Controlla il nome del file e riprova."); 
    }
  }

  getData(): Array<Array<string>> {
    if (!this.sheet) {
      throw new Error("Il foglio non è stato inizializzato correttamente.");
    }
    
    return this.sheet.getDataRange().getDisplayValues();
  }

  getRawData(): any[][] {
    if (!this.sheet) {
      throw new Error("Il foglio non è stato inizializzato correttamente.");
    }
    
    return this.sheet.getDataRange().getValues();
  }
}