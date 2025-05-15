class Sheet{
    protected sheet: GoogleAppsScript.Spreadsheet.Sheet | null;
    
  constructor(sheetName: string){
    //the constructor tries to open the correct sheet
    try
    {
      this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      
      if (!this.sheet)
      {
        throw new Error("Il file specificato non esiste o non è accessibile.");
      }
    } 
    catch (e)
    {
      //if it fails it displays an error message to the user and terminates
      SpreadsheetApp.getUi().alert
      (
        "Errore",
        "Non è stato possibile accedere al file " + sheetName + ". Controlla il nome del file e riprova.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );

      throw e; 
    }
  }
  //method that returns the values ​​contained in the open sheet
  getData()
  {
    if(!this.sheet)
    {
        throw new Error("Il foglio non è stato inizializzato correttamente.");
    }
    
    return this.sheet.getDataRange().getDisplayValues();
  }
  
}