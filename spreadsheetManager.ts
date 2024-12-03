//base class that opens the correct sheet and extracts data from it
class SpreadsheetManager
{
  protected sheet: GoogleAppsScript.Spreadsheet.Sheet | null;

  constructor(nomeSheet: string)
  {
    //the constructor tries to open the correct sheet
    try
    {
      this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeSheet);
      
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
        "Non è stato possibile accedere al file " + nomeSheet + ". Controlla il nome del file e riprova.",
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
    
    Logger.log(this.sheet.getDataRange().getValues())
    return this.sheet.getDataRange().getValues();
  }
}

function prova()
{
    let p = new SpreadsheetManager('ElencoCandidati');
    p.getData();
}