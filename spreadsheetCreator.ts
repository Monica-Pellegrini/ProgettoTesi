//class to create a new spreadsheet and do operations on it
class SpreadsheetCreator
{
  private spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
  private sheet: GoogleAppsScript.Spreadsheet.Sheet | null = null;
  private timestamp: string | null = null;

  constructor()
  {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  //FORSE NON VA QUI
  //method that returns the timestamp to add to the name of the new spreadsheet created
  //MM is capitalized because otherwise it gets confused with the minutes, HH is capitalized to make it use military time
  getCurrentTimestamp(): string
  {
    let currentDate = new Date();

    this.timestamp = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    return this.timestamp;
  }

  //method that creates a file Valutazione Candidati
  createFileValutazione(): GoogleAppsScript.Spreadsheet.Sheet
  {
    this.spreadsheet.setSpreadsheetLocale("en_US"); 
    this.sheet = this.spreadsheet.insertSheet("ValutazioneCandidati " + this.getCurrentTimestamp());
    return this.sheet;
  }

  getRange(row: number, column: number)
  {
    if(!this.sheet)
    {
      throw new Error("Lo Sheet non è ancora stato creato.");
    }

    return this.sheet.getRange(row, column)
  }

  //method that sets the values ​​in the created sheet
  setValori(row: number, column: number, numRow: number, numColumn: number, values: any[][])
  {
    if(!this.sheet)
    {
      throw new Error("Lo Sheet non è ancora stato creato.");
    }

    this.sheet.getRange(row, column, numRow, numColumn).setValues(values);
  }

  //method that sets the formulas
  setFormula(row: number, column: number, formula: string): void
  {
    if(!this.sheet)
    {
      throw new Error("Lo Sheet non è ancora stato creato.");
    }

    this.sheet.getRange(row, column).setFormula(formula);
  }

  //method that changes the background of the cells that the user must fill in
  setBackground(headers: string[], tabCandidati: any[][]): void
  {
    if(!this.sheet)
    {
        throw new Error("Lo Sheet non è ancora stato creato.");
    }

    for(let i = 4; i < headers.length; i++)
    {
      if(!headers[i].startsWith("Punti") && !headers[i].startsWith("Tot"))
      {
        for(let r = 2; r <= tabCandidati.length + 1; r++)
        {
          if(tabCandidati[r-2][0] !== "")
          {
            this.sheet.getRange(r, i + 1).setBackground("#FEF2CB")
          }
        }
      }
    }
  }

}

function provaCreator()
{
    let p = new SpreadsheetCreator()
    let prova = [1]
    p.createFileValutazione();
    p.setValori(1, 1, 1, 1, [prova]);
}