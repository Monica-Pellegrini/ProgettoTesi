//class to manage the tabellaPunteggi
class SheetPunteggi extends SpreadsheetManager
{
  private header: string[] | null;
  private punteggiData: any[][];

  constructor(sheetName: string)
  {
    super(sheetName);
    this.header = null;
    this.punteggiData = [];
    this.initialize();
  }

  initialize(): void
  {
    let rawData = this.getData();

    this.header = rawData.shift() as string[]; //get the header line
    this.punteggiData = rawData;
  }

  //method that returns the criteria contained in the TabellaPunteggi
  getPunteggiData(): any[][]
  {
    return this.punteggiData;
  }

  //method that returns the header contained in the TabellaPunteggi
  getHeader()
  {
    return this.header;
  }
}

function prova()
{
    let p = new SheetPunteggi('ElencoCandidati');
    p.getPunteggiData();
}