//class to manage the tabellaPunteggi
class SheetPunteggi extends SpreadsheetManager
{
  private header: string[] | null;
  private punteggiData: any[][];
  private min: number;
  private max: number;

  constructor(sheetName: string)
  {
    super(sheetName);
    this.header = null;
    this.punteggiData = [];
    this.min = 0;
    this.max = 0;
    this.initialize();
  }

  initialize(): void
  {
    let rawData = this.getData();

    this.max = rawData.shift()?.[1] ?? null; //get the max point that can be assign to a candidate
    this.min = rawData.shift()?.[1] ?? null; //get the minimum point that a candiate can have to pass the selection

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

  getMaximumScore(): number 
  {
    return this.max;
  }

  getMinimumScore(): number 
  {
    return this.min;
  }
}

function prova()
{
    let p = new SheetPunteggi('ElencoCandidati');
    p.getPunteggiData();
}