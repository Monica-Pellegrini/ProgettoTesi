//class to manage the tabellaCandidati
class SheetCandidati extends Sheet
{
  private header: string[] | null;
  private candidatiData: any[][];

  constructor(sheetName: string)
  {
    super(sheetName);
    this.header = null;
    this.candidatiData = [];
    this.initialize();
  }

  initialize()
  {
    let rawData = this.getRawData();

    this.header = rawData.shift() as string[]; //save the table header in a specific variable
    this.candidatiData = rawData;
  }

  getCandidatiData()
  {
    return this.candidatiData;
  }

  getHeader()
  {
    return this.header;
  }
}