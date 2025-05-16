class SheetDatiVerbale extends Sheet {
  
  head: string[] | null | undefined;
  data: string[][] | null | undefined;

  constructor() {
    super('DatiVerbale');
    this.head = null;
    this.data = null;
    this.initialize();
  }

  initialize(): void {
    const rawData: string[][] = this.getData();
    this.head = rawData.shift();
    this.data = rawData;

    if (this.head?.toString() === "" || this.data?.toString() === "") {
      throw new Error("Assenza di dati nel foglio DatiVerbale.");
    }
  }

  getDatiVerbale(): Record<string, string> {
    const datiVerbale: Record<string, string> = {};
    for (let j in this.head) {
      datiVerbale[this.head[j].toString()] = this.data?[0][j].toString() ?? null: "";
      datiVerbale[this.head[j].toString().toUpperCase()] = this.data?[0][j].toString().toUpperCase() ?? null: "";
    }
    return datiVerbale;
  }

  getTemplateUrl(): string {
    if (!this.head?.includes('templateURL')) {
      throw new Error("Manca voce 'templateURL' nel foglio DatiVerbale.");
    } else {
      const url: string = this.data?[0][this.head.indexOf('templateURL')].toString() ?? null: "";
      return url;
    }
  }
}