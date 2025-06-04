class SheetCommissione extends Sheet {

  head: string[] | null;
  data: string[][] | null;
  presidente: string;
  segretario: string;

  constructor() {
    super('Commissione');
    this.head = null;
    this.data = null;
    this.presidente = "DA INSERIRE";
    this.segretario = "DA INSERIRE";
    this.initialize();
  }

  initialize(): void {
    const rawData: string[][] = this.getData();
    this.head = rawData.shift() || [];
    this.data = rawData || [];

    if (this.head.toString() === "" || this.data.toString() === "") {
      throw new Error("Assenza di dati nel foglio Commissione.");
    }
  }

  getCommissione(): Array<Record<string, string>> {
    const commissione: Array<Record<string, string>> = [];
    for (let i = 0; i < (this.data?.length || 0); i++) {
      commissione.push(this.getCommissario(i));
    }
    return commissione;
  }

  getCommissario(i: number): Record<string, string> {
    const commissario: Record<string, string> = {};
    for (let j in this.head) {
      commissario[this.head[j].toString()] = this.data[i][j].toString();
    }

    return commissario;
  }

  getRuoli(): { Presidente: string; Segretario: string } {
    if (!this.head?.toString().match("Ruolo")) {
      throw new Error("Manca la voce 'Ruolo' nel foglio Commissione.");
    } else {

      for (let j in this.head) {
        if ((this.data[j][this.head.indexOf("Ruolo")].toString()) === "Segretario"){
          this.segretario = this.data[j][this.head.indexOf("NomeCommissario")].toString();
        } else if ((this.data[j][this.head.indexOf("Ruolo")].toString()) === "Presidente") {
          this.presidente = this.data[j][this.head.indexOf("NomeCommissario")].toString();
        }
      }

      const ruoli = { Presidente: this.presidente, Segretario: this.segretario };
      return ruoli;
    }
  }
}