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

  getCommissione(): object[] {
    const commissione: object[] = [];
    for (let i = 0; i < (this.data?.length || 0); i++) {
      commissione.push(this.getCommissario(i));
    }
    return commissione;
  }

  getCommissario(i: number): { [key: string]: string } {
    const commissario: { [key: string]: string } = {};
    for (let j in this.head) {
      commissario[this.head[j].toString()] = this.data?[i][j].toString() ?? null: "";
    }

    return commissario;
  }

  getRuoli(): { Presidente: string; Segretario: string } {
    if (!this.head?.includes("Ruolo")) {
      throw new Error("Manca la voce 'Ruolo' nel foglio Commissione.");
    } else {

      for (let j in this.head) {
        if ((this.data?[j][this.head.indexOf("Ruolo")].toString() ?? null: "") === "Segretario"){
          this.segretario = this.data?[j][this.head.indexOf("NomeCommissario")].toString() ?? null: "";
        } else if ((this.data?[j][this.head.indexOf("Ruolo")].toString() ?? null: "") === "Presidente") {
          this.presidente = this.data?[j][this.head.indexOf("NomeCommissario")].toString() ?? null: "";
        }
      }

      const ruoli = { Presidente: this.presidente, Segretario: this.segretario };
      return ruoli;
    }
  }
}