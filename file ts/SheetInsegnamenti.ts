class SheetInsegnamenti extends Sheet {

  head: string[] | null;
  data: string[][] | null;

  constructor() {
    super('ElencoInsegnamenti');
    this.head = null;
    this.data = null;
    this.initialize();
  }

  initialize(): void {
    const rawData: string[][] = this.getData();
    this.head = rawData.shift();
    this.data = rawData;

    if (this.head?.toString() === "" || this.data?.toString() === "") {
      throw new Error("Assenza di dati nel foglio ElencoInsegnamenti.");
    }
  }

  getInsegnamenti(): object[] {
    const insegnamenti: object[] = [];
    for (let i = 0; i < this.data.length; i++) {
      insegnamenti.push(this.getInsegnamento(i));
    }
    return insegnamenti;
  }

  getInsegnamento(i: number): object {
    const insegnamento: { [key: string]: string } = {};
    for (let j in this.head) {
      insegnamento[this.head[j].toString()] = this.data[i][j].toString();
    }
    return insegnamento;
  }

}