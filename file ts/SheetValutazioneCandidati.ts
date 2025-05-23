class SheetValutazioneCandidati extends Sheet {

  head: string[] | null;
  data: any[] | null;

  constructor(ValutazioneCandidati: any) {
    super(ValutazioneCandidati);
    this.head = null;
    this.data = null;
    this.initialize();
  }

  initialize(): void {
    const rawData: any[] = this.getData();
    this.head = rawData.shift() as string[];
    this.data = rawData;

    if (this.head.toString() === "" || this.data.toString() === "") {
      throw new Error("Assenza di dati nel foglio ValutazioneCandidati scelto.");
    }
  }

  getCandidati(insegnamento: string): any[] {
    if (!this.head!.toString().match('Insegnamento')) {
      throw new Error("Manca la voce 'Insegnamento' nel foglio ValutazioneCandidati scelto.");
    } else {
      const candidati: any[] = [];
      for (let i = 0; i < this.data!.length; i++) {
        if (this.data![i][this.head!.indexOf('Insegnamento')].toString() === insegnamento) {
          candidati.push(this.getCandidato(i));
        }
      }
      return candidati;
    }
  }

  getCandidato(i: number): { [key: string]: string } {
    const candidato: { [key: string]: string } = {};
    for (const j in this.head!) {
      candidato[this.head![j].toString()] = this.data![i][j].toString();
    }
    return candidato;
  }

  getSufficienti(candidati: any[], pmin: number): { CandidatoSuff: string }[] {
    try {
      const suff: string[] = [], sufficienti: { CandidatoSuff: string }[] = [];
      for (let i = 0; i < candidati.length; i++) {
        if (parseInt(candidati[i].PunteggioTotale) >= pmin) {
          suff.push(candidati[i].NomeCandidato.toString());
        }
      }

      suff.sort();
      for (let i = 0; i < suff.length; i++) {
        sufficienti.push({ CandidatoSuff: suff[i] });
      }
      return sufficienti;

    } catch (e) { throw new Error("Mancano alcune voci nel foglio ValutazioneCandidati scelto."); }
  }

  getPunteggiCandidato(nomeCandidato: string): any[] {
    const punteggi: any[] = [];
    for (let j = 0; j < this.data!.length; j++) {
      if (this.data![j].toString().includes(nomeCandidato)) {
        for (let i = 0; i < this.head!.length; i++) {
          if (this.head![i].toString().startsWith("Tot")) {
            punteggi.push(this.data![j][i]);
          }
        }
        break;
      }
    }

    return punteggi;
  }

}