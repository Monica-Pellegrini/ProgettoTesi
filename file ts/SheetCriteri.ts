class SheetCriteri extends Sheet {

  max: number;
  min: number;
  head: string[] | null;
  data: string[][] | null;

  constructor() {
    super('TabellaPunteggi');
    this.max = 0;
    this.min = 0;
    this.head = null;
    this.data = null;
    this.initialize();
  }

  initialize(): void {
    const rawData: (string | null)[][] = this.getData();
    this.max = Number(rawData.shift()?.[1]);
    this.min = Number(rawData.shift()?.[1]);
    this.head = rawData.shift()? null: [];
    this.data = rawData? null: [];

    if (this.max.toString() === "" || this.min.toString() === "" || !this.head || (this.data?.length ?? 0) === 0) {
      throw new Error("Assenza di dati nel foglio ElencoInsegnamenti.");
    }
  }

  getAllCriteri(): Array<Record<string, string>> {
    
    if (!this.head?.includes('PUNTI')) {
      throw new Error("Manca la voce 'PUNTI' nel foglio TabellaPunteggi.");
    } else if (!this.head?.includes('CRITERIO')) {
      throw new Error("Manca la voce 'CRITERIO' nel foglio TabellaPunteggi.");
    }

    const criteri: Array<Record<string, string>> = [];
    
    for (let i: number = 0; i < (this.data?.length ?? 0); i++) {
      criteri.push(this.getCriterio(i));
    }
    
    return criteri;
  }

  getCriterio(i: number): Record<string, string> {
    
    const criterio: { [key: string]: any } = {};
    
    criterio['PuntiCriterio'] = this.data?[i][this.head?.indexOf('PUNTI') ?? ""] ?? null:"";

    if (i === 0) {
      criterio['Tipo'] = "Macro";
      criterio['TitoloCriterio'] = this.data?[i][this.head?.indexOf('CRITERIO')?? ""] ?? null:"";
      criterio['MacroCriterio'] = this.data?[i][this.head?.indexOf('CRITERIO')?? ""] ?? null:"";
      
    } else if ((this.data?[i - 1][this.head?.indexOf('CRITERIO')?? ""].toString()?? null:"") === "") {
      criterio['Tipo'] = "Macro";
      criterio['TitoloCriterio'] = this.data?[i][this.head?.indexOf('CRITERIO')?? ""] ?? null:"";
      criterio['MacroCriterio'] = this.data?[i][this.head?.indexOf('CRITERIO')?? ""] ?? null:"";

    } else if ((this.data?[i][this.head?.indexOf('CRITERIO')?? ""].toString()?? null:"") === "") {
      criterio['Tipo'] = "Vuoto";
      
    } else {

      if ((this.data?[i][this.head?.indexOf('FORMULA')?? ""].toString()?? null:"") === '') {
        criterio['Tipo'] = "Sub";
        criterio['SubCriterio'] = this.data?[i][this.head?.indexOf('CRITERIO')?? ""] ?? null:"";

      } else {
        criterio['Tipo'] = "Normale";
        criterio['Criterio'] = this.data?[i][this.head?.indexOf('CRITERIO')?? ""] ?? null:"";
      }
      
    }
    
    return criterio;
  }

  getPuntMaxMin(): Record<string, string> {
    
   return { Pmax: String(this.max), Pmin: String(this.min) };
   
  }

  getValoriPerTabella(): { TitoliCriteri: string[], LettereCriteri: string[], Pmin: number } {

   const titoliCriteri: string[] = [];
   const lettereCriteri: string[] = [];

   for (let i: number=0; i < (this.data?.length ?? 0); i++) {
     if (i === 0 || (this.data?[i - 1][this.head?.indexOf('CRITERIO')?? ""].toString()?? null:"") === "") {
      titoliCriteri.push(this.data?[i][this.head?.indexOf('CRITERIO')??""].toString().substring(4)?? null:"");
      lettereCriteri.push(this.data?[i][this.head?.indexOf('CRITERIO')??""][0]?? null:"");
     }
   }

   return { TitoliCriteri: titoliCriteri, LettereCriteri: lettereCriteri, Pmin:this.min };
  
 }
}