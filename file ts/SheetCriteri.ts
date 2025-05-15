class SheetCriteri extends Sheet{
    private head: string[];
    private data: any[][];
    private max: string[];
    private min: string[];

  constructor(){
    super('TabellaPunteggi');
    this.initialize();
  }

  initialize(){
    [this.max, this.min, this.head, ...this.data] = this.getData();
  }

  getAllCriteri(){
    var criteri = [];
    for(var i = 0; i < this.data.length; i++){
      criteri.push(this.getCriterio(i))
    }
    return criteri;
  }

  getCriterio(i:number){
    var criterio = {};
    
    criterio['PuntiCriterio'] = this.data[i][this.head.indexOf('PUNTI')];
    //se è un macro criterio
    if(i === 0){
      criterio['Tipo'] = "Macro";
      criterio['TitoloCriterio'] = this.data[i][this.head.indexOf('CRITERIO')];
      criterio['MacroCriterio'] = this.data[i][this.head.indexOf('CRITERIO')];
    }else if(this.data[i-1][this.head.indexOf('CRITERIO')].toString() === ""){
      criterio['Tipo'] = "Macro";
      criterio['TitoloCriterio'] = this.data[i][this.head.indexOf('CRITERIO')];
      criterio['MacroCriterio'] = this.data[i][this.head.indexOf('CRITERIO')];
    //se è una casella vuota
    }else if(this.data[i][this.head.indexOf('CRITERIO')].toString() === ""){
      criterio['Tipo'] = "Vuoto"
    }else{
        //se è un sub criterio
        if(this.data[i][this.head.indexOf('FORMULA')].toString() === ''){
          criterio['Tipo'] = "Sub"
          criterio['SubCriterio'] = this.data[i][this.head.indexOf('CRITERIO')];
        //se è un criterio
        }else{
          criterio['Tipo'] = "Normale"
          criterio['Criterio'] = this.data[i][this.head.indexOf('CRITERIO')];
        }
    }
    return criterio;
  }

  getPuntMaxMin(){
    var p = {
      Pmax: this.max[1],
      Pmin: this.min[1]
    }
    return p;
  }

  getValoriPerTabella(){
    var titoliCriteri = [], lettereCriteri = [];
    for(var i = 0; i < this.data.length; i++){
      if(i === 0){
        titoliCriteri.push(this.data[i][this.head.indexOf('CRITERIO')].toString().substring(4));
        lettereCriteri.push(this.data[i][this.head.indexOf('CRITERIO')][0]);
      }else if(this.data[i-1][this.head.indexOf('CRITERIO')].toString() === ""){
        titoliCriteri.push(this.data[i][this.head.indexOf('CRITERIO')].toString().substring(4));
        lettereCriteri.push(this.data[i][this.head.indexOf('CRITERIO')][0]);
      } 
    }

    var valoriPerTabella = {TitoliCriteri: titoliCriteri, LettereCriteri: lettereCriteri, Pmin: this.min[1]}
    return valoriPerTabella;
  }
}