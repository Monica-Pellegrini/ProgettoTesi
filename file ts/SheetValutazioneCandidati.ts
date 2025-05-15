class SheetValutazioneCandidati extends Sheet{
    private head: string[];
    private data: any[][];

  constructor(ValutazioneCandidati:string){
    super(ValutazioneCandidati);

    this.initialize();
  }

  initialize(){
    [this.head, ...this.data] = this.getData();
  }

  getCandidati(insegnamento:string){
    var candidati = [];
    for(var i = 0; i < this.data.length; i++){
      if(this.data[i][this.head.indexOf('Insegnamento')].toString() === insegnamento){
        candidati.push(this.getCandidato(i));
      }
    }
    return candidati;
  }

  getCandidato(i:string){
    var candidato = {};
    for(var j in this.head){
      candidato[this.head[j].toString()] = this.data[i][j].toString();
    }
    return candidato;
  }

  getSufficienti(candidati:any[], pmin:number){
    var suff = [], sufficienti = [];
    for(var i = 0; i < candidati.length; i++){
      if(parseInt(candidati[i].PunteggioTotale) >= parseInt(pmin)){
        suff.push(candidati[i].NomeCandidato.toString());
      }
    }
    suff.sort();
    for(var i = 0; i < suff.length; i++){
      sufficienti.push({CandidatoSuff: suff[i]});
    }
    return sufficienti;
  }
  
  getPunteggiCandidato(nomeCandidato:string){
    var punteggi = [];
    for(var j = 0; j < this.data.length; j++){
      if(this.data[j].toString().includes(nomeCandidato)){
        for(var i = 0; i < this.head.length; i++){
          if(this.head[i].toString().startsWith("Tot")){
            punteggi.push(this.data[j][i]);
          }
        }
        break;
      }
    }
    
    return punteggi;
  }

}