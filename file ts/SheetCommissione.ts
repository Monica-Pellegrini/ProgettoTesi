class SheetCommissione extends Sheet{
    private head: string[];
    private data: any[][];
    private presidente: string;
    private segretario: string;

  constructor(){
    super('Commissione');
    this.presidente = "DA INSERIRE";
    this.segretario = "DA INSERIRE";
    this.initialize();
  }

  initialize(): void{
    [this.head, ...this.data] = this.getData();
  }

  getCommissione(){
    var commissione = [];
    for(var i = 0; i < this.data.length; i++){
      commissione.push(this.getCommissario(i));
    }
    return commissione;
  }

  getCommissario(i: number){
    var commissario = {};
    for(var j in this.head){
      commissario[this.head[j].toString()] = this.data[i][j].toString();
    }
    if(this.data[i][this.head.indexOf("Ruolo")].toString() === "Segretario"){
      this.segretario = this.data[i][this.head.indexOf("NomeCommissario")].toString()
    }else if(this.data[i][this.head.indexOf("Ruolo")].toString() === "Presidente"){
      this.presidente = this.data[i][this.head.indexOf("NomeCommissario")].toString()
    }
    return commissario;
  }

  getRuoli(){
    var ruoli = {Presidente: this.presidente, Segretario: this.segretario};
    return ruoli;
  }
}