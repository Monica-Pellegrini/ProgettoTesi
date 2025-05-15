class SheetDatiVerbale extends Sheet{
    private head: string[]
    private data: any[][];


  constructor(){
    super('DatiVerbale');
    this.initialize();
  }

  initialize(): void{
    [this.head, ...this.data] = this.getData();
  }

  getDatiVerbale(){
    var datiVerbale = {};
    for( var j in this.head){
      datiVerbale[this.head[j].toString()] = this.data[0][j].toString();
      datiVerbale[this.head[j].toString().toUpperCase()] = this.data[0][j].toString().toUpperCase();
    }
    return datiVerbale;
  }
}