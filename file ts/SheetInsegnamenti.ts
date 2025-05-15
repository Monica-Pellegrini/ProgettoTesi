class SheetInsegnamenti extends Sheet{
    private head: string[];
    private data: any[][];

  constructor(){
    super('ElencoInsegnamenti');
    this.initialize();
  }

  initialize(){
    [this.head, ...this.data] = this.getData();
  }

  getInsegnamenti(){
    var insegnamenti = [];
    for(var i = 0; i < this.data.length; i++){
      insegnamenti.push(this.getInsegnamento(i));
    }
    return insegnamenti;
  }

  getInsegnamento(i: number){
    var insegnamento = {};
    for(var j in this.head){
      insegnamento[this.head[j].toString()] = this.data[i][j].toString();
    }
    return insegnamento;
  }

}