function onOpen(){
  var menu = SpreadsheetApp.getUi().createMenu("Funzioni Aggiuntive");
  menu.addItem("Genera Valutazione Candidati", "main");
  menu.addItem("Genera Verbale", "generaVerbale");
  menu.addToUi();

}

function generaVerbale(){
  try{
    var verbale = new VerbalGenerator();
    verbale.generaVerbale();
  }catch(e){
    SpreadsheetApp.getUi().alert("Errore",e,SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

class UserInterface{
  constructor(){}

  makeOutputBox(link, nome){
    //crea box con link cliccabile
    var html=Utilities.formatString('<style>input{margin: 3px 0;}</style><h3>%s</h3><a href="%s" target="_blank">%s</a><br />','Link del verbale:',link,nome);
    var userInterface=HtmlService.createHtmlOutput(html).setHeight(100);
    SpreadsheetApp.getUi().showModelessDialog(userInterface, "Verbale creato");
  }
}

class VerbalGenerator{

  constructor(){
    this.activeSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    this.linkTemplate = null;
    this.initialize()
  }

  initialize(){
    this.interface = new UserInterface();
    var sheet = new SheetDatiVerbale();
    this.linkTemplate = sheet.getTemplateUrl();
  }

  generaVerbale(){

    if(this.linkTemplate !== ""){
      if(this.linkTemplate.includes("https://docs.google.com/document/")){
        if(this.activeSheetName.includes("ValutazioneCandidati")){

          var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");

          try{
            var templateId = DocumentApp.openByUrl(this.linkTemplate).getId();
          }catch(e){
            throw new Error("Non è stato possibile accedere al documento template.");
          }
            
          var verbale = new Verbale(templateId, timestamp, this.activeSheetName);
          verbale.replaceAll();

          this.interface.makeOutputBox(verbale.getLink(),verbale.getName());
          

        }else{SpreadsheetApp.getUi().alert('ATTENZIONE!\nAvviare la funzione da un foglio contenente "ValutazioneCandidati" nel nome.')}
      }else{SpreadsheetApp.getUi().alert("ATTENZIONE!\nInserire un URL valido.")}
    }else{SpreadsheetApp.getUi().alert("ATTENZIONE!\nInserire un URL valido.")}

  }
}

class Verbale{

  constructor(templateId, timestamp, sheetVC){
    this.verbaleId = DriveApp.getFileById(templateId).makeCopy("Verbale " + timestamp).getId();
    this.body = DocumentApp.openById(this.verbaleId).getBody();
    this.sheetVC = sheetVC;
  }

  getLink(){
    return DocumentApp.openById(this.verbaleId).getUrl();
  }

  getName(){
    return DocumentApp.openById(this.verbaleId).getName();
  }

  replaceAll(){
    try{
      this.replaceDatiVerbale();
      this.replaceCommissione();
      this.replaceElencoPunteggi();
      this.replaceElencoInsegnamenti();
    }catch(e){
      //in caso di errore durante la scrittura, cestina il documento
      DriveApp.getFileById(this.verbaleId).setTrashed(true);
      throw e;
    }
  }

  replaceDatiVerbale(){
    try{
      var sheet = new SheetDatiVerbale();

      var datiVerbale = sheet.getDatiVerbale();
      
      this.replacePlaceholder(datiVerbale);
    }catch(e){throw e;} 
  }

  replaceCommissione(){
    try{
      const sheet = new SheetCommissione();

      var paragraphsId = [], paragraphs = [];
      var allPar = this.body.getParagraphs();
      for(var i in allPar){
        if(allPar[i].getText().toString().includes("«NomeCommissario»") && allPar[i].getParent().getType() === DocumentApp.ElementType.BODY_SECTION){
          paragraphsId.push(this.body.getChildIndex(allPar[i]));
          paragraphs.push(allPar[i].copy());
        }
      }

      var allTabs = this.body.getTables();
      var tables = [], tRow = [];
      for( var j in allTabs){
        if(allTabs[j].findText('«NomeCommissario»')){
          tables.push(allTabs[j]);
          tRow.push(allTabs[j].findText('«NomeCommissario»').getElement().getParent().getParent().getParent().copy());
        }
      }
      
      var commissione = sheet.getCommissione();

      //per ogni commissario
      for(var i = 0; i < commissione.length; i++){
        
        this.replacePlaceholder(commissione[i]);

        if(i < commissione.length - 1){
          for(var j in paragraphsId){
            paragraphsId[j]++;
            this.body.insertParagraph(paragraphsId[j], paragraphs[j].copy());
          }
          
          for( var j in tables){
            tables[j].appendTableRow(tRow[j].copy());
          }
        }
      }
      
      var ruoli = sheet.getRuoli();

      this.replacePlaceholder(ruoli);

    }catch(e){throw e}
  }

  replaceElencoPunteggi(){
    try{
      var sheet = new SheetCriteri();

      this.replacePlaceholder(sheet.getPuntMaxMin());

      if(this.body.findText("«TitoloCriterio»")){
        var titoloCriterioId = this.body.getChildIndex(this.body.findText("«TitoloCriterio»").getElement().getParent());
        var titoloCriterioParagraph = this.body.getChild(titoloCriterioId).asParagraph().copy();        
      }

      
      if(this.body.findText("«Criterio»")){
        var critParagraph = this.body.findText("«Criterio»").getElement().getParent().asListItem().copy();
        var critId = this.body.getChildIndex(this.body.findText("«Criterio»").getElement().getParent());
        this.body.removeChild(this.body.findText("«Criterio»").getElement().getParent());
        
        if(this.body.findText("«MacroCriterio»")){
          var macroCritParagraph = this.body.findText("«MacroCriterio»").getElement().getParent().asParagraph().copy();
          this.body.removeChild(this.body.findText("«MacroCriterio»").getElement().getParent());
        }

        if(this.body.findText("«SubCriterio»")){
          var subCritParagraph = this.body.findText("«SubCriterio»").getElement().getParent().asParagraph().copy();
          this.body.removeChild(this.body.findText("«SubCriterio»").getElement().getParent());
        }
      }

      var criteri = sheet.getAllCriteri();

      for(var i = 0; i < criteri.length; i++){ 

        if(i === 0){

          this.body.insertParagraph(critId - 1, macroCritParagraph.copy());
          this.replacePlaceholder(criteri[i]);

        }else if(criteri[i].Tipo === "Macro"){

          titoloCriterioId++;

          this.body.insertParagraph(critId - 1, macroCritParagraph.copy());
          this.body.insertParagraph(titoloCriterioId , titoloCriterioParagraph.copy());
          this.replacePlaceholder(criteri[i]);

          critId++;

        }else if(criteri[i].Tipo === "Vuoto"){

          this.body.insertParagraph(critId, "\n")
          critId+= 2;

        }else{

          if(criteri[i].Tipo === "Sub"){

            this.body.insertParagraph(critId, subCritParagraph.copy())
            this.replacePlaceholder(criteri[i]);

          }else{

            if(criteri[i].PuntiCriterio === ""){
              this.body.insertListItem(critId, critParagraph.copy().replaceText('punti', ''));
            }else{
              this.body.insertListItem(critId, critParagraph.copy());
            }
            
            this.replacePlaceholder(criteri[i]); 
          }
          critId++;

        }

      }
    }catch(e){throw e;}
  }


  replaceElencoInsegnamenti(){
    try{
      var sheet = new SheetInsegnamenti();
      var insegnamenti = sheet.getInsegnamenti();
      
      var insParagraph = [], insegnamentoId = 0;
      if(this.body.findText("«INIZIO»") !== null && this.body.findText("«FINE»") !== null){

        var idInizio = this.body.getChildIndex(this.body.findText("«INIZIO»").getElement().getParent());
        var idFine = this.body.getChildIndex(this.body.findText("«FINE»").getElement().getParent());
        insegnamentoId = idInizio + 1;
        
        while(insegnamentoId !== idFine){
          insParagraph.push(this.body.getChild(insegnamentoId).copy());
          insegnamentoId++;
        }

        this.body.removeChild(this.body.getChild(idFine));
        this.body.removeChild(this.body.getChild(idInizio));
        insegnamentoId--;
      }
      

      //per ogni insegnamento
      for(var i = 0; i < insegnamenti.length; i++){

        this.replacePlaceholder(insegnamenti[i]);
        
        insegnamentoId = this.replaceValutazioneCandidati(insegnamenti[i].DescrizioneCopertura, insegnamentoId);
        
        if(parseInt(insParagraph.length) !== 0 && i < insegnamenti.length - 1){        
          for(var t in insParagraph){

            if(insParagraph[t].getType() === DocumentApp.ElementType.PARAGRAPH){
              this.body.insertParagraph(insegnamentoId , insParagraph[t].asParagraph().copy());
            }else if(insParagraph[t].getType() === DocumentApp.ElementType.LIST_ITEM){
              this.body.insertListItem(insegnamentoId , insParagraph[t].asListItem().copy());
            }else if(insParagraph[t].getType() === DocumentApp.ElementType.TABLE){
              this.body.insertTable(insegnamentoId , insParagraph[t].asTable().copy());
            }

            insegnamentoId++;
          }
        }
        
      }
    }catch(e){throw e;}
  }

  replaceValutazioneCandidati(insegnamento, insegnamentoId){
    try{
      var sheet = new SheetValutazioneCandidati(this.sheetVC);
      var candiati = sheet.getCandidati(insegnamento);

      var sheetPunti = new SheetCriteri();
      var datiPunteggi = sheetPunti.getValoriPerTabella()

      var sufficienti = sheet.getSufficienti(candiati, datiPunteggi.Pmin);

      var candParagraph = [], candId = [];
      while(this.body.findText('«NomeCandidato»')){
        candId.push(this.body.getChildIndex(this.body.findText("«NomeCandidato»").getElement().getParent()) + candId.length);
        candParagraph.push(this.body.findText("«NomeCandidato»").getElement().getParent().copy());
        this.body.removeChild(this.body.findText("«NomeCandidato»").getElement().getParent());
      }

      var tabCand = null;
      if(this.body.findText("«Lettera»")){
        tabCand = this.body.findText("«Lettera»").getElement().getParent().getParent().getParentTable().copy();
        var f = tabCand.getChildIndex(tabCand.findText("«Lettera»").getElement().getParent().getParent().getParent());
        var tabRow = tabCand.getChild(f).copy();
        
        this.body.removeChild(this.body.findText("«Lettera»").getElement().getParent().getParent().getParentTable());
      }
      


      for(var i = 0; i < candiati.length; i++){

        if(parseInt(candParagraph.length) !== 0){
          for(var t in candParagraph){
            if(candParagraph[t].getType() === DocumentApp.ElementType.LIST_ITEM){
              this.body.insertListItem(candId[t], candParagraph[t].copy());
            }else{
              this.body.insertParagraph(candId[t], candParagraph[t].copy());
            }
            candId[t]++;
            candId[t]+= parseInt(t);
            insegnamentoId++;
          }
        }

        if(tabCand !== null){
          var tabella = this.body.insertTable(candId[t], tabCand.copy());
          candId[t]+=2;
          this.body.insertParagraph(candId[t], "");
          candId[t]++;
          insegnamentoId++;
        }
        
        this.replacePlaceholder(candiati[i]);

        //inserisco dati tabella
        if(tabCand !== null){
          var punteggi = sheet.getPunteggiCandidato(candiati[i].NomeCandidato);
          var r = f;
          
          for(var t in datiPunteggi['TitoliCriteri']){
            var riga = {
              Lettera: datiPunteggi['LettereCriteri'][t],
              Titolo: datiPunteggi['TitoliCriteri'][t],
              Punti: punteggi[t]
            }
            this.replacePlaceholder(riga);

            if(t < datiPunteggi['TitoliCriteri'].length -1){
              r++;
              tabella.insertTableRow(r, tabRow.copy());
            }
          }
        }
        
      }

      //sufficienti
      if(this.body.findText("«CandidatoSuff»") !== null){
        var suffParagraph = this.body.findText("«CandidatoSuff»").getElement().getParent().copy();
        var suffId = this.body.getChildIndex(this.body.findText("«CandidatoSuff»").getElement().getParent());
        this.body.removeChild(this.body.findText("«CandidatoSuff»").getElement().getParent());
        insegnamentoId--;
        
        for(var j in sufficienti){
          this.body.insertParagraph(suffId, suffParagraph.copy())
          this.replacePlaceholder(sufficienti[j]);
          suffId++;
          insegnamentoId++;
        }
      }
      
      return insegnamentoId;
    }catch(e){throw e;}
  }

  //rimpiazza i tag nel documento
  replacePlaceholder(replacements){
    for (var key in replacements) {
      this.body.replaceText("«" + key + "»", replacements[key]);
    }
  }

}

class Sheet{

  constructor(sheetName){

    try{
      this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      
      if (!this.sheet){
        throw new Error("Il file specificato non esiste o non è accessibile.");
      }
    }catch (e){
      throw new Error("Non è stato possibile accedere al file " + sheetName + ". Controlla il nome del file e riprova."); 
    }
  }

  getData(){
    if(!this.sheet){
      throw new Error("Il foglio non è stato inizializzato correttamente.");
    }
    
    return this.sheet.getDataRange().getDisplayValues();
  }
  
}

class SheetDatiVerbale extends Sheet{

  constructor(){
    super('DatiVerbale');
    this.head = null;
    this.data = null;
    this.initialize();
  }

  initialize(){
    var rawData = this.getData();
    this.head = rawData.shift();
    this.data = rawData;
    
    if(this.head.toString() === "" || this.data.toString() === ""){
      throw new Error("Assenza di dati nel foglio DatiVerbale.")
    }
  }

  getDatiVerbale(){
    var datiVerbale = {};
    for( var j in this.head){
      datiVerbale[this.head[j].toString()] = this.data[0][j].toString();
      datiVerbale[this.head[j].toString().toUpperCase()] = this.data[0][j].toString().toUpperCase();
    }
    return datiVerbale;
  }

  getTemplateUrl(){
    if(!this.head.includes('templateURL')){
      throw new Error("Manca voce 'templateURL' nel foglio DatiVerbale.");
    }else{
      var url = this.data[0][this.head.indexOf('templateURL')].toString();
      return url;
    }
  }
}

class SheetCommissione extends Sheet{

  constructor(){
    super('Commissione');
    this.head = null;
    this.data = null;
    this.presidente = "DA INSERIRE";
    this.segretario = "DA INSERIRE";
    this.initialize();
  }

  initialize(){
    var rawData = this.getData();
    this.head = rawData.shift();
    this.data = rawData;

    if(this.head.toString() === "" || this.data.toString() === ""){
      throw new Error("Assenza di dati nel foglio Commissione.")
    }
  }

  getCommissione(){
    var commissione = [];
    for(var i = 0; i < this.data.length; i++){
      commissione.push(this.getCommissario(i));
    }
    return commissione;
  }

  getCommissario(i){
    var commissario = {};
    for(var j in this.head){
      commissario[this.head[j].toString()] = this.data[i][j].toString();
    }
    
    return commissario;
  }

  getRuoli(){
    if(!this.head.includes("Ruolo")){
      return new Error("Manca la voce 'Ruolo' nel foglio Commissione.")
    }else{

      for(var j in this.head){
        if(this.data[j][this.head.indexOf("Ruolo")].toString() === "Segretario"){
          this.segretario = this.data[j][this.head.indexOf("NomeCommissario")].toString()
        }else if(this.data[j][this.head.indexOf("Ruolo")].toString() === "Presidente"){
          this.presidente = this.data[j][this.head.indexOf("NomeCommissario")].toString()
        }
      }

      var ruoli = {Presidente: this.presidente, Segretario: this.segretario};
      return ruoli;
    }
  }
}

class SheetCriteri extends Sheet{

  constructor(){
    super('TabellaPunteggi');
    this.max = 0;
    this.min = 0;
    this.head = null;
    this.data = null;
    this.initialize();
  }

  initialize(){
    var rawData = this.getData();
    this.max = rawData.shift()[1];
    this.min = rawData.shift()[1];
    this.head = rawData.shift();
    this.data = rawData;
    
    if(this.max.toString() === "" || this.min.toString() === "" || this.head.toString() === "" || this.data.toString() === ""){
      throw new Error("Assenza di dati nel foglio ElencoInsegnamenti.")
    }
  }

  getAllCriteri(){

    if(!this.head.includes('PUNTI')){
      throw new Error("Manca la voce 'PUNTI' nel foglio TabellaPunteggi.");
    }else if(!this.head.includes('CRITERIO')){
      throw new Error("Manca la voce 'CRITERIO' nel foglio TabellaPunteggi.");
    }

    var criteri = [];
    for(var i = 0; i < this.data.length; i++){
      criteri.push(this.getCriterio(i))
    }
    return criteri;
  }

  getCriterio(i){
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
      Pmax: this.max,
      Pmin: this.min
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

    var valoriPerTabella = {TitoliCriteri: titoliCriteri, LettereCriteri: lettereCriteri, Pmin: this.min}
    return valoriPerTabella;
  }
}

class SheetInsegnamenti extends Sheet{

  constructor(){
    super('ElencoInsegnamenti');
    this.head = null;
    this.data = null;
    this.initialize();
  }

  initialize(){
    var rawData = this.getData();
    this.head = rawData.shift();
    this.data = rawData;
    
    if(this.head.toString() === "" || this.data.toString() === ""){
      throw new Error("Assenza di dati nel foglio ElencoInsegnamenti.")
    }
  }

  getInsegnamenti(){
    var insegnamenti = [];
    for(var i = 0; i < this.data.length; i++){
      insegnamenti.push(this.getInsegnamento(i));
    }
    return insegnamenti;
  }

  getInsegnamento(i){
    var insegnamento = {};
    for(var j in this.head){
      insegnamento[this.head[j].toString()] = this.data[i][j].toString();
    }
    return insegnamento;
  }

}
class SheetValutazioneCandidati extends Sheet{

  constructor(ValutazioneCandidati){
    super(ValutazioneCandidati);
    this.head = null;
    this.data = null;
    this.initialize();
  }

  initialize(){
    var rawData = this.getData();
    this.head = rawData.shift();
    this.data = rawData;
    
    if(this.head.toString() === "" || this.data.toString() === ""){
      throw new Error("Assenza di dati nel foglio ValutazioneCandidati scelto.")
    }
  }

  getCandidati(insegnamento){
    if(!this.head.includes('Insegnamento')){
      throw new Error("Manca la voce 'Insegnamento' nel foglio ValutazioneCandidati scelto.");
    }else{
      var candidati = [];
      for(var i = 0; i < this.data.length; i++){
        if(this.data[i][this.head.indexOf('Insegnamento')].toString() === insegnamento){
          candidati.push(this.getCandidato(i));
        }
      }
      return candidati;
    }
  }

  getCandidato(i){
    var candidato = {};
    for(var j in this.head){
      candidato[this.head[j].toString()] = this.data[i][j].toString();
    }
    return candidato;
  }

  getSufficienti(candidati, pmin){
    try{
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
      
    }catch(e){ throw new Error("Mancano alcune voci nel foglio ValutazioneCandidati scelto.")}
  }
  
  getPunteggiCandidato(nomeCandidato){
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