class Verbale{
    private verbaleId: string;
    private sheetVC: string;
    private body: Body;

  constructor(templateId: string, timestamp: string, sheetVC: string){
    this.verbaleId = DriveApp.getFileById(templateId).makeCopy("Verbale " + timestamp).getId();
    this.body = DocumentApp.openById(this.verbaleId).getBody();
    this.sheetVC = sheetVC;
  }

  getLink(): string{
    return DocumentApp.openById(this.verbaleId).getUrl();
  }

  getName(): string{
    return DocumentApp.openById(this.verbaleId).getName();
  }

  replaceAll(): void{
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

  replaceDatiVerbale(): void{
    var sheet = new SheetDatiVerbale();

    var datiVerbale = sheet.getDatiVerbale();
    
    this.replacePlaceholder(datiVerbale); 
  }

  replaceCommissione(): void{
    const sheet = new SheetCommissione();


    var paragraphsId = [], paragraphs = [], allPar = [];
    allPar = this.body.getParagraphs();
    for(var i in allPar){
      if(allPar[i].getText().toString().match("«NomeCommissario»") && allPar[i].getParent().getType() === DocumentApp.ElementType.BODY_SECTION){
        paragraphsId.push(this.body.getChildIndex(allPar[i]));
        paragraphs.push(allPar[i].copy());
      }
    }


    var allTabs = this.body.getTables();
    var tables = [];
    for( var j in allTabs){
      if(allTabs[j].getCell(0,0).getText() === '«NomeCommissario»'){
        tables.push(allTabs[j]);
        var tRow = allTabs[j].getRow(0).copy();
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
          tables[j].appendTableRow(tRow.copy());
        }
      }
    }
    
    var ruoli = sheet.getRuoli();

    this.replacePlaceholder(ruoli);
  }

  replaceElencoPunteggi(): void{
    var sheet = new SheetCriteri();

    this.replacePlaceholder(sheet.getPuntMaxMin());

    var titoloCriterioId = this.body.getChildIndex(this.body.findText("«TitoloCriterio»").getElement().getParent());
    var titoloCriterioParagraph = this.body.getChild(titoloCriterioId).asParagraph().copy();

    var macroCritParagraph = this.body.findText("«MacroCriterio»").getElement().getParent().asParagraph().copy();
    var critParagraph = this.body.findText("«Criterio»").getElement().getParent().asListItem().copy();
    var subCritParagraph = this.body.findText("«SubCriterio»").getElement().getParent().asParagraph().copy();
    var critId = this.body.getChildIndex(this.body.findText("«Criterio»").getElement().getParent());

    this.body.removeChild(this.body.findText("«MacroCriterio»").getElement().getParent());
    this.body.removeChild(this.body.findText("«Criterio»").getElement().getParent());
    this.body.removeChild(this.body.findText("«SubCriterio»").getElement().getParent());

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
        critId++;
        this.body.insertParagraph(critId, "\n")
        critId++;

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

    };

  }


  replaceElencoInsegnamenti(): void{
    var sheet = new SheetInsegnamenti();
    var insegnamenti = sheet.getInsegnamenti();
    
    var insParagraph = [];
    var idInizio = this.body.getChildIndex(this.body.findText("«INIZIO»").getElement().getParent());
    var idFine = this.body.getChildIndex(this.body.findText("«FINE»").getElement().getParent());
    var insegnamentoId = idInizio + 1;
    while(insegnamentoId !== idFine){
      insParagraph.push(this.body.getChild(insegnamentoId).copy());
      insegnamentoId++;
    }

    this.body.removeChild(this.body.getChild(idFine));
    this.body.removeChild(this.body.getChild(idInizio));
    insegnamentoId--;

    //per ogni insegnamento
    for(var i = 0; i < insegnamenti.length; i++){

      this.replacePlaceholder(insegnamenti[i]);
      
      insegnamentoId = this.replaceValutazioneCandidati(insegnamenti[i].DescrizioneCopertura, insegnamentoId);

      if(i < insegnamenti.length - 1){        
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
  }

  replaceValutazioneCandidati(insegnamento, insegnamentoId): string{
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

    var tabCand = this.body.findText("«Lettera»").getElement().getParent().getParent().getParentTable().copy();
    var tabRow = tabCand.getRow(1).copy();
    this.body.removeChild(this.body.findText("«Lettera»").getElement().getParent().getParent().getParentTable());




    for(var i = 0; i < candiati.length; i++){

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

        var tabella = this.body.insertTable(candId[t], tabCand.copy());
        candId[t]+=2;
        this.body.insertParagraph(candId[t], "");
        candId[t]++;
        insegnamentoId++;

        this.replacePlaceholder(candiati[i]);

        //inserisco dati tabella

        var punteggi = sheet.getPunteggiCandidato(candiati[i].NomeCandidato);
        
        var f = 2
        for(var t in datiPunteggi['TitoliCriteri']){
          var riga = {
            Lettera: datiPunteggi['LettereCriteri'][t],
            Titolo: datiPunteggi['TitoliCriteri'][t],
            Punti: punteggi[t]
          }
          this.replacePlaceholder(riga);

          if(t < datiPunteggi['TitoliCriteri'].length -1){
            tabella.insertTableRow(f, tabRow.copy());
            f++;
          }
        }

    }

    //sufficienti
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

    return insegnamentoId;
  }

  //rimpiazza i tag nel documento
  replacePlaceholder(replacements): void{
    for (var key in replacements) {
      this.body.replaceText("«" + key + "»", replacements[key]);
    }
  }
}