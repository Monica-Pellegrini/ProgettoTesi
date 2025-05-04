//rimane da fare il controllo degli errori
var i, j, k, h;

//stile titolo
const stileTitolo = {};
stileTitolo[DocumentApp.Attribute.BOLD] = true;
stileTitolo[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.JUSTIFY;
stileTitolo[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
stileTitolo[DocumentApp.Attribute.FONT_SIZE] = 11;
stileTitolo[DocumentApp.Attribute.LINE_SPACING] = 1;

//stile corpo
const stileBody = {};
stileBody[DocumentApp.Attribute.BOLD] = false;
stileBody[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.JUSTIFY;
stileBody[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
stileBody[DocumentApp.Attribute.FONT_SIZE] = 11;
stileBody[DocumentApp.Attribute.LINE_SPACING] = 1;

//stile tabella
const stileTab = {};
stileTab[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;

//classe che estrae i valori dei fogli
class ValueRetriver{
  
  constructor(){
    this.sheet;
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  getVal(sheetName){
    this.sheet = this.spreadsheet.getSheetByName(sheetName); 
    return this.sheet.getDataRange().getDisplayValues();
  }

}


//classe candidato
class Candidato{

  constructor(tabVal, i){
    this.nome = tabVal[i][0].toString();
    this.insegnamento = tabVal[i][1].toString();
    this.punteggioTot = tabVal[i][2].valueOf();

    this.totCriterio = [];
    for (k = 4; k < tabVal[0].length; k++){
      this.s = tabVal[0][k].toString();
      if(this.s.match("Tot")){
        this.totCriterio.push(tabVal[i][k].valueOf());
      }
    }
  }

  getNome(){
    return this.nome;
  }

  getInsegnamento(){
    return this.insegnamento;
  }

  getPunteggioTot(){
    return this.punteggioTot;
  }

  getTotCriterio(i){
    return this.totCriterio[i];
  }

  getTotCriterioLenght(){
    return this.totCriterio.length;
  }
}

//classe insegnamento
class Insegnamento{

  constructor(tabIns, i){
    this.idCop = tabIns[i][0].toString();
    this.descrizioneCop = tabIns[i][1].toString();
    this.ssd = tabIns[i][2].toString();
    this.oreLezione = tabIns[i][3].toString();  
  }

  getInsIDCop(){
    return this.idCop;
  }

  getInsDescrizioneCop(){
    return this.descrizioneCop;
  }

  getInsSSD(){
    return this.ssd;
  }

  getInsOreLezione(){
    return this.oreLezione;
  }
}

//classe commissionario
class Commissionario{
  constructor(tabCandidati, i){
    this.nomeC = tabCandidati[i][0].toString();
    this.titoloC = tabCandidati[i][1].toString();
    this.ruoloC = "nulla";
    if(tabCandidati[i][2][0] !== undefined && tabCandidati[i][2][0] !== " "){
      this.ruoloC = tabCandidati[i][2].toString();
    }

  }

  getNomeC(){
    return this.nomeC;
  }

  getTitoloC(){
    return this.titoloC;
  }

  getRuoloC(){
    return this.ruoloC;
  }
}

//classe che contiene le informazioni da inserire nel verbale
class DatiVerbale{

  constructor(){

    this.v = new ValueRetriver();


    //carichiamo i dati dallo sheet DatiVerbale
    this.tabDati = this.v.getVal("DatiVerbale");

    //carichiamo i dati dallo sheet ElencoInsegnamenti
    this.tabInsegnamenti = this.v.getVal("ElencoInsegnamenti");

    //carichiamo i dati dallo sheet TabellaPunteggi
    this.tabPunteggi = this.v.getVal("TabellaPunteggi");

    //carichiamo i dati dallo sheet ValutazioneCandidati corrente
    this.tabVal = this.v.getVal(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName());

    //carichiamo i dati dallo sheet Commissione
    this.tabCommissione = this.v.getVal("Commissione");


    this.cds = this.tabDati[1][0].toString();
    this.dataAffissione = this.tabDati[1][1].toString();
    this.giorno = this.tabDati[1][2].toString();
    this.oraInizio = this.tabDati[1][3].toString();

    this.insegnamenti = [];
    for(i = 1; i < this.tabInsegnamenti.length; i++){
      this.insegnamenti[i-1] = new Insegnamento(this.tabInsegnamenti, i);
    }

    this.criteri = this.tabPunteggi;
    this.pMax = this.tabPunteggi[0][1].valueOf();
    this.pMin = this.tabPunteggi[1][1].valueOf();

    this.candidati = [];
    for(i = 1; i < this.tabVal.length; i++){
      this.candidati[i-1] = new Candidato(this.tabVal, i);
    }
    
    this.commissione = [];
    for(i = 1; i < this.tabCommissione.length; i++){
      this.commissione[i-1] = new Commissionario(this.tabCommissione, i);

      if(this.commissione[i-1].getRuoloC() === 'segretario'){
        this.segretario = this.commissione[i-1].getNomeC();
      }else if(this.commissione[i-1].getRuoloC() === 'presidente'){
        this.presidente = this.commissione[i-1].getNomeC();        
      }
    }

  }

  getCDS(){
    return this.cds;
  }

  getDataAffissione(){
    return this.dataAffissione;
  }

  getGiorno(){
    return this.giorno;
  }

  getOraInizio(){
    return this.oraInizio;
  }

  getIDCop(i){
    return this.insegnamenti[i].getInsIDCop();
  }

  getDescrizioneCop(i){
    return this.insegnamenti[i].getInsDescrizioneCop();
  }

  getSSD(i){
    return this.insegnamenti[i].getInsSSD();
  }

  getOreLezione(i){
    return this.insegnamenti[i].getInsOreLezione();
  }

  getInsegnamentiLenght(){
    return this.insegnamenti.length;
  }

  getCriteri(i){
    return this.criteri[i][1].toString();
  }

  getCriteriLenght(){
    return this.criteri.length;
  }

  getPMax(){
    return this.pMax;
  }

  getPMin(){
    return this.pMin;
  }

  getNomeC(i){
    return this.commissione[i].getNomeC();
  }

  getTitoloC(i){
    return this.commissione[i].getTitoloC();
  }

  getRuoloC(i){
    return this.commissione[i].getRuoloC();
  }

  getCandidato(i){
    return this.candidati[i];
  }

  getCandidatiLenght(){
    return this.candidati.length;
  }

  getCommissione(i){
    return this.commissione[i];
  }

  getCommissioneLenght(){
    return this.commissione.length;
  }

  getSegretario(){
    return this.segretario;
  }

  getPresidente(){
    return this.presidente;
  }
}

//classe che crea e scrive il verbale
class Doc{

  //crea il verbale ed estrae le informazioni dai fogli
  constructor(){
    this.dataCorrente = new Date();
    this.timestamp = Utilities.formatDate(this.dataCorrente, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
    this.doc = DocumentApp.create("Verbale "+this.timestamp);
    this.dati = new DatiVerbale();
  }

  //scrive il verbale
  fillDoc(){
    this.body = this.doc.getBody();
    this.body.setMarginLeft(28);
    this.body.setMarginRight(28);
    this.body.setMarginTop(28);

    //tabella firme
    this.firme = [];
    for(i=0; i < this.dati.getCommissioneLenght(); i++){
      this.firme.push([this.dati.getNomeC(i)+'\n', 'Firmato digitalmente\n'])
    }
    
    this.titolo = "SELEZIONE PUBBLICA, PER TITOLI, PER IL CONFERIMENTO DI INCARICHI DI INSEGNAMENTO PER IL CORSO DI LAUREA IN "+ this.dati.getCDS().toUpperCase() +" PRESSO IL DIPARTIMENTO DI INGEGNERIA (AVVISO AFFISSO ALL'ALBO DELL'UNIVERSITÀ DEGLI STUDI DI FERRARA IL "+ this.dati.getDataAffissione().toUpperCase()+")";
    this.body.insertParagraph(0,this.titolo).setAttributes(stileTitolo);

    this.st = "\nIl giorno "+this.dati.getGiorno()+" alle ore "+this.dati.getOraInizio()+" presso il Dipartimento di Ingegneria si è riunita la Commissione giudicatrice della selezione pubblica per titolo per il conferimento di incarichi degli incarichi di insegnamento del Corso di laurea in "+this.dati.getCDS()+" (presso il Dipartimento di Ingegneria, avviso affisso all'Albo dell'Università degli studi di Ferrara il "+this.dati.getDataAffissione()+") così composta:";
    this.body.appendParagraph(this.st).setAttributes(stileBody);

    this.comm = "";
    for(i=0; i < this.dati.getCommissioneLenght(); i++){
      this.comm = this.comm + "\n"+ this.dati.getNomeC(i) +",  "+ this.dati.getTitoloC(i) +" presso l'Università di Ferrara";
    }
    this.body.appendParagraph(this.comm).setAttributes(stileBody);

    this.st = "\nÈ stato designato Presidente: " + this.dati.getPresidente() +
    "\nLe funzioni di Segretario sono state assunte da " + this.dati.getSegretario() +
    "\n\nLa Commissione, presa visione dell’avviso, prende atto che costituiscono titoli preferenziali per il conferimento dell’incarico di insegnamento il possesso del titolo di dottore di ricerca, della specializzazione medica, dell’abilitazione, ovvero di titoli equivalenti conseguiti all’estero e che l’avviso prevede che i titoli valutabili siano i seguenti:\n";
    this.body.appendParagraph(this.st).setAttributes(stileBody);

    //elenca i criteri principali
    this.titCriteri = [];
    this.t = 0;
    this.st = '';
    for(i = 3; i < this.dati.getCriteriLenght(); i++){
      if(this.dati.getCriteri(i)[0] !== ' ' && this.dati.getCriteri(i)[0] !== undefined){
        this.st = this.st + this.dati.getCriteri(i) + "\n";
        this.titCriteri[this.t] = this.dati.getCriteri(i).toString();
        this.t++;
      }
    }
    this.body.appendParagraph(this.st).setAttributes(stileBody);

    this.st = "\nAlla valutazione dei titoli sono riservati "+this.dati.getPMax()+" punti. Gli incarichi sono conferiti, entro il numero di quelli messi a selezione, ai candidati che abbiano conseguito almeno "+this.dati.getPMin()+" dei "+this.dati.getPMax()+" punti complessivamente a disposizione secondo l'ordine della graduatoria stessa.\n"+
    "\nI criteri definiti dalla Commissione per l'attribuzione dei punteggi per i titoli sono indicati nell’Allegato 1 al presente verbale che ne costituisce parte integrante e sostanziale.\n";
    this.body.appendParagraph(this.st).setAttributes(stileBody);

    for(h = 0; h < this.dati.getInsegnamentiLenght(); h++){

      this.st = "\nLa Commissione ha preso a questo punto in esame attraverso la piattaforma PICA le domande dei candidati per l’insegnamento di:\n"+
      "\nID_Copertura: "+this.dati.getIDCop(h)+
      "\nDescrizione copertura: "+this.dati.getDescrizioneCop(h)+
      "\nCorso di studio: "+this.dati.getCDS()+
      "\nPartizione studenti:"+
      "\nSSD: "+this.dati.getSSD(h)+
      "\nOre lezione: "+this.dati.getOreLezione(h)+
      "\nI candidati iscritti risultano essere:\n";
      this.body.appendParagraph(this.st).setAttributes(stileBody);


      //elenca i candidati
      for(i = 0; i < this.dati.getCandidatiLenght(); i++){
        if(this.dati.getCandidato(i).getInsegnamento().toUpperCase() === this.dati.getDescrizioneCop(h).toUpperCase()){
          this.body.appendListItem(this.dati.getCandidato(i).getNome()).setGlyphType(DocumentApp.GlyphType.NUMBER);
        }
      }

      this.st = "\nLa Commissione ha constatato l'assenza tra i suoi membri e tra questi ed i concorrenti dell'incompatibilità di cui al secondo comma dell'art.5 del D.L. 7.5.1948, n.1172. Ognuno dei membri dichiara, altresì, che non sussistono le cause di astensione di cui all'art. 51 c.p.c.\n"+
      "\nLa Commissione ha quindi proceduto alla valutazione dei curricula prodotti dai candidati in conformità ai criteri sopra definiti."+
      "\nÈ stato quindi assegnato a ciascun concorrente il seguente punteggio:";
      this.body.appendParagraph(this.st).setAttributes(stileBody);


      //array dei sufficienti
      this.suf = [];
      
      //elenco delle tabelle
      for(i = 0; i < this.dati.getCandidatiLenght(); i++){

        if(this.dati.getCandidato(i).getInsegnamento().toUpperCase() === this.dati.getDescrizioneCop(h).toUpperCase()){

          //check se è sufficiente
          if(this.dati.getCandidato(i).getPunteggioTot() > (this.dati.getPMin()-1)){
            this.suf.push(this.dati.getCandidato(i).getNome());
          }

          this.st = "\n- Dott. "+ this.dati.getCandidato(i).getNome()+" complessivi punti  "+ this.dati.getCandidato(i).getPunteggioTot() +"/"+this.dati.getPMax()+" di cui:\n";

          this.body.appendParagraph(this.st).setAttributes(stileBody);

          //crea tabella
          this.tabella = [
          ['Categoria titoli', 'Titolo presentato', 'Punteggio']];

          for(j = 0; j < this.titCriteri.length; j++){
            this.tabella.push([this.titCriteri[j][0],this.titCriteri[j],this.dati.getCandidato(i).getTotCriterio(j)]);
          }

          this.tabella.push(['','TOTALE',this.dati.getCandidato(i).getPunteggioTot()]);

          this.tab = this.body.appendTable(this.tabella).setAttributes(stileTab);


          //sistema padding e altro
          this.tab.setColumnWidth(0,70).setColumnWidth(1,150).setColumnWidth(2,70);
          
          this.rows = this.tab.getNumRows();
          this.cols = this.tab.getChild(0).asTableRow().getNumChildren();

          for(j = 0 ; j < this.rows; j++)
          {
            for(k = 0; k < this.cols; k++)
            {
              this.tab.getCell(j,k).setPaddingBottom(0).setPaddingTop(0);
            }      
          }
        }
        
      }

      this.st = "\nI candidati che hanno ottenuto un punteggio uguale o superiore a "+this.dati.getPMin()+"/"+this.dati.getPMax()+" sono quindi (in ordine alfabetico):\n";
      this.body.appendParagraph(this.st).setAttributes(stileBody);

      //elenca i sufficienti in ordine
      this.suf.sort();
      for(i=0; i < this.suf.length; i++){
        this.st = "- " + this.suf[i];
        this.body.appendParagraph(this.st).setAttributes(stileBody);
      }
    }
    

    this.st = "\n\nIl risultato della valutazione dei titoli viene inviato al Direttore di Dipartimento per l’approvazione della graduatoria in Consiglio e la successiva pubblicazione sul sito web del Dipartimento.\n"+
    "\nLa riunione ha avuto termine alle ore _________________\n";
    this.body.appendParagraph(this.st).setAttributes(stileBody);

    this.st = "\nLA COMMISSIONE\n\n";
    this.body.appendParagraph(this.st).setAttributes(stileBody).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    this.tab1 = this.body.appendTable(this.firme).setAttributes(stileTab).setBorderColor('#FFFFFF');

    //sistemo l'allinemaneto della tabella firme
    this.rows = this.tab1.getNumRows();
    this.cols = this.tab1.getChild(0).asTableRow().getNumChildren();

    for(j = 0 ; j < this.rows; j++)
    {
      for(k = 0; k < this.cols; k++)
      {
        this.tab1.getCell(j,k).getChild(0).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }      
    }

    this.body.appendPageBreak();




    //seconda pagina
    this.body.appendParagraph(this.titolo).setAttributes(stileTitolo);

    this.st = "\nALLEGATO 1 - CRITERI DI VALUTAZIONE";
    this.body.appendParagraph(this.st).setAttributes(stileTitolo).setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    this.st = "\n\nDopo ampia ed approfondita discussione, la Commissione giudicatrice della procedura di selezione per il conferimento degli incarichi di insegnamento del Corso di laurea in "+this.dati.getCDS()+" presso il Dipartimento di Ingegneria così composta:\n"+ this.comm +
    "\ndelibera l'attribuzione dei punteggi per i titoli con i seguenti criteri:\n";
    this.body.appendParagraph(this.st).setAttributes(stileBody);

    //elenca i criteri in dettaglio
    for(i = 3; i < this.dati.getCriteriLenght(); i++){
      this.body.appendParagraph(this.dati.getCriteri(i)).setAttributes(stileBody);
    }

    this.st = "\n\nLA COMMISSIONE\n\n";
    this.body.appendParagraph(this.st).setAttributes(stileBody).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    this.tab2 = this.body.appendTable(this.firme).setAttributes(stileTab).setBorderColor('#FFFFFF');

    //sistemo l'allinemaneto della tabella firme
    this.rows = this.tab2.getNumRows();
    this.cols = this.tab2.getChild(0).asTableRow().getNumChildren();

    for(j = 0 ; j < this.rows; j++)
    {
      for(k = 0; k < this.cols; k++)
      {
        this.tab2.getCell(j,k).getChild(0).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }      
    }

    this.body.appendPageBreak();



    //terza pagina
    this.body.appendParagraph(this.titolo).setAttributes(stileTitolo);

    this.st = "\n\nIl/La sottoscritt_, Prof./Prof.ssa ______________________, componente della commissione giudicatrice della selezione pubblica, per titoli, per il conferimento di incarichi di insegnamento per il corso di laurea in  "+this.dati.getCDS()+" presso il Dipartimento di Ingegneria (avviso affisso all'albo dell'università degli studi di Ferrara il "+this.dati.getDataAffissione()+") dichiara di aver partecipato, per via telematica, alla seduta della Commissione del "+this.dati.getGiorno()+".\n"+
    "\nDichiara inoltre di concordare con il verbale a firma degli altri membri della Commissione.\n"+
    "\n\n_____________, lì  " + this.dati.getGiorno() + "\n";
    this.body.appendParagraph(this.st).setAttributes(stileBody);

    this.st = "_______________________";
    this.body.appendParagraph(this.st).setAttributes(stileBody).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  }

  getDocUrl(){
    return this.doc.getUrl();
  }

  getDocName(){
    return this.doc.getName();
  }
}


function genereVerbale(){

  if(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName().match('ValutazioneCandidati')){
    let d = new Doc();
  d.fillDoc();

  //crea box con link cliccabile
  let link = d.getDocUrl();
  var desc = d.getDocName();
  var html=Utilities.formatString('<style>input{margin: 3px 0;}</style><h3>%s</h3><a href="%s" target="_blank">%s</a><br />','Link del verbale:',link,desc);
  var userInterface=HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModelessDialog(userInterface, "Verbale creato");

  }else{
    Browser.msgBox("Eseguire lo script su un foglio ValutazioneCandidati");
  }
  
}
