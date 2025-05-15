class VerbalGenerator{
  private activeSheetName: string;
  private linkTemplate: string;
  private interface: userInterface;
  private input: any;

  constructor(){
    this.activeSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    this.linkTemplate = "";
    this.input = null;
    this.initialize()
  }

  initialize(): void{
    this.interface = new userInterface();
    this.input = this.interface.makeInputBox();
    this.linkTemplate = this.input.Url;
  }

  generaVerbale(): void{
    if(this.input.Button == SpreadsheetApp.getUi().Button.OK){
      if(this.linkTemplate !== ""){
        if(this.linkTemplate.includes("https://docs.google.com/document/")){
          if(this.activeSheetName.includes("ValutazioneCandidati")){

            var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
            var templateId = DocumentApp.openByUrl(this.linkTemplate).getId();
            var verbale = new Verbale(templateId, timestamp, this.activeSheetName);
            verbale.replaceAll();


            this.interface.makeOutputBox(verbale.getLink(),verbale.getName());

          }else{
            SpreadsheetApp.getUi().alert('ATTENZIONE!\nAvviare la funzione da un foglio contenente "ValutazioneCandidati" nel nome.')
          }
        }else{
          SpreadsheetApp.getUi().alert("ATTENZIONE!\nInserire un URL valido.")
        }
      }
    }else{}



  }

}