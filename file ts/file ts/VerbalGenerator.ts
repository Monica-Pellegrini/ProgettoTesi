class VerbalGenerator {
  private activeSheetName: string;
  private linkTemplate: string;
  private interface: UserInterface;

  constructor() {
    this.activeSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    this.initialize();
  }

  private initialize(): void {
    this.interface = new UserInterface();
    const sheet: SheetDatiVerbale = new SheetDatiVerbale();
    this.linkTemplate = sheet.getTemplateUrl();
  }

  public generaVerbale(): void {
    if (this.linkTemplate !== "") {
      if (this.linkTemplate.match("https://docs.google.com/document/")) {
        if (this.activeSheetName.match("ValutazioneCandidati")) {

          const timestamp: string = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm:ss");
          let templateId: string = "";
          try {
            templateId = DocumentApp.openByUrl(this.linkTemplate).getId();
          } catch (e) {
            throw new Error("Non è stato possibile accedere al documento template.");
          }

          const verbale: Verbale = new Verbale(templateId, timestamp, this.activeSheetName);
          verbale.replaceAll();

          this.interface.makeOutputBox(verbale.getLink(), verbale.getName());
        } else {
          SpreadsheetApp.getUi().alert('ATTENZIONE!\nAvviare la funzione da un foglio contenente "ValutazioneCandidati" nel nome.');
        }
      } else {
        SpreadsheetApp.getUi().alert("ATTENZIONE!\nInserire un URL valido.");
      }
    } else {
      SpreadsheetApp.getUi().alert("ATTENZIONE!\nInserire un URL valido.");
    }
  }
}