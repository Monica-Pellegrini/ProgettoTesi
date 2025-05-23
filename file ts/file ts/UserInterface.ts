class UserInterface {
  constructor() {}

  makeOutputBox(link: string, nome: string): void {
    // crea box con link cliccabile
    const html: string = Utilities.formatString('<style>input{margin: 3px 0;}</style><h3>%s</h3><a href="%s" target="_blank">%s</a><br />', 'Link del verbale:', link, nome);
    const userInterface: GoogleAppsScript.HTML.HtmlOutput = HtmlService.createHtmlOutput(html).setHeight(100);
    SpreadsheetApp.getUi().showModelessDialog(userInterface, "Verbale creato");
  }
}