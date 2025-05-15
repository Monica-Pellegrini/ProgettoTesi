class userInterface{
  constructor(){}

  makeInputBox(){
    var result = SpreadsheetApp.getUi().prompt('Setup','Prima di iniziare, assicurati che i nomi dei fogli da cui vuoi prendere i seguenti dati coincidano coi seguenti:\n\n- "DatiVerbale" per i dati della selezione\n- "ElencoInsegnamenti" per gli insegnamenti\n- "TabellaPunteggi" per i criteri di valutazione e i loro punteggi\n- "Commissione" per la commissione\n\nAssicurati che il nome del foglio attivo includa "ValutazioneCandidati" nel nome.\n\nInserire l\'URL del documento da usare come template:',SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);

    var button = result.getSelectedButton();
    var url = result.getResponseText();
    var input = {Button: button, Url: url};
    return input;
  }

  makeOutputBox(link: string, nome: string): void{
    //crea box con link cliccabile
    var html=Utilities.formatString('<style>input{margin: 3px 0;}</style><h3>%s</h3><a href="%s" target="_blank">%s</a><br />','Link del verbale:',link,nome);
    var userInterface=HtmlService.createHtmlOutput(html).setHeight(100);
    SpreadsheetApp.getUi().showModelessDialog(userInterface, "Verbale creato");
  }
}