function onOpen(){
  const menu = SpreadsheetApp.getUi().createMenu("Funzioni Aggiuntive");
  menu.addItem("Genera Valutazione Candidati", "main");
  menu.addToUi();
  menu.addItem("Genera Verbale", "generaVerbale");
  menu.addToUi();
}