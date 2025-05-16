function onOpen(): void {
  const menu = SpreadsheetApp.getUi().createMenu("Funzioni Aggiuntive");
  menu.addItem("Genera Verbale", "generaVerbale");
  menu.addToUi();
}