
function onOpen(): void
{
  let menu = SpreadsheetApp.getUi().createMenu("Funzioni Aggiuntive");
  menu.addItem("Genera Valutazione Candidati", "main");
  menu.addToUi();
}

function main(): void
{
  const ui = SpreadsheetApp.getUi();
  ui.alert("Ciao! Questa è una semplice notifica.");
}