function generaVerbale(): void {
  try {
    const verbale: VerbalGenerator = new VerbalGenerator();
    verbale.generaVerbale();
  } catch (e: unknown) {
    SpreadsheetApp.getUi().alert("Errore", String(e), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}