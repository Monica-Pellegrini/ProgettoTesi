//class that asks for the name of the sheet to open
class nameManager
{
  private nome_scores_sheet: string;
  private nome_candidates_sheet: string;
  
  constructor()
  {
    this.nome_scores_sheet = Browser.inputBox("Inserire il nome dello sheet contentente la tabella PUNTEGGI");
    this.nome_candidates_sheet = Browser.inputBox("Inserire il nome dello sheet contentente la tabella CANDIDATI");
  }

  getNamePunteggi()
  {
    return this.nome_scores_sheet;
  }

  getNameCandidati()
  {
    return this.nome_candidates_sheet;
  }
}