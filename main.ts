function main()
{
  //let nameManager = new nameManager()
  //let nomeScoresSheet = nameManager.getNamePunteggi();
  //let nomeCandidatesSheet = nameManager.getNameCandidati();
  //let a = new GenerateValutazioneCandidati(nomeScoresSheet, nomeCandidatesSheet);

  let a = new GenerateValutazioneCandidati('Punteggi', 'ElencoCandidati');
  a.generateValutazione();
}