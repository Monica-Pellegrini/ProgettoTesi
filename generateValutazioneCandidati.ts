//class for generating the valutazione candidati
class GenerateValutazioneCandidati
{
  private spreadsheet: SpreadsheetCreator;
  private punteggiData: SheetPunteggi;
  private candidatiData: SheetCandidati;
  private timestamp: string | null;
  private namePunteggiSheet: string;

  constructor(namePunteggiSheet: string, nameCandidatiSheet: string)
  {
    this.spreadsheet = new SpreadsheetCreator();
    this.punteggiData = new SheetPunteggi(namePunteggiSheet);
    this.candidatiData = new SheetCandidati(nameCandidatiSheet);
    this.timestamp = null;
    this.namePunteggiSheet = namePunteggiSheet;
  }
  

  //method that generates the Valutazione Candidati header
  generateHeader(tabPunteggi: any[][]): string[]
  {
    let headers = ["NomeCandidato", "Insegnamento", "PunteggioTotale", "Sufficiente"]; //add the first 4
    let count = 0;

    for(let i = 1; i < tabPunteggi.length; i++)
    {
      //if we are not in a section separator line and are not in a range line
      if(tabPunteggi[i][1] !== "" && !tabPunteggi[i][1].startsWith("-")) 
      {
        headers.push(tabPunteggi[i][1]);

        headers.push("Voce");
        headers.push("Punti " + tabPunteggi[i][1]);
      }
      else if(!tabPunteggi[i][1].startsWith("-")) //I have reached the end of the section
      {
        count++;
        headers.push("Tot"+ count);
        i++; //I skip the next section heading
      }
    } 

    count++;
    headers.push("Tot"+ count);
    return headers;
  }

  //method that saves the maximum scores of each section
  getMax(tabPunteggi: any[][]): number[]
  {
    let max = [];
    max.push(tabPunteggi[0][2]) //Insert the max value of the first section
    
    for(let i = 1; i < tabPunteggi.length; i++)
    {
      if(tabPunteggi[i][1] === "")
      {
        max.push(tabPunteggi[i+1][2]); //inserst the max value of the next sections
      }
    }
    
    return max;
  }

  //method that saves the scores for each criterion
  getScores(tabPunteggi: any[][]): number[]
  {
    let scores = [];

    for(let i = 1; i < tabPunteggi.length; i++)
    {
      if(tabPunteggi[i][1] !== "" && !tabPunteggi[i][1].startsWith("-"))
      {
        scores.push(tabPunteggi[i][2]); //I save the scores for the various criteria
      }
      else if(!tabPunteggi[i][1].startsWith("-"))//I have reached the end of the section
      {
        i++; //I skip the section header because I don't want to save the maximum score here
      }
    }

    return scores;
  }

  //method that saves the formulas of each criterion
  getFormulas(tabPunteggi: any[][]): string[]
  {
    let formulas = [];
    
    for(let i = 1; i < tabPunteggi.length; i++)
    {
      if(tabPunteggi[i][1] !== "" && !tabPunteggi[i][1].startsWith("-"))
      {
        formulas.push(tabPunteggi[i][0].trim()); //save the value of the formula
      }
      else if(!tabPunteggi[i][1].startsWith("-"))//reached the end of the section
      {
        i++; //I skip the header where there is no formula
      }
    }
    
    return formulas;
  }

  //method that stores the indices from which each range begins
  getIndicesRange(tabPunteggi: any[][], tabCandidati: any[][]): number[]
  {
    let numColumns = tabPunteggi.length;
    let numRow = tabCandidati.length;
    let indicesRange = [];

    for(let i = 1; i < numColumns; i++)
    {
      if(tabPunteggi[i][1] !== "" && !tabPunteggi[i][1].startsWith("-")) 
      {
        if(tabPunteggi[i][0].trim() === "RANGE")
        {
          //need to figure out where the column range starts from
          for(var a = 0; a < numRow; a++)
          {
            //enter the formula for all candidates
            indicesRange.push(i+1); //saving the index of the first entry "-"
          }
        }
      }
    }

    return indicesRange;
  }

  checkMaxScore(maxArray: number[]): void
  {
    let maxScore = this.punteggiData.getMaximumScore(); //get the maximum score possible 
    let somma = maxArray.reduce(function(total, num) {
      return total + num;
    }, 0);

    if(somma !== maxScore)
    {
      throw new Error("Errore: la il valore massimo dei punti ottenibili dal candidato inserito nella tabella punteggi: " + maxScore + " non corrisponde alla somma dei valori massimi delle categorie che è: " + somma +
      " ricontrollare e poi rieseguire il servizio");
    }
  }

  generateValutazione(): void
  {
    const tabPunteggi = this.punteggiData.getPunteggiData(); //I read and save the data from the score table
    const minimum = this.punteggiData.getMinimumScore(); //Get the value of the minimum score accepted 
    const tabCandidati = this.candidatiData.getCandidatiData(); //I read and save the data from the candidate table
    const headers = this.generateHeader(tabPunteggi); //generate the headers for the ValutazioneCandidati
    const max = this.getMax(tabPunteggi); //save the max value of each section
    const scores = this.getScores(tabPunteggi); //saves the values of the scores
    const formulas = this.getFormulas(tabPunteggi); //saves the values of the formulas
    const indiciRange = this.getIndicesRange(tabPunteggi, tabCandidati);
    let columnsMax = [];

    this.checkMaxScore(max); //check if the max score possible is correct based on the sum of max score of each section
    this.spreadsheet.createFileValutazione(); //create the new speadsheet
    const formuleManager = new FormuleManager(this.spreadsheet, tabCandidati);

    

    this.spreadsheet.setValori(1, 1, 1, headers.length, [headers]);
    this.spreadsheet.setValori(2, 1, tabCandidati.length, tabCandidati[0].length, tabCandidati); //insert name of the candidates and the subject they apply for in the Valutazione Candidati

    columnsMax = formuleManager.insertCellsTotal(headers, max);
    formuleManager.insertTotalScore(columnsMax);
    formuleManager.insertValuation(minimum);
    formuleManager.insertPointCells(headers, formulas, scores, indiciRange, tabPunteggi, this.namePunteggiSheet);

    this.spreadsheet.setBackground(headers, tabCandidati);
    
  }
}

function provaGenera()
{
    let a = new GenerateValutazioneCandidati('Punteggi', 'ElencoCandidati');
    a.generateValutazione();
}