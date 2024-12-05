class FormuleManager
{
  private sheetManager: SpreadsheetCreator;
  private numRow: number;
  private tabCandidati: any[][];

  constructor(sheetManager: SpreadsheetCreator, tabCandidati: any[][])
  {
    this.sheetManager = sheetManager
    this.numRow = tabCandidati.length;
    this.tabCandidati = tabCandidati;
  }

  //FORSE NON VA QUI
  //method that calculates the previous cell depending on the formula type
  calculatePreviousCell(i: number, r: number, userFormula: string): string
  {
    let previousColumn = "";
    let previousCell = "";

    if(userFormula === "MULTI")
    {
      previousColumn = String.fromCharCode(65 + i - 2); //convert from index to column letter
    }
    else
    {
      previousColumn = String.fromCharCode(65 + i - 1); 
    }

    previousCell = previousColumn + r;
    return previousCell;
  }

  //method that determines the formula to insert
  calculateFormula(userFormula: string, previousCell: string, points: number, indexsRange: number[], tabPunteggi: any[][], sheetNamePunteggi: string)
  {
    let formula = "";

    if(userFormula === "IF")
    {
      formula = '=IF(' + previousCell + ' = "", 0, ' + points + ')';
    }
    else if(userFormula === "MULTI")
    {
      formula = previousCell + "*" + points;
    }
    else if(userFormula === "RANGE")
    {
      let endRange = 0;
      let i = indexsRange[0];
      let startRange = indexsRange[0]; //I save the value of the row from which the range begins (e.g. 18)

      while(i < tabPunteggi.length && tabPunteggi[i][1].startsWith("-"))
      {
        endRange = i;
        i++;
      }

      startRange = parseInt(startRange.toString()) + 2;
      endRange = parseInt(endRange.toString()) + 2; //I have to take into account the fact that we have removed the header and that the arrays start from 0
      indexsRange.shift(); //I delete the index I just created the formula for

      formula = "INDEX(" + sheetNamePunteggi + "!C$" + startRange + ":C$" + endRange + ", MATCH(" + previousCell + "," + sheetNamePunteggi + "!D$" + startRange + ":D$" + endRange + "," + "1))";
    }

    return formula;
  }

  //method that inserts the appropriate formulas into the point cells
  insertPointCells(headers: string[], formule: string[], scores: number[], indexsRange: number[], tabPunteggi: any[][], sheetNamePunteggi: string): void
  {
    let j = 0;

    //scroll through the entire header. If the cell begins with the word "Punti" then it contains a score so I enter the respective formula
    for(let i = 0; i < headers.length; i++) 
    {
      
      if(headers[i].startsWith("Punti"))
      {
        for (let r = 2; r <= this.numRow + 1; r++)
        { 
          let formula = "";
          let userFormula = formule[j].toUpperCase(); 
          let previousCell = this.calculatePreviousCell(i, r, userFormula)

          formula = this.calculateFormula(userFormula, previousCell, scores[j], indexsRange, tabPunteggi, sheetNamePunteggi)

          if(this.tabCandidati[r-2][0] != "")
          {
            this.sheetManager.setFormula(r, i + 1, formula);
          }
        }

        j++;
      }  
    }
  }

  //method that inserts formulas into total cells
  insertCellsTotal(headers: string[], max: number[]): number[]
  {
    let columnsMax = []; //this store the columns that contain the Tot points of a category
    let indexTotPrevious = -1;  //Index of the previous "Tot" column
    let count = 0; //count to scroll through the array where the maximum values ​​for each category are saved

    for (let i = 0; i < headers.length; i++)
    {
      if (headers[i].startsWith("Tot"))
      {
        let columnsPointsToSum = [];

        columnsMax.push(i+1);

        //We look for the "Punti" columns belonging to the current section
        for (let j = indexTotPrevious + 1; j < i; j++)
        {
          if (headers[j].startsWith("Punti"))
          {
            columnsPointsToSum.push(j + 1);  //Memorize the index of the columns "Punti"
          }
        }

        //Create the formula to insert into the cells of the Tot column
        //let's start by scrolling through all the lines
        for(let r = 2; r <= this.numRow + 1; r++)
        {  
          if (columnsPointsToSum.length > 0)
          {
            let formula = "=MIN(" + max[count] + ";";
            
            //Build the sum formula for the current row
            //we use the map function that calls the function in the parameter for each element of the array
            //sumColumns will be an array containing the A1 notation of the cells to be added. I build the formula by joining the elements of the array (and therefore the cells to be added) with a +
            let sumColumns = columnsPointsToSum.map((colIndex) => 
            {
              //returns the A1 format notation of the cells
              return this.sheetManager.getRange(r, colIndex).getA1Notation();
            });

            formula += sumColumns.join('+') + ')';
            
            
            if(this.tabCandidati[r-2][0] !== "")
            {
              this.sheetManager.setFormula(r, i + 1, formula); //Insert the formula on the cell of the Tot column
            }
          }
          else
          {
            this.sheetManager.getRange(r, i + 1).setValue('');
          }
        }

        count++;
        indexTotPrevious = i; //update the index of the previous "Tot" column
      }
    }

    return columnsMax;
  } 

  //method that enters the total score obtained by a candidate
  insertTotalScore(columnsMax: number[]): void
  {
    //we calculate TotPoints based on the values ​​present in the TotX cells
    for(let r = 2; r <= this.numRow + 1; r++)
    {
      let formula = "=SUM(";

      let cellsTotalSections = columnsMax.map((colIndex) => 
      {
        //returns A1 format notation of the cells
        return this.sheetManager.getRange(r, colIndex).getA1Notation();
      });


      formula += cellsTotalSections.join("+") + ")";

      if(this.tabCandidati[r-2][0] !== "")
      {
        this.sheetManager.getRange(r, 2).setFormula(formula);
      }
    }
  }

  //method that enters the candidate's rating (sufficient or not)
  insertValuation()
  {
    //In the sufficiente column I insert a formula that checks the score in the PunteggioTotale cell and tells me whether the candidate is sufficient or not
    for(let r = 2; r <= this.numRow +1; r++)
    {
      let checkCell = "B"+r;
      let formula = "=IF(" + checkCell + "<35;\"NO!!!\";\"SÍ\")";

      if(this.tabCandidati[r-2][0] !== "")
      {
        this.sheetManager.getRange(r, 3).setFormula(formula);
      }
    }
  }
  
}
