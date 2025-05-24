class Verbale{
  verbaleId: string;
  body: GoogleAppsScript.Document.Body;
  sheetVC: any;

  constructor(templateId: string, timestamp: string, sheetVC: any) {
    this.verbaleId = DriveApp.getFileById(templateId).makeCopy("Verbale " + timestamp).getId();
    this.body = DocumentApp.openById(this.verbaleId).getBody();
    this.sheetVC = sheetVC;
  }

  getLink(): string {
    return DocumentApp.openById(this.verbaleId).getUrl();
  }

  getName(): string {
    return DocumentApp.openById(this.verbaleId).getName();
  }

  replaceAll(): void {
    try {
      this.replaceDatiVerbale();
      this.replaceCommissione();
      this.replaceElencoPunteggi();
      this.replaceElencoInsegnamenti();
    } catch (e: unknown) {
      DriveApp.getFileById(this.verbaleId).setTrashed(true);
      throw e;
    }
  }

  replaceDatiVerbale(): void {
    try {
      const sheet: SheetDatiVerbale = new SheetDatiVerbale();
      const datiVerbale: any = sheet.getDatiVerbale();

      this.replacePlaceholder(datiVerbale);
    } catch (e: unknown) {
      throw e;
    }
  }

  replaceCommissione(): void {
    try {
      const sheet: SheetCommissione = new SheetCommissione();

      const paragraphsId: number[] = [];
      const paragraphs: GoogleAppsScript.Document.Paragraph[] = [];
      const allPar: GoogleAppsScript.Document.Paragraph[] = this.body.getParagraphs();
      
      for (let i in allPar) {
        if (allPar[i].getText().toString().match("«NomeCommissario»") && 
            allPar[i].getParent().getType() === DocumentApp.ElementType.BODY_SECTION) {
          paragraphsId.push(this.body.getChildIndex(allPar[i]));
          paragraphs.push(allPar[i].copy());
        }
      }

      const allTabs: GoogleAppsScript.Document.Table[] = this.body.getTables();
      const tables: GoogleAppsScript.Document.Table[] = [];
      const tRow: GoogleAppsScript.Document.TableRow[] = [];

      for (let j in allTabs) {
        if (allTabs[j].findText('«NomeCommissario»')) {
          tables.push(allTabs[j]);
          tRow.push(allTabs[j].findText('«NomeCommissario»').getElement().getParent().getParent().getParent().asTableRow().copy());
        }
      }

      const commissione: any[] = sheet.getCommissione();

      for (let i = 0; i < commissione.length; i++) {
        this.replacePlaceholder(commissione[i]);

        if (i < commissione.length - 1) {
          for (let j in paragraphsId) {
            paragraphsId[j]++;
            this.body.insertParagraph(paragraphsId[j], paragraphs[j].copy());
          }

          for (let j in tables) {
            tables[j].appendTableRow(tRow[j].copy());
          }
        }
      }

      const ruoli: Record<string, string> = sheet.getRuoli();

      this.replacePlaceholder(ruoli);
      
    } catch (e) { throw e; }
  }

  replaceElencoPunteggi(): void {
    try {
      const sheet: SheetCriteri = new SheetCriteri();

      this.replacePlaceholder(sheet.getPuntMaxMin());

      let titoloCriterioId: number = 0;
      let titoloCriterioParagraph: GoogleAppsScript.Document.Paragraph | null = null;
      if (this.body.findText("«TitoloCriterio»")) {
        titoloCriterioId = this.body.getChildIndex(this.body.findText("«TitoloCriterio»").getElement().getParent());
        titoloCriterioParagraph = this.body.getChild(titoloCriterioId).asParagraph().copy();
      }

      let critId: number = 1;
      let critParagraph: GoogleAppsScript.Document.ListItem | null = null;
      let macroCritParagraph: GoogleAppsScript.Document.Paragraph | null = null;
      let subCritParagraph: GoogleAppsScript.Document.Paragraph | null = null;
      
      if (this.body.findText("«Criterio»")) {
        critParagraph = this.body.findText("«Criterio»").getElement().getParent().asListItem().copy();
        critId = this.body.getChildIndex(this.body.findText("«Criterio»").getElement().getParent());
        this.body.removeChild(this.body.findText("«Criterio»").getElement().getParent());
        
        if (this.body.findText("«MacroCriterio»")) {
          macroCritParagraph = this.body.findText("«MacroCriterio»").getElement().getParent().asParagraph().copy();
          this.body.removeChild(this.body.findText("«MacroCriterio»").getElement().getParent());
        }

        if (this.body.findText("«SubCriterio»")) {
          subCritParagraph = this.body.findText("«SubCriterio»").getElement().getParent().asParagraph().copy();
          this.body.removeChild(this.body.findText("«SubCriterio»").getElement().getParent());
        }
      }

      const criteri: Array<Record<string, string>> = sheet.getAllCriteri();

      for (let i: number = 0; i < criteri.length; i++) {

        if (i === 0) {

            if (macroCritParagraph !== null) {
                this.body.insertParagraph(critId - 1, macroCritParagraph.copy());
            }
            this.replacePlaceholder(criteri[i]);

        } else if (criteri[i].Tipo === "Macro") {

          if (macroCritParagraph !== null) {
            this.body.insertParagraph(critId - 1, macroCritParagraph.copy());
            critId++;
          }
          if (titoloCriterioParagraph !== null) {
            titoloCriterioId++;
            this.body.insertParagraph(titoloCriterioId, titoloCriterioParagraph.copy());
          }
          this.replacePlaceholder(criteri[i]);

        } else if (criteri[i].Tipo === "Vuoto") {

          if (critId > 1) {
            this.body.insertParagraph(critId, "\n");
            critId += 2;
          }

        } else {

          if (criteri[i].Tipo === "Sub") {

            if (subCritParagraph !== null) {
              this.body.insertParagraph(critId, subCritParagraph.copy());
              this.replacePlaceholder(criteri[i]);
              critId++;
            }

          } else {

            if (critParagraph !== null) {
              if (criteri[i].PuntiCriterio === "") {
                this.body.insertListItem(critId, critParagraph.copy().replaceText('punti', '').asListItem());
              } else {
                this.body.insertListItem(critId, critParagraph.copy());
              }
              critId++;
              this.replacePlaceholder(criteri[i]);
            }

          }

        }

      }
    } catch(e: any) { throw e; }
  }


  replaceElencoInsegnamenti(): void {
    try {
      const sheet: SheetInsegnamenti = new SheetInsegnamenti();
      const insegnamenti: Array<any> = sheet.getInsegnamenti();

      const insParagraph: Array<any> = [];
      let insegnamentoId: number = 0;

      if (this.body.findText("«INIZIO»") !== null && this.body.findText("«FINE»") !== null) {
        const idInizio: number = this.body.getChildIndex(this.body.findText("«INIZIO»").getElement().getParent());
        const idFine: number = this.body.getChildIndex(this.body.findText("«FINE»").getElement().getParent());
        insegnamentoId = idInizio + 1;

        while (insegnamentoId !== idFine) {
          insParagraph.push(this.body.getChild(insegnamentoId).copy());
          insegnamentoId++;
        }

        this.body.removeChild(this.body.getChild(idFine));
        this.body.removeChild(this.body.getChild(idInizio));
        insegnamentoId--;
      }

      for (let i: number = 0; i < insegnamenti.length; i++) {
        this.replacePlaceholder(insegnamenti[i]);

        insegnamentoId = this.replaceValutazioneCandidati(insegnamenti[i].DescrizioneCopertura, insegnamentoId);

        if (parseInt(insParagraph.length.toString()) !== 0 && i < insegnamenti.length - 1) {
          for (const t in insParagraph) {
              if (insParagraph[t].getType() === DocumentApp.ElementType.PARAGRAPH) {
                this.body.insertParagraph(insegnamentoId, insParagraph[t].asParagraph().copy());
              } else if (insParagraph[t].getType() === DocumentApp.ElementType.LIST_ITEM) {
                this.body.insertListItem(insegnamentoId, insParagraph[t].asListItem().copy());
              } else if (insParagraph[t].getType() === DocumentApp.ElementType.TABLE) {
                this.body.insertTable(insegnamentoId, insParagraph[t].asTable().copy());
              }
            insegnamentoId++;
          }
        }
      }
    } catch (e) { throw e; }
  }

  replaceValutazioneCandidati(insegnamento: string, insegnamentoId: number): number {
    try {
      const sheet: SheetValutazioneCandidati = new SheetValutazioneCandidati(this.sheetVC);
      const candiati: Array<Record<string, string>> = sheet.getCandidati(insegnamento);

      const sheetPunti: SheetCriteri = new SheetCriteri();
      const datiPunteggi: { TitoliCriteri: string[], LettereCriteri: string[], Pmin: number } = sheetPunti.getValoriPerTabella();

      const sufficienti: Array<Record<string, string>> = sheet.getSufficienti(candiati, datiPunteggi.Pmin);

      const candParagraph: Array<GoogleAppsScript.Document.Element> = [];
      const candId: Array<number> = [];
      
      while (this.body.findText('«NomeCandidato»')) {
        candId.push(this.body.getChildIndex(this.body.findText("«NomeCandidato»").getElement().getParent()) + candId.length);
        candParagraph.push(this.body.findText("«NomeCandidato»").getElement().getParent().copy());
        this.body.removeChild(this.body.findText("«NomeCandidato»").getElement().getParent());
      }

      let tabCand: GoogleAppsScript.Document.Table | null = null;
      let tabRow: GoogleAppsScript.Document.TableRow | null = null;
      let f: number = 1;
      if (this.body.findText("«Lettera»")) {
        tabCand = this.body.findText("«Lettera»").getElement().getParent().getParent().asTableCell().getParentTable().copy();
        f= tabCand.getChildIndex(tabCand.findText("«Lettera»").getElement().getParent().getParent().getParent());
        tabRow = tabCand.getChild(f).asTableRow().copy();

        this.body.removeChild(this.body.findText("«Lettera»").getElement().getParent().getParent().asTableCell().getParentTable());
      }

      for (let i: number = 0; i < candiati.length; i++) {
        if (parseInt(candParagraph.length.toString()) !== 0) {
          for (const t in candParagraph) {
            if (candParagraph[t].getType() === DocumentApp.ElementType.LIST_ITEM) {
              this.body.insertListItem(candId[t], candParagraph[t].asListItem().copy());
            } else {
              this.body.insertParagraph(candId[t], candParagraph[t].asParagraph().copy());
            }
            candId[t]++;
            candId[t] += parseInt(t);
            insegnamentoId++;
          }
        }

        const j: number = candParagraph.length - 1;
        let tab: GoogleAppsScript.Document.Table | null = null;
        if (tabCand !== null) {
          const tabella: GoogleAppsScript.Document.Table = this.body.insertTable(candId[j], tabCand.copy());
          candId[j] += 2;
          this.body.insertParagraph(candId[j], "");
          candId[j]++;
          insegnamentoId++;
          tab = tabella;
        }

        this.replacePlaceholder(candiati[i]);

        //inserisco dati tabella
        if (tabCand !== null) {
          const punteggi: Array<number> = sheet.getPunteggiCandidato(candiati[i].NomeCandidato);
          let r: number = f;
          let c = 0;
          for (let t in datiPunteggi['TitoliCriteri']) {
            const riga: Record<string, string> = {
              Lettera: datiPunteggi['LettereCriteri'][t],
              Titolo: datiPunteggi['TitoliCriteri'][t],
              Punti: String(punteggi[t])
            };
            
            this.replacePlaceholder(riga);

            
            if (c < datiPunteggi['TitoliCriteri'].length - 1) {
              r++;
              if(tab !== null && tabRow !== null){
                tab.insertTableRow(r, tabRow.copy());
              }
            }
            c++;
          }
        }
          
      }

      //sufficienti
      if (this.body.findText("«CandidatoSuff»") !== null) {
        const suffParagraph: GoogleAppsScript.Document.Element = this.body.findText("«CandidatoSuff»").getElement().getParent().copy();
        let suffId: number = this.body.getChildIndex(this.body.findText("«CandidatoSuff»").getElement().getParent());
        
        this.body.removeChild(this.body.findText("«CandidatoSuff»").getElement().getParent());
        insegnamentoId--;

        for (const j in sufficienti) {
          this.body.insertParagraph(suffId, suffParagraph.asParagraph().copy())
          this.replacePlaceholder(sufficienti[j]);
          suffId++;
          insegnamentoId++;
        }
      }

      return insegnamentoId;
    } catch(e) { throw e; }
  }

  //rimpiazza i tag nel documento
  replacePlaceholder(replacements: Record<string, string>): void {
    for (const key in replacements) {
      this.body.replaceText(`«${key}»`, replacements[key]);
    }
  }
}