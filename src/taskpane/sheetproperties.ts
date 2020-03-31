import CellOperations from "./celloperations";
import CellProperties from "./cellproperties";
import UIOptions from "./uioptions";
/* global console, document, Excel, Office */
export default class SheetProperties {

  private isInputRelationship: boolean = false;
  private isOutputRelationship: boolean = false;
  private isRelationshipIcons: boolean = false;
  private isImpact: boolean = false;
  private isLikelihood: boolean = false;
  private isSpread: boolean = false;
  private isReferenceCell: boolean = false;
  private degreeOfNeighbourhood: number = 1;
  private isCheatSheetExist: boolean = false;
  private cellOp: CellOperations;
  private cellProp = new CellProperties();
  private cells: CellProperties[];
  private newCells: CellProperties[] = null;
  private referenceCell: CellProperties = null;
  private isSheetParsed = false;
  private newValues: any[][];
  private newFormulas: any[][];
  private originalTopBorder: Excel.RangeBorder;
  private originalBottomBorder: Excel.RangeBorder;
  private originalLeftBorder: Excel.RangeBorder;
  private originalRightBorder: Excel.RangeBorder;
  private uiOptions: UIOptions;

  constructor() {
    this.uiOptions = new UIOptions();
    this.cellProp = new CellProperties();
    this.cells = new Array<CellProperties>();
    this.cellOp = new CellOperations(null, null, null);
  }

  public async parseSheet() {

    this.isSheetParsed = true;

    try {

      console.log('Parsing the sheet');
      this.uiOptions.hideOptions();

      this.cells = await this.cellProp.getCells();
      this.cellProp.getRelationshipOfCells(this.cells);

      this.uiOptions.showReferenceCellOption();

      console.log('Done parsing the sheet');

    } catch (error) {
      console.log(error);
    }
  }

  public markAsReferenceCell() {

    try {

      if (this.isReferenceCell) {
        this.cellOp.removeShapesReferenceCellWise();
        this.setBorderToOriginal();
      }

      let range: Excel.Range;

      Excel.run(async (context) => {

        range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        this.setNewBorder();

        console.log('Marking a reference cell');

        this.referenceCell = this.cellProp.getReferenceAndNeighbouringCells(range.address);
        this.cellProp.checkUncertainty(this.cells);
        this.cellProp.addVarianceAndLikelihoodInfo(this.cells);
        this.cellOp = new CellOperations(this.cells, this.referenceCell, 1);
        this.isReferenceCell = true;

        console.log('Done Marking a reference cell');

        this.uiOptions.showVisualizationOption();
        this.registerCellSelectionChangedEvent();
      });


      this.displayOptions();

    } catch (error) {
      console.error(error);
    }
  }

  private setBorderToOriginal() {
    try {
      Excel.run(async context => {

        let range = context.workbook.worksheets.getActiveWorksheet().getRange(this.referenceCell.address);

        range.format.borders.getItem('EdgeTop').color = this.originalTopBorder.color;
        range.format.borders.getItem('EdgeBottom').color = this.originalBottomBorder.color;
        range.format.borders.getItem("EdgeLeft").color = this.originalLeftBorder.color;
        range.format.borders.getItem('EdgeRight').color = this.originalRightBorder.color;

        range.format.borders.getItem('EdgeTop').weight = this.originalTopBorder.weight;
        range.format.borders.getItem('EdgeBottom').weight = this.originalBottomBorder.weight;
        range.format.borders.getItem("EdgeLeft").weight = this.originalLeftBorder.weight;
        range.format.borders.getItem('EdgeRight').weight = this.originalRightBorder.weight;

        range.format.borders.getItem('EdgeTop').style = this.originalTopBorder.style;
        range.format.borders.getItem('EdgeBottom').style = this.originalBottomBorder.style;
        range.format.borders.getItem("EdgeLeft").style = this.originalLeftBorder.style;
        range.format.borders.getItem('EdgeRight').style = this.originalRightBorder.style;

        return context.sync().then(() => { }).catch((reason: any) => console.log(reason));
      })
    } catch (error) {
      console.log(error);
    }
  }

  private setNewBorder() {
    try {
      Excel.run(async context => {

        let color: string = 'orange';

        let range = context.workbook.getSelectedRange();

        this.getOriginalBorder();

        range.format.borders.getItem('EdgeTop').color = color;
        range.format.borders.getItem('EdgeBottom').color = color;
        range.format.borders.getItem("EdgeLeft").color = color;
        range.format.borders.getItem('EdgeRight').color = color;

        range.format.borders.getItem('EdgeTop').weight = "Thick";
        range.format.borders.getItem('EdgeBottom').weight = "Thick";
        range.format.borders.getItem('EdgeLeft').weight = "Thick";
        range.format.borders.getItem('EdgeRight').weight = "Thick";


        return context.sync().then(() => { }).catch((reason: any) => console.log(reason));
      })
    } catch (error) {
      console.log(error);
    }
  }

  private getOriginalBorder() {
    try {
      Excel.run(async context => {

        let range = context.workbook.getSelectedRange();

        this.originalTopBorder = range.format.borders.getItem('EdgeTop');
        this.originalBottomBorder = range.format.borders.getItem('EdgeBottom');
        this.originalLeftBorder = range.format.borders.getItem('EdgeLeft');
        this.originalRightBorder = range.format.borders.getItem('EdgeRight');

        this.originalTopBorder.load(['color', 'weight', 'style']);
        this.originalBottomBorder.load(['color', 'weight', 'style']);
        this.originalLeftBorder.load(['color', 'weight', 'style']);
        this.originalRightBorder.load(['color', 'weight', 'style']);

        return context.sync().then(() => { }).catch((reason: any) => console.log(reason));
      })
    } catch (error) {
      console.log(error);
    }
  }

  private displayOptions() {

    try {
      // this.cellOp.removeShapesNeighbourWise(this.degreeOfNeighbourhood);

      if (this.isImpact) {
        this.impact();
      }

      if (this.isLikelihood) {
        this.likelihood();
      }

      if (this.isRelationshipIcons) {
        this.relationshipIcons();
      }

      if (this.isSpread) {
        this.spread();
      }
    } catch (error) {
      console.log(error);
    }
  }

  public inputRelationship() {

    try {

      this.isInputRelationship = this.uiOptions.isElementChecked('inputRelationship');

      if (this.isInputRelationship) {
        this.uiOptions.showAllOptions();
        this.displayOptions();
      } else {
        this.cellOp.removeShapesInfluenceWise('Input');
      }
    } catch (error) {
      console.log(error);
    }
  }

  public outputRelationship() {

    try {
      this.isOutputRelationship = this.uiOptions.isElementChecked('outputRelationship');

      if (this.isOutputRelationship) {
        this.uiOptions.showAllOptions();
        this.displayOptions();
      } else {
        this.cellOp.removeShapesInfluenceWise('Output');
      }

    } catch (error) {
      console.log(error);
    }
  }

  public setDegreeOfNeighbourhood(n: number) {
    this.degreeOfNeighbourhood = n;
  }

  public impact() {

    try {
      this.isImpact = this.uiOptions.isElementChecked('impact');

      if (this.impact) {

        if (this.isInputRelationship) {
          this.cellOp.showInputImpact(this.degreeOfNeighbourhood);
        }

        if (this.isOutputRelationship) {
          this.cellOp.showOutputImpact(this.degreeOfNeighbourhood);
        }

      } else {
        this.cellOp.removeShapesOptionWise('Impact');

      }
    } catch (error) {
      console.log(error);
    }
  }

  public likelihood() {

    try {
      this.isLikelihood = this.uiOptions.isElementChecked('likelihood');

      if (this.isLikelihood) {

        if (this.isInputRelationship) {
          this.cellOp.showInputLikelihood(this.degreeOfNeighbourhood);
        }

        if (this.isOutputRelationship) {
          this.cellOp.showOutputLikelihood(this.degreeOfNeighbourhood);
        }

      } else {
        this.cellOp.removeShapesOptionWise('likelihood');

      }
    } catch (error) {
      console.log(error);
    }
  }

  public relationshipIcons() {

    try {
      this.isRelationshipIcons = this.uiOptions.isElementChecked('relationship');

      if (this.isRelationshipIcons) {

        if (this.isInputRelationship) {
          this.cellOp.showInputRelationship(this.degreeOfNeighbourhood);
        }

        if (this.isOutputRelationship) {
          this.cellOp.showOutputRelationship(this.degreeOfNeighbourhood);
        }

      } else {
        this.cellOp.removeShapesOptionWise('Relationship');

      }
    } catch (error) {
      console.log(error);
    }
  }

  public spread() {

    try {
      this.isSpread = this.uiOptions.isElementChecked('spread');

      if (this.isSpread) {
        this.cellOp.showSpread(this.degreeOfNeighbourhood, this.isInputRelationship, this.isOutputRelationship);

      } else {
        this.cellOp.removeShapesOptionWise('Relationship');

      }
    } catch (error) {
      console.log(error);
    }
  }

  public registerCellSelectionChangedEvent() {

    Excel.run(async (context) => {
      let worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.onSelectionChanged.add((event) => this.handleSelectionChange(event));

      await context.sync();
      console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
    }).catch((reason: any) => { console.log(reason) });
  }

  async handleSelectionChange(event) {

    try {

      this.cells.forEach((cell: CellProperties) => {

        if (cell.address.includes(event.address)) {

          this.uiOptions.removeImpactInfoInTaskpane();
          this.uiOptions.removeRelationshipInfoInTaskpane();

          if (cell.isImpact) {
            this.uiOptions.addImpactPercentage(cell);
            this.uiOptions.drawImpactLegend(cell.impact, cell.rectColor);
          }

          if (cell.isLikelihood) {
            this.uiOptions.addLikelihoodPercentage(cell);
          }

          if (cell.isInputRelationship) {
            this.uiOptions.highlightInputRelationshipInfo(cell, this.degreeOfNeighbourhood);
          }

          if (cell.isOutputRelationship) {
            this.uiOptions.highlightOutputRelationshipInfo(cell, this.degreeOfNeighbourhood);
          }

          if (cell.isSpread) {
            this.uiOptions.showSpreadInTaskPane(cell);
            this.uiOptions.showMeanAndStdDevValueInTaskpane(cell);
          } else {
            this.uiOptions.removeHtmlSpreadInfoForOriginalChart();
            this.uiOptions.removeHtmlSpreadInfoForNewChart();

          }
          return;
        }
      })
    } catch (error) {
      console.log(error);
    }
  }

  public startWhatIf() {

  }

  public keepNewValues() {

  }

  public dismissNewValues() {

  }
}
