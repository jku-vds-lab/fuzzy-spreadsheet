import CellOperations from "../cell/celloperations";
import CellProperties from "../cell/cellproperties";
import UIOptions from "../ui/uioptions";

// changed sheet name
/* global console, setTimeout, Excel */
export default class SheetProp {

  protected static isInputRelationship: boolean = false;
  protected static isOutputRelationship: boolean = false;
  protected static isImpact: boolean = false;
  protected static isLikelihood: boolean = false;
  protected static isSpread: boolean = false;
  protected static isReferenceCell: boolean = false;
  protected static degreeOfNeighbourhood: number = 0;
  protected static newCells: CellProperties[] = null;

  protected cellOp: CellOperations;
  protected cellProp = new CellProperties();
  protected cells: CellProperties[];
  protected referenceCell: CellProperties;
  protected uiOptions: UIOptions;

  private originalTopBorder: Excel.RangeBorder;
  private originalBottomBorder: Excel.RangeBorder;
  private originalLeftBorder: Excel.RangeBorder;
  private originalRightBorder: Excel.RangeBorder;


  constructor() {
    this.uiOptions = new UIOptions();
    this.cellProp = new CellProperties();
    this.cells = new Array<CellProperties>();
    this.cellOp = new CellOperations(null, null, null);
    this.referenceCell = null;
  }

  public getCells() {
    return this.cells;
  }

  public getReferenceCell() {
    return this.referenceCell;
  }


  public resetApp() {
    SheetProp.isInputRelationship = false;
    SheetProp.isOutputRelationship = false;
    SheetProp.isImpact = false;
    SheetProp.isLikelihood = false;
    SheetProp.isSpread = false;
    SheetProp.isReferenceCell = false;
    SheetProp.degreeOfNeighbourhood = 0;
    SheetProp.newCells = null;
    this.uiOptions.deSelectAllOoptions();
  }

  public async processNewValues() {

    this.cells = await this.cellProp.getCells();
    this.cellProp.getRelationshipOfCells(this.cells);
    this.addPropertiesToCells(this.referenceCell.address);

    this.cellOp.removeShapesInfluenceWise('Input');
    this.cellOp.removeShapesInfluenceWise('Output');

    this.cellOp.removeShapesReferenceCellWise();

    setTimeout(() => this.displayOptions(), 1000);
    this.protectSheet();
  }

  public async keepOldValues() {
    this.cellOp.removeShapesInfluenceWise('Input');
    this.cellOp.removeShapesInfluenceWise('Output');

    this.cellOp.removeShapesReferenceCellWise();

    setTimeout(() => this.displayOptions(), 1000);
    // this.protectSheet();
  }


  public async parseSheet() {
    try {

      console.log('Parsing the sheet');

      this.uiOptions.hideOptions();
      // eslint-disable-next-line require-atomic-updates
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

      this.unprotectSheet();

      if (SheetProp.isReferenceCell) {
        this.cellOp.setOptions(SheetProp.isImpact, SheetProp.isLikelihood, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);
        this.cellOp.removeAllShapes();
        this.setBorderToOriginal();
      }

      let range: Excel.Range;

      Excel.run(async (context) => {

        range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
        this.setNewBorder(range.address);
        const address = range.address;
        let index = address.indexOf('!');
        let localAddress = address.substr(index + 1);

        this.uiOptions.addRefCellAddressInTaskpane(localAddress);

        console.log('Marking a reference cell');

        this.addPropertiesToCells(range.address);

        console.log('Done Marking a reference cell');

        this.uiOptions.showVisualizationOption();
        this.registerCellSelectionChangedEvent();
        this.displayOptions();
        this.protectSheet();

      });
    } catch (error) {
      console.error(error);
    }
  }

  public addPropertiesToCells(address: string) {

    this.referenceCell = this.cellProp.getReferenceAndNeighbouringCells(address);
    this.cellProp.checkUncertainty(this.cells);
    this.cellProp.addVarianceAndLikelihoodInfo(this.cells);
    this.cellOp = new CellOperations(this.cells, this.referenceCell, 0);
    SheetProp.isReferenceCell = true;
    this.cellOp.setCells(this.cells);
  }

  private setBorderToOriginal() {
    try {
      Excel.run(async context => {

        console.log('Setting back the original border of: ' + this.referenceCell.address);

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

  private setNewBorder(address: string) {
    try {
      Excel.run(async context => {

        let color: string = 'orange';

        let range = context.workbook.worksheets.getActiveWorksheet().getRange(address);

        // this.getOriginalBorder();

        this.originalTopBorder = range.format.borders.getItem('EdgeTop');
        this.originalBottomBorder = range.format.borders.getItem('EdgeBottom');
        this.originalLeftBorder = range.format.borders.getItem('EdgeLeft');
        this.originalRightBorder = range.format.borders.getItem('EdgeRight');

        this.originalTopBorder.load(['color', 'weight', 'style']);
        this.originalBottomBorder.load(['color', 'weight', 'style']);
        this.originalLeftBorder.load(['color', 'weight', 'style']);
        this.originalRightBorder.load(['color', 'weight', 'style']);

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

  public displayOptions() {

    try {

      this.cellOp.setOptions(SheetProp.isImpact, SheetProp.isLikelihood, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);

      if (SheetProp.degreeOfNeighbourhood == 0) {
        if (SheetProp.isSpread) {
          this.spread();
        }
        return;
      }

      this.handleImpactLikelihood();

      if (SheetProp.isInputRelationship) {
        this.showInputRelationship();
      }

      if (SheetProp.isOutputRelationship) {
        this.showOutputRelationship();
      }

      if (SheetProp.isSpread) {
        this.spread();
      }

    } catch (error) {
      console.log(error);
    }
  }

  public protectSheet() {
    Excel.run(function (context) {
      var activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("protection/protected");

      return context.sync().then(function () {
        if (!activeSheet.protection.protected) {
          activeSheet.protection.protect();
          console.log('Sheet is protected');
        }
      })
    }).catch((reason) => console.log(reason));
  }

  public unprotectSheet() {
    Excel.run(async (context) => {
      let workbook = context.workbook;
      // workbook.protection.unprotect();
      workbook.worksheets.getActiveWorksheet().protection.unprotect();
      return context.sync().then(() => (console.log('Sheet is unprotected'))).catch((reason) => console.log(reason));
    });
  }

  public inputRelationship() {

    try {

      SheetProp.isInputRelationship = this.uiOptions.isElementChecked('inputRelationship');
      this.cellOp.setOptions(SheetProp.isImpact, SheetProp.isLikelihood, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);

      if (SheetProp.isInputRelationship) {

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
      SheetProp.isOutputRelationship = this.uiOptions.isElementChecked('outputRelationship');
      this.cellOp.setOptions(SheetProp.isImpact, SheetProp.isLikelihood, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);

      if (SheetProp.isOutputRelationship) {

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
    SheetProp.degreeOfNeighbourhood = n;
    console.log('DON: ' + n);
    this.cellOp.setOptions(SheetProp.isImpact, SheetProp.isLikelihood, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);
    this.cellOp.removeShapesNeighbourWise(n);
    setTimeout(() => this.displayOptions(), 1000);
  }

  public impact() {

    try {
      SheetProp.isImpact = this.uiOptions.isElementChecked('impact');
      this.cellOp.setOptions(SheetProp.isImpact, SheetProp.isLikelihood, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);
      this.displayOptions();
    } catch (error) {
      console.log(error);
    }
  }

  public likelihood() {

    try {
      SheetProp.isLikelihood = this.uiOptions.isElementChecked('likelihood');
      this.cellOp.setOptions(SheetProp.isImpact, SheetProp.isLikelihood, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);
      this.displayOptions();
    } catch (error) {
      console.log(error);
    }
  }

  public handleImpactLikelihood() {

    try {

      if (SheetProp.isImpact && SheetProp.isLikelihood) {

        this.cellOp.setOptions(false, false, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);
        this.cellOp.removeShapesOptionWise('Likelihood');
        this.cellOp.removeShapesOptionWise('Impact');

        if (SheetProp.isInputRelationship) {
          this.cellOp.showInputImpact(SheetProp.degreeOfNeighbourhood, true);
          this.cellOp.showInputLikelihood(SheetProp.degreeOfNeighbourhood, false);
        }

        if (SheetProp.isOutputRelationship) {
          this.cellOp.showOutputImpact(SheetProp.degreeOfNeighbourhood, true);
          this.cellOp.showOutputLikelihood(SheetProp.degreeOfNeighbourhood, false);
        }

        this.cellOp.setOptions(true, true, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);

      } else if (SheetProp.isImpact) {

        this.cellOp.setOptions(false, SheetProp.isLikelihood, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);
        this.cellOp.removeShapesOptionWise('Likelihood');
        this.cellOp.removeShapesOptionWise('Impact');

        if (SheetProp.isInputRelationship) {
          this.cellOp.showInputImpact(SheetProp.degreeOfNeighbourhood, true);
        }

        if (SheetProp.isOutputRelationship) {
          this.cellOp.showOutputImpact(SheetProp.degreeOfNeighbourhood, true);
        }

        this.cellOp.setOptions(true, SheetProp.isLikelihood, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);

      } else if (SheetProp.isLikelihood) {

        this.cellOp.setOptions(SheetProp.isImpact, false, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);
        this.cellOp.removeShapesOptionWise('Likelihood');
        this.cellOp.removeShapesOptionWise('Impact');

        if (SheetProp.isInputRelationship) {
          this.cellOp.showInputLikelihood(SheetProp.degreeOfNeighbourhood, true);
        }

        if (SheetProp.isOutputRelationship) {
          this.cellOp.showOutputLikelihood(SheetProp.degreeOfNeighbourhood, true);
        }

        this.cellOp.setOptions(SheetProp.isImpact, true, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);
      } else {

        this.cellOp.setOptions(SheetProp.isImpact, SheetProp.isLikelihood, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);
        this.cellOp.removeShapesOptionWise('Impact');
        this.cellOp.removeShapesOptionWise('Likelihood');
      }
    } catch (error) {
      console.log(error);
    }
  }

  public hideRelationshipMarkers() {


  }

  public showInputRelationship() {
    if (SheetProp.isInputRelationship) {
      this.cellOp.showInputRelationship(SheetProp.degreeOfNeighbourhood);
    }
  }

  public showOutputRelationship() {
    if (SheetProp.isOutputRelationship) {
      this.cellOp.showOutputRelationship(SheetProp.degreeOfNeighbourhood);
    }
  }



  public spread() {

    try {
      SheetProp.isSpread = this.uiOptions.isElementChecked('spread');

      this.cellOp.setOptions(SheetProp.isImpact, SheetProp.isLikelihood, SheetProp.isSpread, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);

      if (SheetProp.isSpread) {
        this.cellOp.showSpread(SheetProp.degreeOfNeighbourhood, SheetProp.isInputRelationship, SheetProp.isOutputRelationship);

      } else {
        this.cellOp.removeShapesOptionWise('Spread');

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

      this.uiOptions.addSelCellAddressInTaskpane(event.address);

      this.cells.forEach((cell: CellProperties, index: number) => {

        if (cell.address.includes(event.address)) {

          this.uiOptions.removeImpactInfoInTaskpane();
          this.uiOptions.removeRelationshipInfoInTaskpane();

          if (cell.isImpact) {
            if (SheetProp.newCells != null) {
              this.uiOptions.drawImpactLegend(cell.impact, SheetProp.newCells[index].impact, cell.rectColor);
            } else {
              this.uiOptions.addImpactPercentage(cell);
              this.uiOptions.drawImpactLegend(cell.impact, -1, cell.rectColor);
            }

          }

          if (cell.isLikelihood) {
            this.uiOptions.addLikelihoodPercentage(cell);
            this.uiOptions.drawLikelihoodLegend(cell.likelihood, -1);
          }

          if (cell.isInputRelationship) {
            this.uiOptions.highlightInputRelationshipInfo(cell, SheetProp.degreeOfNeighbourhood);
          }

          if (cell.isOutputRelationship) {
            this.uiOptions.highlightOutputRelationshipInfo(cell, SheetProp.degreeOfNeighbourhood);
          }

          if (cell.isSpread) {
            this.uiOptions.removeHtmlSpreadInfoForOriginalChart();
            this.uiOptions.showSpreadInTaskPane(cell);
          } else {
            this.uiOptions.removeHtmlSpreadInfoForOriginalChart();
            // this.uiOptions.removeHtmlSpreadInfoForNewChart();

          }
          return;
        }
      })
    } catch (error) {
      console.log(error);
    }
  }
}
