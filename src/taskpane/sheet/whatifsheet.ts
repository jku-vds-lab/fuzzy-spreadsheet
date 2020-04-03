import CellProperties from "../cell/cellproperties";
import Spread from "../operations/spread";
import SheetProperties from "./sheetproperties";
import CellOperations from "../cell/celloperations";
import UIOptions from "../ui/uioptions";

/* global console, Excel */
export default class WhatIf {

  private cells: CellProperties[];
  private oldCells: CellProperties[];
  private referenceCell: CellProperties;
  private oldReferenceCell: CellProperties;
  protected isInputRelationship: boolean = false;
  protected isOutputRelationship: boolean = false;
  protected isRelationshipIcons: boolean = false;
  protected isImpact: boolean = false;
  protected isLikelihood: boolean = false;
  protected isSpread: boolean = false;
  protected degreeOfNeighbourhood: number = 1;
  protected cellOp: CellOperations;
  protected uiOptions: UIOptions;
  protected cellProp: CellProperties;
  private sheetEventResult = null;

  constructor(oldCells: CellProperties[], oldReferenceCell: CellProperties) {

    this.uiOptions = new UIOptions();
    this.oldCells = oldCells;
    this.cellProp = new CellProperties();
    this.oldReferenceCell = oldReferenceCell;
    this.cells = new Array<CellProperties>();
    this.referenceCell = new CellProperties();
  }
  public protectSheet() {
    Excel.run((context) => {
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
  registerSheetCalculatedEvent() {
    this.unprotectSheet();
    Excel.run(async (context) => {

      var worksheet = context.workbook.worksheets.getActiveWorksheet();
      this.sheetEventResult = worksheet.onChanged.add(() => this.processWhatIf());

      return context.sync()
        .then(() => {
          console.log("Event handler successfully registered for onChanged event in the worksheet.");
        });
    }).catch((reason: any) => { console.log(reason) });
  }

  async processWhatIf() {
    await this.parseSheet();
  }

  async parseSheet() {
    try {
      this.cells = await this.cellProp.getCells();
      this.cellProp.getRelationshipOfCells(this.cells);
      console.log('Old Reference Cell:', this.oldReferenceCell);
      this.addPropertiesToCells(this.oldReferenceCell.address);
    } catch (error) {
      console.log(error);
    }
  }

  public addPropertiesToCells(address: string) {

    try {
      this.referenceCell = this.cellProp.getReferenceAndNeighbouringCells(address);
      this.cellProp.checkUncertainty(this.cells);
      this.cellProp.addVarianceAndLikelihoodInfo(this.cells);
      this.cellOp = new CellOperations(this.cells, this.referenceCell, 1);
      this.cellOp.setCells(this.cells);
      this.registerCellSelectionChangedEvent();

    } catch (error) {
      console.log(error);
    }

  }

  public registerCellSelectionChangedEvent() {

    Excel.run(async (context) => {
      let worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.onSelectionChanged.add((event) => this.handleSelectionChange(event));

      await context.sync();
      console.log("What-if Event handler successfully registered for onSelectionChanged event in the worksheet.");
    }).catch((reason: any) => { console.log(reason) });
  }

  async handleSelectionChange(event) {

    try {

      this.cells.forEach((cell: CellProperties) => {

        if (cell.address.includes(event.address)) {

          this.uiOptions.removeImpactInfoInTaskpane('newImpactPercentage');


          if (cell.isImpact) {
            this.uiOptions.addImpactPercentage(cell, 'newImpactPercentage');
          }

          if (cell.isLikelihood) {
            this.uiOptions.addLikelihoodPercentage(cell);
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

  spread() {
    this.isSpread = this.uiOptions.isElementChecked('spread');

  }

  impact() {


    this.isImpact = this.uiOptions.isElementChecked('impact');
    if (this.impact) {
      // addinfo
      this.cellOp.showInputImpact(this.degreeOfNeighbourhood, false);
    } else {
      // remove info
    }
  }

  likelihood() {
    this.isLikelihood = this.uiOptions.isElementChecked('likelihood');
  }

  setDegreeOfNeighbourhood(n: number) {
    this.degreeOfNeighbourhood = n;
  }

  relationshipIcons() {
    this.isRelationshipIcons = this.uiOptions.isElementChecked('relationship');
  }

  inputRelationship() {
    this.isInputRelationship = this.uiOptions.isElementChecked('inputRelationship');
  }

  outputRelationship() {
    this.isOutputRelationship = this.uiOptions.isElementChecked('outputRelationship');
  }




  startWhatIfAnalysis() {

  }



  calculateChange() {

    let i = 0;
    try {
      this.cells.forEach((newCell: CellProperties, index: number) => {

        newCell.updatedValue = newCell.value - this.oldCells[index].value;
        if (newCell.updatedValue != 0) {
          //

        }

        if (this.oldReferenceCell.id == newCell.id) {
          this.referenceCell = newCell;
        }
        i++;
      })

    } catch (error) {
      console.log('calculateChange Error at ' + this.cells[i].address, error);
    }
  }

  // public displayOptions() {
  //   if (this.isImpact) {
  //     // calculate new impact
  //   }

  //   if (this.isLikelihood) {
  //     // calculate new likelihood
  //   }

  //   if (this.isRelationshipIcons) {
  //     // do nothing at the moment
  //   }

  //   if (this.isSpread) {
  //     // delete old spread
  //     // add original spread in first half
  //     // compute samples for new spread
  //     // add new spread in second half
  //   }
  // }

  dismissWhatIf(isImpact: boolean, isLikelihood: boolean, isRelationship: boolean, isSpread: boolean) {

    if (isImpact) {
      // remove new impact, only change the newcells.impact to zero

    }

    if (isLikelihood) {
      // remove new likelihood, only change the newcells.impact to zero
    }

    if (isRelationship) {
      // do nothing at the moment
    }

    if (isSpread) {
      // delete old spread & mark is spread as false
      // delete new spread
      // add old spread again
    }
  }


  keepWhatIf(isImpact: boolean, isLikelihood: boolean, isRelationship: boolean, isSpread: boolean) {
    if (isImpact) {
      // do something
      // calculate new impact
    }

    if (isLikelihood) {
      // calculate new likelihood
    }

    if (isRelationship) {
      // do nothing at the moment
    }

    if (isSpread) {
      // delete old spread
      // add original spread in first half
      // compute samples for new spread
      // add new spread in second half
    }
  }


  showUpdateTextInCells(n: number, isInput: boolean, isOutput: boolean) {

    try {

      this.showUpdateTextInReferenceCell();

      if (isInput) {
        this.showUpdateTextInInputCells(this.referenceCell.inputCells, n)
      }

      if (isOutput) {
        this.showUpdateTextInOutputCells(this.referenceCell.outputCells, n)
      }

    } catch (error) {
      console.log(error);
    }
  }

  showNewSpread(degreeOfNeighbourhood: number, isInput: boolean, isOutput: boolean) {
    // try {
    //   const spread: Spread = new Spread(this.newCells, this.oldCells, this.newReferenceCell);
    //   spread.showSpread(degreeOfNeighbourhood, isInput, isOutput);

    // } catch (error) {
    //   console.log(error);
    // }
  }

  deleteNewSpread(degreeOfNeighbourhood: number, isInput: boolean, isOutput: boolean) {
    try {

      let namesToBeDeleted = new Array<string>();

      this.cells.forEach((newCell: CellProperties, index: number) => {
        if (newCell.isSpread) {
          namesToBeDeleted.push(newCell.address);
          newCell.samples = null;
          this.oldCells[index].isSpread = false;

        }
      })

      this.deleteSpreadNameWise(namesToBeDeleted);

      // const spread = new Spread(this.oldCells, null, this.referenceCell);
      // spread.showSpread(degreeOfNeighbourhood, isInput, isOutput);

    } catch (error) {
      console.log(error);
    }
  }

  keepNewSpread(degreeOfNeighbourhood: number, isInput: boolean, isOutput: boolean) {
    try {

      let namesToBeDeleted = new Array<string>();

      this.cells.forEach((newCell: CellProperties) => {
        if (newCell.isSpread) {
          namesToBeDeleted.push(newCell.address);
          newCell.isSpread = false;
        }
      })

      this.deleteSpreadNameWise(namesToBeDeleted);

      // const spread = new Spread(this.newCells, null, this.newReferenceCell);
      // spread.showSpread(degreeOfNeighbourhood, isInput, isOutput);

    } catch (error) {
      console.log(error);
    }
  }



  deleteSpreadNameWise(namesToBeDeleted: string[]) {

    try {
      Excel.run((context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let shapes = sheet.shapes;
        shapes.load("items/name");

        return context.sync().then(() => {
          namesToBeDeleted.forEach((name: string) => {
            shapes.items.forEach((shape) => {
              if (shape.name.includes(name)) {
                shape.delete();
              }
            })
          })
        }).catch((reason: any) => console.log(reason));
      });
    } catch (error) {
      console.log('Async Delete Error:', error);
    }
  }

  showUpdateTextInReferenceCell() {

    try {

      const updatedValue = this.referenceCell.updatedValue;

      if (updatedValue == 0) {
        return;
      }

      this.addTextBoxOnUpdate(this.referenceCell, updatedValue);

    } catch (error) {
      console.log('showUpdateTextInReferenceCell: ' + error);
    }
  }

  showUpdateTextInInputCells(cells: CellProperties[], n: number) {

    try {

      cells.forEach((inCell: CellProperties) => {

        const updatedValue = inCell.updatedValue;

        if (updatedValue == 0) {
          return;
        }

        this.addTextBoxOnUpdate(inCell, updatedValue);

        if (n == 1) {
          return;
        }

        this.showUpdateTextInInputCells(inCell.inputCells, n - 1);
      })
    } catch (error) {
      console.log('showUpdateTextInInputCells: ' + error);
    }
  }

  showUpdateTextInOutputCells(cells: CellProperties[], n: number) {

    try {
      cells.forEach((outCell: CellProperties) => {

        const updatedValue = outCell.updatedValue;

        if (updatedValue == 0) {
          return;
        }

        this.addTextBoxOnUpdate(outCell, updatedValue);

        if (n == 1) {
          return;
        }

        this.showUpdateTextInOutputCells(outCell.outputCells, n - 1);
      })
    } catch (error) {
      console.log('showUpdateTextInOutputCells: ' + error);
    }
  }


  addTextBoxOnUpdate(cell: CellProperties, updatedValue: number) {

    try {

      Excel.run((context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        let text = '';

        let color = 'red';
        if (updatedValue > 0) {
          color = 'green';
          text += '+';
        }

        if (updatedValue == Math.ceil(updatedValue)) {
          text += updatedValue;
        } else {
          text += updatedValue.toFixed(2);
        }

        const textbox = sheet.shapes.addTextBox(text);
        textbox.name = "Update1";
        textbox.left = cell.left + 5;
        textbox.top = cell.top + 2;
        textbox.height = cell.height + 4;
        textbox.width = cell.width - 5;
        textbox.lineFormat.visible = false;
        textbox.fill.transparency = 1;
        textbox.textFrame.verticalAlignment = "Distributed";

        let rotation = 0;

        if (color == 'red') {
          rotation = 180;
        }

        let arrow = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.triangle);
        arrow.name = 'Update2';
        arrow.width = 5;
        arrow.height = cell.height / 3;
        arrow.top = cell.top + cell.height / 2 + 2;
        arrow.left = cell.left + 5;
        arrow.lineFormat.color = color;
        arrow.rotation = rotation;
        arrow.fill.setSolidColor(color);
        return context.sync().then(() => console.log('Updated shapes')).catch((reason: any) => console.log('Failed to draw the updated shape: ' + reason));
      });
    } catch (error) {
      console.log(error);
    }
  }
}