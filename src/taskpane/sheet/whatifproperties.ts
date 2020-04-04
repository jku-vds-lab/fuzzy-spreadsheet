import CellProperties from "../cell/cellproperties";
import Spread from "../operations/spread";
import SheetProp from "./sheetproperties";
import CellOperations from "../cell/celloperations";
import UIOptions from "../ui/uioptions";
import { tickStep } from "d3";

/* global console, Excel */
export default class WhatIfProps extends SheetProp {

  private oldCells: CellProperties[];
  private oldReferenceCell: CellProperties;
  private newCells: CellProperties[];
  private newReferenceCell: CellProperties;
  private sheetEventResult = null;
  protected uiOptions: UIOptions;

  constructor(oldCells: CellProperties[], oldReferenceCell: CellProperties) {

    super();
    this.uiOptions = new UIOptions();
    this.cellProp = new CellProperties();
    this.cellOp = new CellOperations(null, null, null);
    this.oldCells = oldCells;
    this.oldReferenceCell = oldReferenceCell;
    this.newCells = new Array<CellProperties>();
    this.newReferenceCell = new CellProperties();
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

    try {
      await this.parseSheet();
      this.addPropertiesToCells(this.oldReferenceCell.address);
      this.displayOptions();
    } catch (error) {
      console.log(error);
    }
  }

  async parseSheet() {

    try {
      this.newCells = await this.cellProp.getCells();
      this.cellProp.getRelationshipOfCells(this.newCells);
    } catch (error) {
      console.log(error);
    }
  }

  public addPropertiesToCells(address: string) {

    try {

      this.newReferenceCell = this.cellProp.getReferenceAndNeighbouringCells(address);
      this.cellProp.checkUncertainty(this.newCells);
      this.cellProp.addVarianceAndLikelihoodInfo(this.newCells);
      this.cellOp = new CellOperations(this.newCells, this.newReferenceCell, 1, false);
      this.cellOp.setCells(this.newCells);
      this.registerCellSelectionChangedEvent();

      console.log('Old Ref: ' + this.oldReferenceCell.value + 'New Ref: ' + this.newReferenceCell.value);

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

      this.newCells.forEach((newCell: CellProperties) => {

        if (newCell.address.includes(event.address)) {

          this.uiOptions.removeImpactInfoInTaskpane('newImpactPercentage');


          if (newCell.isImpact) {
            this.uiOptions.addImpactPercentage(newCell, 'newImpactPercentage');
          }

          if (newCell.isLikelihood) {
            this.uiOptions.addLikelihoodPercentage(newCell, 'newLikelihoodPercentage');
          }

          if (newCell.isSpread) {
            this.uiOptions.showSpreadInTaskPane(newCell);
            this.uiOptions.showMeanAndStdDevValueInTaskpane(newCell);
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


  displayOptions() {
    this.impact();
    this.likelihood();
    this.spread();
    this.showTextBoxInCells();
  }

  impact() {

    if (SheetProp.isImpact) {

      if (SheetProp.isInputRelationship) {
        this.cellOp.showInputImpact(SheetProp.degreeOfNeighbourhood, false);
      }

      if (SheetProp.isOutputRelationship) {
        this.cellOp.showOutputImpact(SheetProp.degreeOfNeighbourhood, false);
      }
    } else {
      this.cellOp.removeShapesOptionWise('Impact');
    }
  }

  likelihood() {

    if (SheetProp.isImpact) {

      if (SheetProp.isInputRelationship) {
        this.cellOp.showInputLikelihood(SheetProp.degreeOfNeighbourhood, false);
      }

      if (SheetProp.isOutputRelationship) {
        this.cellOp.showOutputLikelihood(SheetProp.degreeOfNeighbourhood, false);
      }
    } else {
      this.cellOp.removeShapesOptionWise('Likelihood');
    }

  }

  spread() {

  }

  relationshipIcons() {

    if (SheetProp.isRelationshipIcons) {

      if (SheetProp.isInputRelationship) {
        this.cellOp.showInputRelationship(SheetProp.degreeOfNeighbourhood);
      }

      if (SheetProp.isOutputRelationship) {
        this.cellOp.showOutputRelationship(SheetProp.degreeOfNeighbourhood);
      }

    } else {
      this.cellOp.removeShapesOptionWise('Relationship');
    }
  }

  inputRelationship() {

    if (SheetProp.isInputRelationship) {
      this.displayOptions();
    } else {
      this.cellOp.removeShapesInfluenceWise('Input');
    }
  }

  outputRelationship() {

    if (SheetProp.isOutputRelationship) {
      this.displayOptions();
    } else {
      this.cellOp.removeShapesInfluenceWise('Output');
    }
  }



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

      this.newCells.forEach((newCell: CellProperties, index: number) => {
        if (newCell.isSpread) {
          namesToBeDeleted.push(newCell.address);
          newCell.samples = null;
          this.oldCells[index].isSpread = false;

        }
      })

      this.deleteSpreadNameWise(namesToBeDeleted);

      // const spread = new Spread(this.oldCells, null, this.newReferenceCell);
      // spread.showSpread(degreeOfNeighbourhood, isInput, isOutput);

    } catch (error) {
      console.log(error);
    }
  }

  keepNewSpread(degreeOfNeighbourhood: number, isInput: boolean, isOutput: boolean) {
    try {

      let namesToBeDeleted = new Array<string>();

      this.newCells.forEach((newCell: CellProperties) => {
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

  showTextBoxInCells() {

    try {

      this.cellOp.removeShapesUpdatedWise();
      this.calculateUpdatedValue();

      this.showUpdateTextInReferenceCell();

      if (SheetProp.isInputRelationship) {
        this.showUpdateTextInInputCells(this.newReferenceCell.inputCells, SheetProp.degreeOfNeighbourhood)
      }

      if (SheetProp.isOutputRelationship) {
        this.showUpdateTextInOutputCells(this.newReferenceCell.outputCells, SheetProp.degreeOfNeighbourhood)
      }

    } catch (error) {
      console.log(error);
    }
  }

  public calculateUpdatedValue() {

    try {
      this.newCells.forEach((newCell: CellProperties, index: number) => {
        newCell.updatedValue = newCell.value - this.oldCells[index].value;
      })

    } catch (error) {
      console.log('calculateChange Error: ', error);
    }
  }

  showUpdateTextInReferenceCell() {

    try {

      const updatedValue = this.newReferenceCell.updatedValue;

      if (updatedValue == 0) {
        return;
      }

      this.addTextBoxOnUpdate(this.newReferenceCell, updatedValue);

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
        textbox.name = cell.address + 'TextBoxUpdate';
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
        arrow.name = cell.address + 'TextBoxUpdate';
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