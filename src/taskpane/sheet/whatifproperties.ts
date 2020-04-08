import CellProperties from "../cell/cellproperties";
import Spread from "../operations/spread";
import SheetProp from "./sheetproperties";
import CellOperations from "../cell/celloperations";
import UIOptions from "../ui/uioptions";

// Protect the sheet
/* global setTimeout, console, Excel */
export default class WhatIfProps extends SheetProp {

  private oldCells: CellProperties[];
  private oldReferenceCell: CellProperties;
  private newCells: CellProperties[];
  private newReferenceCell: CellProperties;
  private sheetEventResult = null;
  private cellSelectionEvent = null;
  protected uiOptions: UIOptions;
  private changedRefCell: { oldCell: CellProperties, newCell: CellProperties };
  private changedInputCells = new Array<{ oldCell: CellProperties, newCell: CellProperties }>();
  private changedOutputCells = new Array<{ oldCell: CellProperties, newCell: CellProperties }>();

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

  showUIOptionsForWhatIf() {
    this.uiOptions.showWhatIfOptions();
  }

  startWhatIf() {
    this.unprotectSheet();
    this.registerSheetCalculatedEvent();
  }

  dismissNewValues() {
    this.uiOptions.hideWhatIfOptions();
    this.removeHandler();
    this.cellOp.removeShapesUpdatedWise();
    this.uiOptions.removeHtmlSpreadInfoForNewChart();
    this.newCells = null;
    this.cellProp.writeCellsToSheet(this.oldCells);
  }

  keepNewValues() {
    this.uiOptions.hideWhatIfOptions();
    this.removeHandler();
    this.cellOp.removeShapesUpdatedWise();
    this.uiOptions.removeHtmlSpreadInfoForNewChart();
    this.newCells = null;
  }

  public removeHandler() {
    return Excel.run(this.sheetEventResult.context, (context) => {
      this.sheetEventResult.remove();
      this.cellSelectionEvent.remove();

      return context.sync()
        .then(() => {
          this.sheetEventResult = null;
          this.cellSelectionEvent = null;
          console.log("Event handler successfully removed.");
        });
    }).catch((reason: any) => { console.log(reason) });
  }


  registerSheetCalculatedEvent() {

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
      this.cellOp.removeAllShapes(); // important to delete every information
      this.changedRefCell = null;
      this.changedInputCells = new Array<{ oldCell: CellProperties, newCell: CellProperties }>();
      this.changedOutputCells = new Array<{ oldCell: CellProperties, newCell: CellProperties }>();

      this.registerCellSelectionChangedEvent();

    } catch (error) {
      console.log(error);
    }

  }

  public registerCellSelectionChangedEvent() {

    Excel.run(async (context) => {
      let worksheet = context.workbook.worksheets.getActiveWorksheet();
      this.cellSelectionEvent = worksheet.onSelectionChanged.add((event) => this.handleSelectionChange(event));

      await context.sync();
      console.log("What-if Event handler successfully registered for onSelectionChanged event in the worksheet.");
    }).catch((reason: any) => { console.log(reason) });
  }

  async handleSelectionChange(event) {

    try {

      if (this.newCells == null) {
        return;
      }

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

            if ((this.changedRefCell == null) || (this.changedInputCells == null) || (this.changedOutputCells == null)) {
              return;
            }
            console.log('Still being called');
            if (this.checkIfCellBelongsToChangedCells(newCell)) {
              this.uiOptions.removeHtmlSpreadInfoForNewChart();
              this.uiOptions.addHtmlSpreadInfoForNewChart();
              this.uiOptions.showSpreadInTaskPane(newCell, '.what-if-chart', '#FFA500', true);
            } else {
              this.uiOptions.removeHtmlSpreadInfoForNewChart();
            }
          }
          else {
            this.uiOptions.removeHtmlSpreadInfoForNewChart();
          }
          return;
        }
      })
    } catch (error) {
      console.log(error);
    }
  }

  checkIfCellBelongsToChangedCells(cell: CellProperties) {

    let isIncluded = false;

    let isRef = this.checkIfCellBelongsToReferenceChangedCells(cell);
    let isInput = this.checkIfCellBelongsToInputChangedCells(cell);
    let isOutput = this.checkIfCellBelongsToOutputChangedCells(cell);
    isIncluded = isRef || isInput || isOutput;
    return isIncluded;
  }

  checkIfCellBelongsToReferenceChangedCells(cell: CellProperties) {
    let isIncluded = false;

    if (cell.address == this.changedRefCell.newCell.address) {
      isIncluded = true;
    }
    return isIncluded;
  }

  checkIfCellBelongsToInputChangedCells(cell: CellProperties) {
    let isIncluded = false;
    this.changedInputCells.forEach((inCell) => {
      if (cell.address == inCell.newCell.address) {
        isIncluded = true;
        return;
      }
    });

    return isIncluded;

  }

  checkIfCellBelongsToOutputChangedCells(cell: CellProperties) {
    let isIncluded = false;
    this.changedOutputCells.forEach((outCell) => {
      if (cell.address == outCell.newCell.address) {
        isIncluded = true;
        return;
      }
    })

    return isIncluded;
  }

  setDegreeOfNeighbourhood(n: number) {
    this.cellOp.removeShapesNeighbourWise(n);
    this.displayOptions();
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
      this.cellOp.removeShapesOptionWise('ImpactUpdate');
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
      this.cellOp.removeShapesOptionWise('LikelihoodUpdate');
    }

  }

  spread() {

    try {
      if (SheetProp.isSpread) {
        this.cellOp.showSpread(SheetProp.degreeOfNeighbourhood, SheetProp.isInputRelationship, SheetProp.isOutputRelationship, false);
        setTimeout(() => this.compareSpread(), 1000);
      } else {
        this.cellOp.removeShapesOptionWise('SpreadUpdate');
      }
    } catch (error) {
      console.log('Error in what-if spread: ', error);
    }
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
      this.cellOp.removeShapesOptionWise('RelationshipUpdate');
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

  compareSpread() {

    try {

      this.changedInputCells = new Array<{ oldCell: CellProperties, newCell: CellProperties }>();
      this.changedOutputCells = new Array<{ oldCell: CellProperties, newCell: CellProperties }>();

      let n = SheetProp.degreeOfNeighbourhood;

      this.changedRefCell = this.getChangedCell(this.oldReferenceCell, this.newReferenceCell);

      if (!(this.changedRefCell == null)) {

        this.redrawSpread([this.changedRefCell.oldCell], [this.changedRefCell.newCell], 'ReferenceSpread');
      }

      if (SheetProp.isInputRelationship) {

        this.compareInputSpread(this.oldReferenceCell.inputCells, this.newReferenceCell.inputCells, n);

        if (this.changedInputCells.length > 0) {

          let oldUnchangedCells = new Array<CellProperties>();
          this.changedInputCells.forEach((cell) => oldUnchangedCells.push(cell.oldCell));

          let newChangedCells = new Array<CellProperties>();
          this.changedInputCells.forEach((cell) => newChangedCells.push(cell.newCell));

          this.redrawSpread(oldUnchangedCells, newChangedCells, 'InputSpread');
        }
      }

      if (SheetProp.isOutputRelationship) {

        this.compareOutputSpread(this.oldReferenceCell.outputCells, this.newReferenceCell.outputCells, n);

        if (this.changedOutputCells.length > 0) {

          let oldUnchangedCells = new Array<CellProperties>();
          this.changedOutputCells.forEach((cell) => oldUnchangedCells.push(cell.oldCell));

          let newChangedCells = new Array<CellProperties>();
          this.changedOutputCells.forEach((cell) => newChangedCells.push(cell.newCell));

          this.redrawSpread(oldUnchangedCells, newChangedCells, 'OutputSpread');
        }
      }

    } catch (error) {
      console.log(error);
    }
  }

  redrawSpread(oldCells: CellProperties[], newCells: CellProperties[], name: string) {
    try {

      this.cellOp.removeSpreadCellWise(oldCells, name);

      oldCells.forEach((oldCell: CellProperties) => {
        oldCell.isSpread = true;
      })
      this.cellOp.drawSpread(oldCells, name, 'blue', true);
      newCells.forEach((newCell: CellProperties) => {
        newCell.isSpread = true;
      })
      this.cellOp.drawSpread(newCells, name + 'Update', 'orange', false, true);

    } catch (error) {
      console.log(error);
    }
  }

  compareInputSpread(oldInputCells: CellProperties[], newInputCells: CellProperties[], n: number) {

    newInputCells.forEach((newCell: CellProperties) => {

      oldInputCells.forEach((oldCell: CellProperties) => {

        if (newCell.address == oldCell.address) {
          if (newCell.isSpread) {

            let changedCell = this.getChangedCell(oldCell, newCell);

            if (changedCell == null) {
              return;
            }

            if (!this.checkIfCellBelongsToInputChangedCells(changedCell.newCell)) {
              this.changedInputCells.push(changedCell);
            }

            if (n == 1) {
              return;
            }

            this.compareInputSpread(oldCell.inputCells, newCell.inputCells, n - 1);
          }
        }
      })
      if (n == 1) {
        return;
      }
    })
  }


  compareOutputSpread(oldOutputCells: CellProperties[], newOutputCells: CellProperties[], n: number) {

    newOutputCells.forEach((newCell: CellProperties) => {

      oldOutputCells.forEach((oldCell: CellProperties) => {

        if (newCell.address == oldCell.address) {
          if (newCell.isSpread) {

            let changedCell = this.getChangedCell(oldCell, newCell);

            if (changedCell == null) {
              return;
            }

            if (!this.checkIfCellBelongsToOutputChangedCells(changedCell.newCell)) {
              this.changedOutputCells.push(changedCell);
            }

            if (n == 1) {
              return;
            }

            this.compareOutputSpread(oldCell.outputCells, newCell.outputCells, n - 1);
          }
        }
      })
      if (n == 1) {
        return;
      }
    })
  }

  getChangedCell(oldCell: CellProperties, newCell: CellProperties) {
    let changedCell: { oldCell: CellProperties, newCell: CellProperties };

    if ((oldCell.value == newCell.value) && (oldCell.stdev == newCell.stdev) && (oldCell.likelihood == newCell.likelihood)) {
      changedCell = null;
    } else {
      changedCell = { oldCell: oldCell, newCell: newCell };
    }

    return changedCell;
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
        // textbox.setZOrder(Excel.ShapeZOrder.sendToBack);
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

        let range = sheet.getRange(this.newReferenceCell.address);
        range.select();
        return context.sync().then(() => console.log('Updated shapes')).catch((reason: any) => console.log('Failed to draw the updated shape: ' + reason));
      });
    } catch (error) {
      console.log(error);
    }
  }
}