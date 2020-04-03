import CellProperties from "../cell/cellproperties";
import Spread from "../operations/spread";
import { increment } from "src/functions/functions";
import { image, utcFormat } from "d3";

// only new cells contain what if values
/* global console, Excel */
export default class WhatIf {
  public value: number = 0;
  public variance: number = 0;
  public likelihood: number = 0;
  public spreadRange: string = null;
  private newCells: CellProperties[];
  private oldCells: CellProperties[];
  private referenceCell: CellProperties;
  private newReferenceCell: CellProperties;

  constructor(newCells: CellProperties[] = null, oldCells: CellProperties[] = null, referenceCell: CellProperties = null) {
    this.newCells = newCells;
    this.oldCells = oldCells;
    this.referenceCell = referenceCell;
    this.newReferenceCell = referenceCell;
  }

  calculateChange() {

    let i = 0;
    try {
      this.newCells.forEach((newCell: CellProperties, index: number) => {

        newCell.whatIf.value = newCell.value - this.oldCells[index].value;
        if (newCell.whatIf.value > 0) {
          console.log('For cell: ' + newCell.address + ' changed Value: ' + newCell.whatIf.value);
        }

        if (this.referenceCell.id == newCell.id) {
          this.newReferenceCell = newCell;
          console.log('Reference Cell Formula', this.newReferenceCell.formula);
        }
        i++;
      })

    } catch (error) {
      console.log('calculateChange Error at ' + this.newCells[i].address, error);
    }
  }

  getChangedCells() {
    let changedCells = new Array<{ oldCell: CellProperties, newCell: CellProperties }>();

    this.newCells.forEach((newCell: CellProperties, index: number) => {

      if (newCell.value == this.oldCells[index].value) {
        if (newCell.stdev == this.oldCells[index].stdev) {
          if (newCell.likelihood == this.oldCells[index].likelihood) {
            return;
          }
        }
      }

      changedCells.push({ oldCell: this.oldCells[index], newCell: newCell });

    });
  }

  displayOptionsForChangeCells(isImpact: boolean, isLikelihood: boolean, isRelationship: boolean, isSpread: boolean) {
    if (isImpact) {
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
        this.showUpdateTextInInputCells(this.newReferenceCell.inputCells, n)
      }

      if (isOutput) {
        this.showUpdateTextInOutputCells(this.newReferenceCell.outputCells, n)
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

      this.newCells.forEach((newCell: CellProperties, index: number) => {
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

  showUpdateTextInReferenceCell() {

    try {

      const updatedValue = this.newReferenceCell.whatIf.value;

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

        const updatedValue = inCell.whatIf.value;

        if (updatedValue == 0) {
          return;
        }

        this.addTextBoxOnUpdate(inCell, inCell.whatIf.value);

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

        const updatedValue = outCell.whatIf.value;

        if (updatedValue == 0) {
          return;
        }

        this.addTextBoxOnUpdate(outCell, outCell.whatIf.value);

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