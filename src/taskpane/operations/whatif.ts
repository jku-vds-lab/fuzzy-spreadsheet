import CellProperties from "../cellproperties";
import Spread from "./spread";
import { increment } from "src/functions/functions";

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
        // newCell.whatIf.variance = this.newCells[index + 1].value - newCell.variance;

        if (this.referenceCell.id == newCell.id) {
          this.newReferenceCell = newCell;
        }
        i++;
      })

    } catch (error) {
      console.log('calculateChange Error at ' + this.newCells[i].address, error);
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

    try {
      const spread: Spread = new Spread(this.newCells, this.oldCells, this.newReferenceCell, 'orange');

      console.log('Computing new spread. Input: ' + isInput + ' Output: ' + isOutput);
      spread.showSpread(degreeOfNeighbourhood, isInput, isOutput);

    } catch (error) {
      console.log(error);
    }
  }

  deleteNewSpread(degreeOfNeighbourhood: number, isInput: boolean, isOutput: boolean) {
    try {

      let namesToBeDeleted = new Array<string>();

      this.newCells.forEach((newCell: CellProperties, index: number) => {
        if (newCell.isSpread) {
          namesToBeDeleted.push(newCell.address);
          newCell.samples = null;
          console.log('Old Cell: ' + this.oldCells[index].address)
          this.oldCells[index].isSpread = false;// Check this and then check if the spread is set to false, then call the showSpread method
          // namesToBeDeleted.push(this.oldCells[index].address);
        }
      })

      this.deleteSpreadNameWise(namesToBeDeleted);

      const spread = new Spread(this.oldCells, null, this.referenceCell);
      spread.showSpread(degreeOfNeighbourhood, isInput, isOutput);

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

        text = updatedValue.toPrecision(1).toString();

        const textbox = sheet.shapes.addTextBox(text);
        textbox.name = "Update1";
        textbox.left = cell.left + 5;
        textbox.top = cell.top;
        textbox.height = cell.height + 4;
        textbox.width = cell.width / 2;
        textbox.lineFormat.visible = false;
        textbox.fill.transparency = 1;
        textbox.textFrame.verticalAlignment = "Distributed";

        let arrow: Excel.Shape;

        if (color == 'red') {
          arrow = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.downArrow);
        } else {
          arrow = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.upArrow);
        }

        arrow.name = 'Update2';
        arrow.width = 5;
        arrow.height = cell.height;
        arrow.top = cell.top;
        arrow.left = cell.left;
        arrow.lineFormat.color = color;
        arrow.fill.setSolidColor(color);
        return context.sync().then(() => console.log('Updated shapes')).catch((reason: any) => console.log('Failed to draw the updated shape: ' + reason));
      });
    } catch (error) {
      console.log(error);
    }
  }
}