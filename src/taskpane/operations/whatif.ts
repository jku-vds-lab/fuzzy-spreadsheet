import CellProperties from "../cellproperties";
import Spread from "./spread";
import { increment } from "src/functions/functions";

/* global console, Excel */
export default class WhatIf {
  public value: number = 0;
  public variance: number = 0;
  public likelihood: number = 0;
  public spreadRange: string = null;
  private newCells: CellProperties[];
  private referenceCell: CellProperties;

  setNewCells(newCells: CellProperties[], referenceCell: CellProperties) {
    this.newCells = newCells;
    this.referenceCell = referenceCell;
  }

  async calculateUpdatedNumber() {

    try {
      this.newCells.forEach(async (newCell: CellProperties, index: number) => {

        newCell.whatIf = new WhatIf();
        newCell.whatIf.value = newCell.value - this.referenceCell.value;


        if (this.referenceCell.id == newCell.id) {
          this.referenceCell.whatIf = new WhatIf();
          this.referenceCell.whatIf.value = newCell.value - this.referenceCell.value;
          console.log('Reference Cell value: ' + this.referenceCell.value + ' and new cell value: ' + newCell.value);
          this.referenceCell.whatIf.variance = this.newCells[index + 1].value - this.referenceCell.variance;

        }
      })

    } catch (error) {
      console.log('Error: ', error);
    }
  }

  showUpdateTextInCells(n: number = 1) {

    try {
      this.newCells.forEach((newCell: CellProperties) => {

        if (n == 1) {

          newCell.inputCells.forEach((inCell: CellProperties) => {
            this.addTextBoxOnUpdate(inCell, inCell.whatIf.value);
          })

          newCell.outputCells.forEach((outCell: CellProperties) => {
            this.addTextBoxOnUpdate(outCell, outCell.whatIf.value);
          })

        }

      })
    } catch (error) {
      console.log(error);
    }
  }

  showUpdateTextInInputCells(cells: CellProperties, n: number) {

    // try {

    //   cells.forEach((newCell: CellProperties) => {
    //     if (newCell.whatIf.value) {

    //     }
    //   })

    // } catch (error) {
    //   console.log(error);
    // }

  }

  showUpdateTextInOutputCells(n: number) {
    // try {

    // } catch (error) {
    //   console.log(error);
    // }
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
        textbox.left = cell.left;
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

  // check the variance & likelihood
  async drawChangedSpread(referenceCell: CellProperties, degreeOfNeighbourhood: number) {
    let newReferenceCell = null;

    this.newCells.forEach((cell: CellProperties) => {
      if (referenceCell.id == cell.id) {
        newReferenceCell = cell;
        return;
      }
    });

    console.log('Spread in reference cell: ' + this.referenceCell.isSpread);

    this.referenceCell.inputCells.forEach((inCell: CellProperties) => {
      inCell.isSpread = false;
      console.log('Spread in input cell: ' + inCell.isSpread);
    })

    this.referenceCell.outputCells.forEach((inCell: CellProperties) => {
      inCell.isSpread = false;
      console.log('Spread in output cell: ' + inCell.isSpread);
    })

    const spread: Spread = new Spread(this.newCells, newReferenceCell, 'red');

    spread.showSpread(degreeOfNeighbourhood);

  }
}