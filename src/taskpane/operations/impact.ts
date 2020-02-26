/* global console, Excel */

// TO-DO's

// fix the red issue
// work on output
// remove duplicates

import { std, abs } from 'mathjs';
import CellProperties from '../cellproperties';
import CommonOperations from './commonops';
import SheetProperties from '../sheetproperties';
import { increment } from 'src/functions/functions';

export default class Impact {

  private referenceCell: CellProperties;
  private commonOps: CommonOperations;

  constructor(referenceCell: CellProperties) {
    this.referenceCell = referenceCell;
    this.commonOps = new CommonOperations();
  }

  public showImpact(n: number = 1) {

    this.addImpactInfo(n);
    const commonOps = new CommonOperations();

    if (SheetProperties.isLikelihood) {
      commonOps.deleteRectangles();
    }

    this.showInputImpact(this.referenceCell, n);
    this.showOutputImpact(this.referenceCell, n);
  }

  public async removeImpact() {

    this.removeImpactInfo();

    await this.commonOps.deleteRectangles();

    if (SheetProperties.isLikelihood) {

      this.commonOps.drawRectangles(this.referenceCell.inputCells);
      this.commonOps.drawRectangles(this.referenceCell.outputCells);
    }
  }

  private showInputImpact(cell: CellProperties, i: number) {

    this.commonOps.drawRectangles(cell.inputCells);

    if (i == 1) {
      return;
    }

    cell.inputCells.forEach((inCell: CellProperties) => {
      this.showInputImpact(inCell, i - 1);
    })
  }

  private showOutputImpact(cell: CellProperties, i: number) {
    this.commonOps.drawRectangles(cell.outputCells);

    if (i == 1) {
      return;
    }

    cell.outputCells.forEach((outCell: CellProperties) => {
      this.showOutputImpact(outCell, i - 1);
    })
  }

  private removeImpactInfo() {

    let color = null;
    let transparency = 0;

    if (SheetProperties.isLikelihood) {
      color = 'gray';
    }

    this.referenceCell.inputCells.forEach((inCell: CellProperties) => {
      inCell.rectColor = color;
      inCell.rectTransparency = transparency;
    })

    this.referenceCell.outputCells.forEach((outCell: CellProperties) => {
      outCell.rectColor = color;
      outCell.rectTransparency = transparency;
    })
  }

  private addImpactInfo(n: number = 1) {

    this.addImpactInfoInputCells(n);
    this.addImpactInfoOutputCells();
  }

  private addImpactInfoInputCells(n: number = 1) {

    const isreferenceCellAverage = this.referenceCell.formula.includes("AVERAGE") || this.referenceCell.formula.includes("MITTELWERT");
    const isreferenceCellSum = this.referenceCell.formula.includes("SUM");
    const isreferenceCellDiff = this.referenceCell.formula.includes('-');
    const divisor = this.getDivisor();
    let subtrahend = null;


    if (isreferenceCellDiff) {
      let formula: string = this.referenceCell.formula;
      let idx = formula.indexOf('-');
      subtrahend = formula.substring(idx + 1, formula.length);
    }
    let color = 'green';
    const inputCells = this.referenceCell.inputCells;

    this.referenceCell.inputCells.forEach((inCell: CellProperties) => {
      // fix color issue here
      if (isreferenceCellDiff) {
        color = 'green';
        if (inCell.address.includes(subtrahend)) {
          color = 'red';
        }
      }

      console.log('Incell Address: ' + inCell.address + ' with color: ' + color);
      this.addInputImpactInfoRecursively(inputCells, n, color, divisor);

    })
  }

  private addInputImpactInfoRecursively(inputCells: CellProperties[], n: number, color: string, divisor: number) {

    console.log('color I am here with: ' + color);

    inputCells.forEach((inCell: CellProperties) => {
      inCell.rectColor = color;
      console.log(inCell.rectColor);
      inCell.rectTransparency = abs(1 - (inCell.value / divisor));
    })

    if (n == 1) {
      return;
    }

    n = n - 1;

    inputCells.forEach((inCell: CellProperties) => {
      this.addInputImpactInfoRecursively(inCell.inputCells, n, color, divisor);
    })
  }

  private addImpactInfoOutputCells() {

    this.referenceCell.outputCells.forEach((outCell: CellProperties) => {

      let isSubtrahend: boolean = false;
      let isMinuend: boolean = false;

      if (outCell.formula.includes('-')) {
        //figure out whether the reference cell is minuend or subtrahend to the outcell
        let formula: string = outCell.formula;
        formula = formula.replace('=', '').replace(' ', '');
        let idx = formula.indexOf('-');
        let subtrahend = formula.substring(idx + 1, formula.length);

        if (this.referenceCell.address.includes(subtrahend)) {
          isSubtrahend = true;
        } else {
          isMinuend = true;
        }
      }

      let colorProperties = this.outputColorProperties(outCell.value, this.referenceCell.value, outCell.inputCells, isSubtrahend, isMinuend);
      outCell.rectColor = colorProperties.color;
      outCell.rectTransparency = colorProperties.transparency;
    })
  }

  private computeColor(cellValue: number, referenceCellValue: number, cells: CellProperties[], isSubtrahend: boolean = false, isMinuend: boolean = false) {

    let color = "green";

    if (isSubtrahend) {
      color = "red";
      return color;
    }

    if (referenceCellValue > 0 && cellValue < 0) {
      if (isMinuend) {
        color = "green";
      } else {
        color = "red";
      }
    }

    if (referenceCellValue < 0 && cellValue < 0) { // because of the negative sign, the smaller the number the higher it is
      let isAnyCellPositive = false;

      cells.forEach((cell: CellProperties) => {
        if (cell.value > 0) {
          isAnyCellPositive = true;
        }
      })

      if (isAnyCellPositive) { // if even one cell is positive, then it means that only that cell is contributing positively and rest all are contributing negatively
        color = "red";
      }
    }
    return color;
  }

  // Fix color values for negative values
  private outputColorProperties(cellValue: number, referenceCellValue: number, cells: CellProperties[], isSubtrahend: boolean = false, isMinuend: boolean = false) {

    let transparency = 0;
    const color = this.computeColor(cellValue, referenceCellValue, cells, isSubtrahend, isMinuend);

    // Make both values positive
    if (cellValue < 0) {
      cellValue = -cellValue;
    }

    if (referenceCellValue < 0) {
      referenceCellValue = -referenceCellValue;
    }

    if (cellValue > referenceCellValue) {

      let value = referenceCellValue / cellValue;

      transparency = 1 - value;

    }
    else {
      let maxValue = cellValue;
      // go through the input cells of the output cell
      cells.forEach((c: CellProperties) => {
        let val = c.value;
        if (val < 0) {
          val = -val;
        }
        if (val > maxValue) {
          maxValue = val;
        }
      })

      transparency = 1 - (referenceCellValue / maxValue);
    }

    return { color: color, transparency: transparency };
  }

  private getDivisor() {
    let divisor = 1;

    this.referenceCell.inputCells.forEach((c: CellProperties) => {
      divisor += c.value;
    })
    return divisor;
  }
}