/* global console, Excel */

// TO-DO's

// fix the red issue
// work on output
// remove duplicates

import { abs } from 'mathjs';
import CellProperties from '../cellproperties';
import CommonOperations from './commonops';
import SheetProperties from '../sheetproperties';

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
    this.addImpactInfoOutputCells(this.referenceCell, n);
  }

  private addImpactInfoInputCells(n: number = 1) {

    const isreferenceCellDiff = this.referenceCell.formula.includes('-');
    const divisor = this.getDivisor(this.referenceCell);
    let subtrahend = null;


    if (isreferenceCellDiff) {
      let formula: string = this.referenceCell.formula;
      let idx = formula.indexOf('-');
      subtrahend = formula.substring(idx + 1, formula.length);
    }
    let color = 'green';
    const inputCells = this.referenceCell.inputCells;


    this.addInputImpactInfoRecursively(inputCells, n, color, divisor, subtrahend);
  }

  private addInputImpactInfoRecursively(inputCells: CellProperties[], n: number, color: string, divisor: number, subtrahend: string = null) {

    inputCells.forEach((inCell: CellProperties) => {

      inCell.rectColor = color;

      if (inCell.address.includes(subtrahend)) {
        inCell.rectColor = 'red';
      }
      console.log(inCell.rectColor);
      inCell.rectTransparency = abs(1 - (inCell.value / divisor));
    })

    if (n == 1) {
      return;
    }

    n = n - 1;

    inputCells.forEach((inCell: CellProperties) => {
      this.addInputImpactInfoRecursively(inCell.inputCells, n, inCell.rectColor, divisor);
    })
  }

  private addImpactInfoOutputCells(cell: CellProperties, n: number) {

    cell.outputCells.forEach((outCell: CellProperties) => {

      if (outCell.formula.includes('-')) {
        const divisor = this.getDivisor(outCell);
        let formula: string = outCell.formula;
        formula = formula.replace('=', '').replace(' ', '');
        const idx = formula.indexOf('-');
        const subtrahend = formula.substring(idx + 1, formula.length);

        outCell.rectColor = this.checkIfCellHasSubtrahend(cell, subtrahend);
        outCell.rectTransparency = abs(1 - this.referenceCell.value / divisor);

      } else if (outCell.formula.includes('AVERAGE')) {
        const divisor = this.getDivisor(outCell);
        const transparency = abs(1 - this.referenceCell.value / divisor);

        outCell.rectTransparency = transparency;
        outCell.rectColor = 'green';
      } else {
        outCell.rectTransparency = abs(1 - this.referenceCell.value / outCell.value);
        outCell.rectColor = 'green';
      }
    })

    if (n == 1) {
      return;
    }

    n = n - 1;
    cell.outputCells.forEach((outCell: CellProperties) => {
      this.addImpactInfoOutputCells(outCell, n);
    });

  }

  private checkIfCellHasSubtrahend(cell: CellProperties, subtrahend: string) {
    let color = 'green';

    if (cell.address.includes(subtrahend)) {
      color = 'red';
      return color;
    }

    cell.outputCells.forEach((outCell: CellProperties) => {
      this.checkIfCellHasSubtrahend(outCell, subtrahend);
    })

    return color;
  }

  private getDivisor(cell: CellProperties) {
    let divisor = 1;

    cell.inputCells.forEach((c: CellProperties) => {
      divisor += c.value;
    })
    return divisor;
  }
}