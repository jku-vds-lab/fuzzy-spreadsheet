/* global console, Excel */

// TO-DO's

// fix the red issue
// work on output
// remove duplicates

import { abs, round } from 'mathjs';
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
      inCell.isImpact = false;
    })

    this.referenceCell.outputCells.forEach((outCell: CellProperties) => {
      outCell.rectColor = color;
      outCell.rectTransparency = transparency;
      outCell.isImpact = false;
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
      const impact = inCell.value / divisor;
      inCell.impact = round(impact * 100, 2);
      inCell.rectTransparency = abs(1 - impact);
      inCell.isImpact = true;
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

        const impact = this.referenceCell.value / divisor;
        outCell.rectColor = this.checkIfCellHasSubtrahend(cell, subtrahend);
        outCell.rectTransparency = abs(1 - impact);
        outCell.impact = round(impact * 100, 2);


      } else if (outCell.formula.includes('AVERAGE')) {
        const divisor = this.getDivisor(outCell);
        const impact = this.referenceCell.value / divisor;

        outCell.rectTransparency = abs(1 - impact);
        outCell.rectColor = 'green';
        outCell.impact = round(impact * 100, 2);
      } else {
        const impact = this.referenceCell.value / outCell.value;
        outCell.impact = round(impact * 100, 2);
        outCell.rectTransparency = abs(1 - impact);
        outCell.rectColor = 'green';
      }
      outCell.isImpact = true;
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