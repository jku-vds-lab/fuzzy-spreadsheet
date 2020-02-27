/* global console, Excel */
import { abs, round } from 'mathjs';
import CellProperties from '../cellproperties';
import CommonOperations from './commonops';
import SheetProperties from '../sheetproperties';
import Likelihood from './likelihood';

export default class Impact {

  private referenceCell: CellProperties;
  private cells: CellProperties[];
  private commonOps: CommonOperations;

  constructor(referenceCell: CellProperties, cells: CellProperties[]) {
    this.referenceCell = referenceCell;
    this.commonOps = new CommonOperations();
    this.cells = cells;
  }

  public showImpact(n: number) {

    this.addImpactInfo(n);

    if (SheetProperties.isLikelihood) {
      this.commonOps.deleteRectangles(this.cells);
    }

    this.showInputImpact(this.referenceCell, n);
    this.showOutputImpact(this.referenceCell, n);
  }

  public async removeImpact(n: number) {

    this.cells.forEach((cell: CellProperties) => {
      cell.isImpact = false;
    })

    await this.commonOps.deleteRectangles(this.cells);

    if (SheetProperties.isLikelihood) {

      const likelihood = new Likelihood(this.cells, this.referenceCell);
      likelihood.showLikelihood(n);
    }
  }

  private showInputImpact(cell: CellProperties, i: number) {

    cell.inputCells.forEach((inCell: CellProperties) => {

      if (inCell.isImpact) {
        return;
      }

      inCell.isImpact = true;
      this.commonOps.drawRectangle(inCell);

      if (i == 1) {
        return;
      }
      this.showInputImpact(inCell, i - 1);
    })
  }

  private showOutputImpact(cell: CellProperties, i: number) {

    cell.outputCells.forEach((outCell: CellProperties) => {

      if (outCell.isImpact) {
        return;
      }

      outCell.isImpact = true;
      this.commonOps.drawRectangle(outCell);

      if (i == 1) {
        return;
      }
      this.showOutputImpact(outCell, i - 1);
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