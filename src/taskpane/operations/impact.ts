/* global console, Excel */
import { abs, round } from 'mathjs';
import CellProperties from '../cell/cellproperties';
import CommonOperations from './commonops';
import Bins from './bins';

export default class Impact {

  private referenceCell: CellProperties;
  private commonOps: CommonOperations;
  private inputCellsWithImpact: CellProperties[];
  private outputCellsWithImpact: CellProperties[];
  private redColors: string[];
  private blueColors: string[];

  constructor(referenceCell: CellProperties) {
    this.referenceCell = referenceCell;
    this.commonOps = new CommonOperations(this.referenceCell);
    this.redColors = Bins.getRedColorsForImpact();
    this.blueColors = Bins.getBlueColorsForImpact();
  }

  public showInputImpact(n: number, isDraw: boolean) {

    try {
      this.addImpactInfoInputCells(n);
      this.inputCellsWithImpact = new Array<CellProperties>();
      this.addInputImpact(this.referenceCell, n);

      if (isDraw) {
        this.commonOps.drawRectangle(this.inputCellsWithImpact, 'InputImpact');
      }
    } catch (error) {
      console.log(error);
    }
  }

  public showOutputImpact(n: number, isDraw: boolean) {

    try {
      this.addImpactInfoOutputCells(this.referenceCell, n);
      this.outputCellsWithImpact = new Array<CellProperties>();
      this.addOutputImpact(this.referenceCell, n);

      if (isDraw) {
        this.commonOps.drawRectangle(this.outputCellsWithImpact, 'OutputImpact');
      }
    } catch (error) {
      console.log(error);
    }
  }

  private addInputImpact(cell: CellProperties, i: number) {

    cell.inputCells.forEach((inCell: CellProperties) => {

      if (!inCell.isImpact) {

        this.inputCellsWithImpact.push(inCell);
        inCell.isImpact = true;
      }

      if (i == 1) {
        return;
      }
      this.addInputImpact(inCell, i - 1);
    })
  }

  private addOutputImpact(cell: CellProperties, i: number) {

    cell.outputCells.forEach((outCell: CellProperties) => {

      if (!outCell.isImpact) {

        this.outputCellsWithImpact.push(outCell);
        outCell.isImpact = true;
      }

      if (i == 1) {
        return;
      }
      this.addOutputImpact(outCell, i - 1);
    })
  }

  private addImpactInfoInputCells(n: number = 1) {

    const isreferenceCellDiff = this.referenceCell.formula.includes('-');
    const divisor = this.getDivisor(this.referenceCell);
    let subtrahend = null;


    if (isreferenceCellDiff) {
      let formula = this.referenceCell.formula;
      let idx = formula.indexOf('-');
      subtrahend = formula.substring(idx + 1, formula.length);
    }
    let isImpactPositive = true;
    const inputCells = this.referenceCell.inputCells;

    this.addInputImpactInfoRecursively(inputCells, n, isImpactPositive, divisor, subtrahend);
  }

  private addInputImpactInfoRecursively(inputCells: CellProperties[], n: number, isImpactPositive: boolean, divisor: number, subtrahend: string = null) {

    inputCells.forEach((inCell: CellProperties) => {

      if (inCell.address.includes(subtrahend)) {
        isImpactPositive = false;
      }

      const impact = inCell.value / divisor;

      inCell.impact = round(impact * 100, 2);
      inCell.isImpactPositive = isImpactPositive;
      let index = Math.ceil(inCell.impact);

      if (inCell.isImpactPositive) {
        inCell.rectColor = this.blueColors[index];
      } else {
        inCell.rectColor = this.redColors[index];
      }
    })

    if (n == 1) {
      return;
    }

    n = n - 1;

    inputCells.forEach((inCell: CellProperties) => {
      this.addInputImpactInfoRecursively(inCell.inputCells, n, inCell.isImpactPositive, divisor);
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

        let isImpactPositive = this.checkIfCellHasSubtrahend(cell, subtrahend);

        outCell.impact = round(impact * 100, 2);
        outCell.isImpactPositive = isImpactPositive;

        if (isImpactPositive) {
          outCell.rectColor = this.blueColors[Math.ceil(outCell.impact)];
        } else {
          outCell.rectColor = this.redColors[Math.ceil(outCell.impact)];
        }
      } else if (outCell.formula.includes('AVERAGE')) {
        const divisor = this.getDivisor(outCell);
        const impact = this.referenceCell.value / divisor;

        outCell.impact = round(impact * 100, 2);
        outCell.rectColor = this.blueColors[Math.ceil(outCell.impact)];
        outCell.isImpactPositive = true;
      } else {
        const impact = this.referenceCell.value / outCell.value;
        outCell.impact = round(impact * 100, 2);
        outCell.rectColor = this.blueColors[Math.ceil(outCell.impact)];
        outCell.isImpactPositive = true;
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
    let isSubtrahend = false;

    if (cell.address.includes(subtrahend)) {
      isSubtrahend = true;
    }

    cell.outputCells.forEach((outCell: CellProperties) => {
      this.checkIfCellHasSubtrahend(outCell, subtrahend);
    })

    let isImpactPositive = !isSubtrahend; // if it contains subtrahend, then the impact is negative

    return isImpactPositive;
  }

  private getDivisor(cell: CellProperties) {
    let divisor = 0;

    cell.inputCells.forEach((c: CellProperties) => {
      divisor += c.value;
    })
    return divisor;
  }
}