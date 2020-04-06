/* global console, Excel */
import CellProperties from '../cell/cellproperties';
import CommonOperations from './commonops';

export default class Likelihood {

  private referenceCell: CellProperties;
  private commonOps: CommonOperations;
  private inputCellsWithLikelihood: CellProperties[];
  private outputCellsWithLikelihood: CellProperties[];

  constructor(referenceCell: CellProperties) {
    this.referenceCell = referenceCell;
    this.commonOps = new CommonOperations(this.referenceCell);
  }

  public showInputLikelihood(n: number, isDraw: boolean) {

    try {
      this.inputCellsWithLikelihood = new Array<CellProperties>();
      this.addInputLikelihood(this.referenceCell, n);
      if (isDraw) {
        this.commonOps.drawRectangle(this.inputCellsWithLikelihood, 'InputLikelihood');
      }
    } catch (error) {
      console.log(error);
    }
  }

  public showOutputLikelihood(n: number, isDraw: boolean) {

    try {
      this.outputCellsWithLikelihood = new Array<CellProperties>();
      this.addOutputLikelihood(this.referenceCell, n);
      if (isDraw) {
        this.commonOps.drawRectangle(this.outputCellsWithLikelihood, 'OutputLikelihood');
      }
    } catch (error) {
      console.log(error);
    }
  }

  private addInputLikelihood(cell: CellProperties, n: number) {

    cell.inputCells.forEach((inCell: CellProperties) => {

      if (!inCell.isLikelihood) {
        this.inputCellsWithLikelihood.push(inCell);
        inCell.isLikelihood = true;
      }

      if (n == 1) {
        return;
      }

      this.addInputLikelihood(inCell, n - 1);
    })
  }

  private addOutputLikelihood(cell: CellProperties, n: number) {

    cell.outputCells.forEach((outCell: CellProperties) => {

      if (!outCell.isLikelihood) {
        this.outputCellsWithLikelihood.push(outCell);
        outCell.isLikelihood = true;
      }

      if (n == 1) {
        return;
      }

      this.addOutputLikelihood(outCell, n - 1);
    })
  }
}