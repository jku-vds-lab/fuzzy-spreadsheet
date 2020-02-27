/* global console, Excel */
import CellOperations from '../celloperations';
import CellProperties from '../cellproperties';
import CommonOperations from './commonops';
import SheetProperties from '../sheetproperties';
import { increment } from 'src/functions/functions';
import Impact from './impact';


export default class Likelihood {

  private cells: CellProperties[];
  private referenceCell: CellProperties;
  private commonOps: CommonOperations;

  constructor(cells: CellProperties[], referenceCell: CellProperties) {
    this.cells = cells;
    this.referenceCell = referenceCell;
    this.commonOps = new CommonOperations();
  }


  public showLikelihood(n: number = 1) {

    this.addLikelihoodInfo();

    try {

      if (SheetProperties.isImpact) {
        this.commonOps.deleteRectangles();
      }

      this.showInputLikelihood(this.referenceCell, n);
      this.showOutputLikelihood(this.referenceCell, n);

    } catch (error) {
      console.log(error);
    }
  }

  public async removeLikelihood() {

    this.cells.forEach((cell: CellProperties) => {
      cell.isLikelihood = false;
    })
    await this.commonOps.deleteRectangles();

    if (SheetProperties.isImpact) {

      const impact = new Impact(this.referenceCell, this.cells);
      impact.showImpact();
    }
  }

  private showInputLikelihood(cell: CellProperties, i: number) {

    cell.inputCells.forEach((inCell: CellProperties) => {
      if (inCell.isLikelihood) {
        return;
      }

      inCell.isLikelihood = true;
      this.commonOps.drawRectangle(inCell);
      if (i == 1) {
        return;
      } else {
        this.showInputLikelihood(inCell, i - 1);
      }
    })
  }

  private showOutputLikelihood(cell: CellProperties, i: number) {

    cell.outputCells.forEach((outCell: CellProperties) => {
      if (outCell.isLikelihood) {
        return;
      }

      outCell.isLikelihood = true;
      this.commonOps.drawRectangle(outCell);
      if (i == 1) {
        return;
      } else {
        this.showOutputLikelihood(outCell, i - 1);
      }
    })
  }

  private addLikelihoodInfo() {

    try {
      for (let i = 0; i < this.cells.length; i++) {

        if (!SheetProperties.isImpact) {

          this.cells[i].rectColor = 'gray';
          this.cells[i].rectTransparency = 0;
        }

        if (this.cells[i].isUncertain) {
          this.cells[i].likelihood = this.cells[i + 2].value / 10;
        }
      }
    } catch (error) {
      console.log(error);
    }
  }
}