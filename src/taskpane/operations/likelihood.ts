/* global console, Excel */
import CellOperations from '../celloperations';
import CellProperties from '../cellproperties';
import CommonOperations from './commonops';
import SheetProperties from '../sheetproperties';


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

    await this.commonOps.deleteRectangles();

    if (SheetProperties.isImpact) {

      this.commonOps.drawRectangles(this.referenceCell.inputCells);
      this.commonOps.drawRectangles(this.referenceCell.outputCells);
    }
  }

  private showInputLikelihood(cell: CellProperties, i: number) {

    this.commonOps.drawRectangles(cell.inputCells);

    if (i == 1) {
      return;
    }

    cell.inputCells.forEach((inCell: CellProperties) => {
      this.showInputLikelihood(inCell, i - 1);
    })
  }

  private showOutputLikelihood(cell: CellProperties, i: number) {
    this.commonOps.drawRectangles(cell.outputCells);

    if (i == 1) {
      return;
    }

    cell.outputCells.forEach((outCell: CellProperties) => {
      this.showOutputLikelihood(outCell, i - 1);
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