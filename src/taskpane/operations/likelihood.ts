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

  // at the moment it will be overwriting
  public showLikelihood(n: number = 1) {

    this.addLikelihoodInfo();

    try {

      const commonOps = new CommonOperations();

      if (SheetProperties.isImpact) {
        commonOps.deleteRectangles();
      }
      commonOps.drawRectangles(this.referenceCell.inputCells);
      this.showInputLikelihood(this.referenceCell, n);
      // this.commonOps.drawRectangles(this.referenceCell.inputCells);
      // this.commonOps.drawRectangles(this.referenceCell.outputCells);

    } catch (error) {
      console.log(error);
    }
  }

  public async removeLikelihood() {

    const commonOps = new CommonOperations();

    await commonOps.deleteRectangles();

    if (SheetProperties.isImpact) {

      commonOps.drawRectangles(this.referenceCell.inputCells);
      commonOps.drawRectangles(this.referenceCell.outputCells);
    }
  }

  private showInputLikelihood(cell: CellProperties, i: number) {

    cell.inputCells.forEach((inCell: CellProperties) => {
      console.log(inCell.address);
      this.commonOps.drawRectangles(inCell.inputCells);
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