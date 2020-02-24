/* global console, Excel */
import CellOperations from '../celloperations';
import CellProperties from '../cellproperties';
import CommonOperations from './commonops';
import SheetProperties from '../sheetproperties';


export default class Likelihood {

  private cells: CellProperties[];
  private referenceCell: CellProperties;

  constructor(cells: CellProperties[], referenceCell: CellProperties) {
    this.cells = cells;
    this.referenceCell = referenceCell;
  }

  public addLikelihood() {

    this.addLikelihoodInfo();

    try {
      let commonOps = new CommonOperations();
      commonOps.drawRectangles(this.referenceCell.inputCells);
      commonOps.drawRectangles(this.referenceCell.outputCells);

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

  private addLikelihoodInfo() {

    try {
      for (let i = 0; i < this.cells.length; i++) {

        if (this.cells[i].isUncertain) {

          if (!SheetProperties.isImpact) {

            this.cells[i].rectColor = 'gray';
            this.cells[i].rectTransparency = 0;
          }

          this.cells[i].likelihood = this.cells[i + 2].value;

          console.log(this.cells[i].value + " has " + this.cells[i].likelihood);
        }
      }
    } catch (error) {
      console.log(error);
    }
  }
}