/* global console, Excel */
import CellOperations from '../celloperations';
import CellProperties from '../cellproperties';
import CommonOperations from './commonops';
import SheetProperties from '../sheetproperties';
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

  public showLikelihood(n: number, isInput: boolean, isOutput: boolean) {

    try {

      this.addLikelihoodInfo();

      if (isInput) {
        if (SheetProperties.isImpact) {
          this.commonOps.deleteRectangles(this.cells, 'Input');
        }
        this.showInputLikelihood(this.referenceCell, n);
      }

      if (isOutput) {
        if (SheetProperties.isImpact) {
          this.commonOps.deleteRectangles(this.cells, 'Output');
        }
        this.showOutputLikelihood(this.referenceCell, n);
      }
    } catch (error) {
      console.log(error);
    }
  }

  public async removeLikelihood(n: number, isInput: boolean, isOutput: boolean, isRemoveAll: boolean) {

    this.cells.forEach((cell: CellProperties) => {
      cell.isLikelihood = false;
    })

    if (!isInput) {
      await this.commonOps.deleteRectangles(this.cells, 'Input');
    }

    if (!isOutput) {
      await this.commonOps.deleteRectangles(this.cells, 'Output');
    }

    if (isRemoveAll) {

      await this.commonOps.deleteRectangles(this.cells, 'Input');
      await this.commonOps.deleteRectangles(this.cells, 'Output');
    }

    // if (SheetProperties.isImpact) {
    //   const impact = new Impact(this.referenceCell, this.cells);
    //   impact.showImpact(n, isInput, isOutput);
    // }
  }

  private showInputLikelihood(cell: CellProperties, n: number) {

    cell.inputCells.forEach((inCell: CellProperties) => {

      if (inCell.isLikelihood) {
        return;
      }

      inCell.isLikelihood = true;
      this.commonOps.drawRectangle(inCell, 'Input');

      if (n == 1) {
        return;
      }

      this.showInputLikelihood(inCell, n - 1);
    })
  }

  private showOutputLikelihood(cell: CellProperties, n: number) {

    cell.outputCells.forEach((outCell: CellProperties) => {

      if (outCell.isLikelihood) {
        return;
      }

      outCell.isLikelihood = true;
      this.commonOps.drawRectangle(outCell, 'Output');

      if (n == 1) {
        return;
      }

      this.showOutputLikelihood(outCell, n - 1);
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
          this.cells[i].likelihood = this.cells[i + 2].value;
        }
      }
    } catch (error) {
      console.log(error);
    }
  }
}