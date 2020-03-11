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

  public showInputLikelihood(n: number) {

    try {
      this.addLikelihoodInfo();
      // if (SheetProperties.isImpact) {
      //   this.commonOps.deleteRectangles(this.cells, 'Input');
      // }
      this.displayInputLikelihood(this.referenceCell, n);
    } catch (error) {
      console.log(error);
    }
  }

  public showOutputLikelihood(n: number) {

    try {
      this.addLikelihoodInfo();
      // if (SheetProperties.isImpact) {
      //   this.commonOps.deleteRectangles(this.cells, 'Output');
      // }
      this.displayOutputLikelihood(this.referenceCell, n);
    } catch (error) {
      console.log(error);
    }
  }

  public async removeInputLikelihood(n: number) {
    try {
      const type = 'Input';
      this.removeInputLikelihoodInfo(this.referenceCell, n);
      await this.commonOps.deleteRectangles(this.cells, type);
    } catch (error) {
      console.log(error);
    }

    // if (SheetProperties.isImpact) {
    //   const impact = new Impact(this.referenceCell, this.cells);
    //   impact.showImpact(n, isInput, isOutput);
    // }
  }

  private removeInputLikelihoodInfo(cell: CellProperties, n: number) {

    cell.inputCells.forEach((inCell: CellProperties) => {

      if (inCell.isLikelihood) {
        inCell.isLikelihood = false;
      }

      if (n == 1) {
        return;
      }
      this.removeInputLikelihoodInfo(inCell, n - 1);
    })
  }


  public async removeOutputLikelihood(n: number) {

    try {
      const type = 'Output';
      this.removeOutputLikelihoodInfo(this.referenceCell, n);
      await this.commonOps.deleteRectangles(this.cells, type);
    } catch (error) {
      console.log(error);
    }

  }

  private removeOutputLikelihoodInfo(cell: CellProperties, n: number) {

    cell.outputCells.forEach((outCell: CellProperties) => {

      if (outCell.isLikelihood) {
        outCell.isLikelihood = false;
      }

      if (n == 1) {
        return;
      }
      this.removeOutputLikelihoodInfo(outCell, n - 1);
    })
  }

  private displayInputLikelihood(cell: CellProperties, n: number) {

    cell.inputCells.forEach((inCell: CellProperties) => {

      if (inCell.isLikelihood) {
        return;
      }

      inCell.isLikelihood = true;
      this.commonOps.drawRectangle(inCell, 'Input');

      if (n == 1) {
        return;
      }

      this.displayInputLikelihood(inCell, n - 1);
    })
  }

  private displayOutputLikelihood(cell: CellProperties, n: number) {

    cell.outputCells.forEach((outCell: CellProperties) => {

      if (outCell.isLikelihood) {
        return;
      }

      outCell.isLikelihood = true;
      this.commonOps.drawRectangle(outCell, 'Output');

      if (n == 1) {
        return;
      }

      this.displayOutputLikelihood(outCell, n - 1);
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