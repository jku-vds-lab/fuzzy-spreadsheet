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

      if (SheetProperties.isImpact) {
        console.log('Removing impact inputs');
        this.commonOps.deleteRectangles(this.cells, 'InputImpact')
      }

      this.displayInputLikelihood(this.referenceCell, n);

    } catch (error) {
      console.log(error);
    }
  }

  public showOutputLikelihood(n: number) {

    try {
      this.addLikelihoodInfo();

      if (SheetProperties.isImpact) {
        console.log('Removing impact outputs');
        this.commonOps.deleteRectangles(this.cells, 'OutputImpact')
      }

      this.displayOutputLikelihood(this.referenceCell, n);

    } catch (error) {
      console.log(error);
    }
  }

  public removeInputLikelihood(n: number) {
    try {
      const type = 'InputLikelihood';
      this.removeInputLikelihoodInfo(this.referenceCell, n);
      this.commonOps.deleteRectangles(this.cells, type);

      if (SheetProperties.isImpact && SheetProperties.isInputRelationship) {
        const impact = new Impact(this.referenceCell, this.cells);
        impact.redrawInputImpact(n);
      }

    } catch (error) {
      console.log(error);
    }
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


  public removeOutputLikelihood(n: number) {

    try {
      const type = 'OutputLikelihood';
      this.removeOutputLikelihoodInfo(this.referenceCell, n);
      this.commonOps.deleteRectangles(this.cells, type);

      if (SheetProperties.isImpact && SheetProperties.isOutputRelationship) {
        const impact = new Impact(this.referenceCell, this.cells);
        impact.redrawOutputImpact(n);
      }

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
      this.commonOps.drawRectangle(inCell, 'InputLikelihood');

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
      this.commonOps.drawRectangle(outCell, 'OutputLikelihood');

      if (n == 1) {
        return;
      }

      this.displayOutputLikelihood(outCell, n - 1);
    })
  }

  public redrawInputLikelihood(n: number) {

    this.commonOps.deleteRectangles(this.cells, 'InputLikelihood');
    this.removeInputLikelihoodInfo(this.referenceCell, n);
    this.addLikelihoodInfo();
    this.displayInputLikelihood(this.referenceCell, n);
  }

  public redrawOutputLikelihood(n: number) {

    this.commonOps.deleteRectangles(this.cells, 'OutputLikelihood');
    this.removeOutputLikelihoodInfo(this.referenceCell, n);
    this.addLikelihoodInfo();
    this.displayOutputLikelihood(this.referenceCell, n);
  }

  public removeAllLikelihoods() {
    this.cells.forEach((cell: CellProperties) => {
      cell.isLikelihood = false;
    })

    console.log('Removing all likelihood inputs');
    this.commonOps.deleteRectangles(this.cells, 'InputLikelihood');
    console.log('Removing all likelihood outputs');
    this.commonOps.deleteRectangles(this.cells, 'OutputLikelihood');

  }

  addLikelihoodInfo() {

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