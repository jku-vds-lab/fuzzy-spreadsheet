import CellOperations from "./celloperations";
import SheetProperties from "./sheetproperties";
import WhatIf from "./operations/whatif";
import Spread from "./operations/spread";

/* global console, Excel */

// Find a way to figure out which cells are uncertain so that we dont have to use their column index anymore
// maybe with the help of their formula?
export default class CellProperties {
  public id: string;
  public address: string;
  public value: number;
  public top: number;
  public left: number;
  public height: number;
  public width: number;
  public formula: any;
  public isFocus: boolean;
  public isUncertain: boolean = false;
  public degreeToFocus: number;
  public inputCells: CellProperties[];
  public outputCells: CellProperties[];
  public impact: number = 0;
  public likelihood: number = 10;
  public spreadRange: string;
  public variance: number = 0;
  public samples: number[];
  public isLineChart: boolean = false;
  // for impact and likelihood
  public rect: Excel.Shape;
  public rectColor: string;
  public rectTransparency: number;

  public isInputRelationship: boolean = false;
  public isOutputRelationship: boolean = false;
  public isImpact: boolean = false;
  public isLikelihood: boolean = false;
  public isSpread: boolean = false;
  public whatIf: WhatIf;

  private cells: CellProperties[];
  private newCells: CellProperties[];

  CellProperties() {
    this.id = "";
    this.address = "";
    this.value = 0;
    this.top = 0;
    this.left = 0;
    this.height = 0;
    this.width = 0;
    this.isFocus = false;
    this.likelihood = 100;
    this.degreeToFocus = -1;
    this.formula = "";
    this.spreadRange = null;
    this.isUncertain = false;
    this.whatIf = new WhatIf();
  }

  async getCells() {

    this.cells = new Array<CellProperties>();

    await Excel.run(async (context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();

      const range = sheet.getUsedRange(true);
      range.load(['formulas', 'values']);
      await context.sync();

      for (let i = 0; i < 20; i++) {
        for (let j = 0; j < 18; j++) {

          if (range.values[i][j] == "") {
            continue;
          }

          let cell = sheet.getCell(i, j);
          cell.load(["top", "left", "address"]); // compute these three as well

          await context.sync();

          let cellProperties = new CellProperties();
          cellProperties.id = "R" + i + "C" + j;
          cellProperties.address = cell.address;
          cellProperties.value = range.values[i][j];
          cellProperties.top = cell.top;
          cellProperties.left = cell.left;
          cellProperties.height = 15;
          cellProperties.width = 75.5;
          cellProperties.formula = range.formulas[i][j];
          cellProperties.degreeToFocus = -1;

          if (cellProperties.formula == cellProperties.value) {
            cellProperties.formula = "";
          }

          cellProperties.inputCells = new Array<CellProperties>();
          cellProperties.outputCells = new Array<CellProperties>();
          this.cells.push(cellProperties);
        }
      }
      // context.sync();
    });
    return this.cells;
  }

  updateNewValues(newValues: any[][], newFormulas: any[][], isUpdate: boolean = false) {

    try {

      this.newCells = new Array<CellProperties>();

      // make a deep copy of the element
      this.cells.forEach((cell: CellProperties) => {
        let newCell = new CellProperties();
        newCell = Object.assign(newCell, cell);
        newCell.id = cell.id;
        this.newCells.push(newCell);
      });

      this.newCells.forEach(function (newCell: CellProperties, index) {

        let id = newCell.id;
        id = id.replace('R', '');
        let c = id.indexOf('C');
        const rowIndex = id.substring(0, c);
        const colIndex = id.substring(c + 1);

        this[index].value = newValues[rowIndex][colIndex];
        this[index].formula = newFormulas[rowIndex][colIndex];
        if (this[index].formula == this[index].value) {
          this[index].formula = "";
        }
      }, this.newCells);

      this.checkUncertainty(this.newCells);
      // check if the reference cell is uncertain or not

      if (isUpdate) {
        console.log('Update everything');
        this.cells = this.newCells;

        this.getRelationshipOfCells();
      }
    } catch (error) {
      console.log('Error: ' + error);
    }
  }

  async calculateUpdatedNumber(referenceCell: CellProperties) {
    let i = 0;
    this.newCells.forEach(async (newCell: CellProperties) => {


      if (referenceCell.id == newCell.id) {
        console.log('New Reference Cell: ', newCell);
        referenceCell.whatIf = new WhatIf();
        referenceCell.whatIf.value = newCell.value - referenceCell.value;
        console.log('Change is value here: ' + referenceCell.whatIf.value);

        if (newCell.isUncertain) {
          newCell.variance = this.newCells[i + 1].value;
        }

        console.log('New cell variance: ' + newCell.variance);

        if (referenceCell.variance == newCell.variance) {
          console.log('No change in variance');
          return;
        } else {
          const spread = new Spread(this.newCells, newCell, 'MyCheatSheet');
          await spread.createCheatSheet();
          spread.showSpread(1);
        }
        // createNewGraph on the reference Cell
        return;
      }
      i++;
    })
  }

  private convertIdToIndices(id: string) {

    id = id.replace('R', '');
    let c = id.indexOf('C');
    let rowIndex = id.substring(0, c);
    let colIndex = id.substring(c + 1);

    return { rowIndex: rowIndex, colIndex: colIndex };
  }

  errorHandlerFunction(callback) {
    try {
      callback();
    } catch (error) {
      console.log(error);
    }
  }

  getRelationshipOfCells() {

    this.cells.forEach((cell: CellProperties) => {
      // eslint-disable-next-line no-empty
      if (cell.formula == "") {

      } else {

        let rangeAddresses: string[] = this.getRangeFromFormula(cell.formula);
        let cellRangeAddresses = new Array<string>();

        rangeAddresses.forEach((rangeAddress: string) => {
          rangeAddress = rangeAddress.trim();
          if (rangeAddress.includes(':')) {
            cellRangeAddresses = new Array<string>();

            let ranges = rangeAddress.split(':');
            if (ranges.length > 1) {
              let rangeStart = ranges[0];
              let rangeEnd = ranges[1];

              let colStartArray = rangeStart.match(/\d+/g);
              let colEndArray = rangeEnd.match(/\d+/g);
              let colStart = '';
              let colEnd = '';

              if (colStartArray != null) {
                colStart = colStartArray[0];
              }
              if (colEndArray != null) {
                colEnd = colEndArray[0];
              }

              let rowStart = rangeStart.replace(colStart, '');
              let rowEnd = rangeEnd.replace(colEnd, '');

              if (rowStart == rowEnd) {
                for (let i = Number(colStart); i <= Number(colEnd); i++) {
                  cellRangeAddresses.push(rowStart + i);
                }
              }
              else {
                let startIndex = rowStart.charCodeAt(0);
                let endIndex = rowEnd.charCodeAt(0);
                for (let i = startIndex; i <= endIndex; i++) {
                  const rowChar = String.fromCharCode(i);
                  cellRangeAddresses.push(rowChar + colStart);
                }
              }
            }

          }
          else {
            cellRangeAddresses = new Array<string>();
            cellRangeAddresses.push(rangeAddress);
          }

          const cellsFromRange = this.getCellsFromRangeAddress(cellRangeAddresses);


          cellsFromRange.forEach((cellInRange: CellProperties) => {
            cell.inputCells.push(cellInRange);
            cellInRange.outputCells.push(cell);
          })
        })
      }
    })
  }

  // can be optimised further
  private getCellsFromRangeAddress(cellRangeAddresses: string[]) {

    let cellsInRange = new Array<CellProperties>();

    for (let i = 0; i < cellRangeAddresses.length; i++) {
      for (let j = 0; j < this.cells.length; j++) {

        if (this.cells[j].address.includes(cellRangeAddresses[i])) {
          cellsInRange.push(this.cells[j]);
          break;
        }
      }
    }

    return cellsInRange;
  }

  getReferenceAndNeighbouringCells(referenceCellAddress: string) {
    let referenceCell = new CellProperties();

    this.cells.forEach((cell: CellProperties) => {
      if (cell.address == referenceCellAddress) {
        referenceCell = cell;
      }
    });

    referenceCell.degreeToFocus = 0;

    this.inCellsDegree(referenceCell.inputCells, 1);
    this.outCellsDegree(referenceCell.outputCells, 1);
    return referenceCell;
  }

  checkUncertainty(cells: CellProperties[]) {

    this.checkAverageValues(cells);

    cells.forEach((cell: CellProperties) => {

      //if input cells are uncertain then there sum or difference will also be uncertain
      if (cell.formula.includes("SUM")) {
        cell.isUncertain = this.checkAverageValues(cell.inputCells);
      }

      if (cell.formula.includes("-")) {
        let result = this.checkAverageValues(cell.inputCells);

        // if the first degree input cells to a difference cell are not uncertain, may be second degree might be uncertain
        if (!result) {
          cell.inputCells.forEach((iCell: CellProperties) => {
            result = this.checkAverageValues(iCell.inputCells);
          })
        }
        cell.isUncertain = result;
      }
    })
  }


  private checkAverageValues(cells: CellProperties[]) {
    let isUncertain = false;
    cells.forEach((cell: CellProperties) => {
      if (cell.formula.includes("AVERAGE") || cell.formula.includes('MITTELWERT')) { // because of german layout
        cell.isUncertain = true;
        isUncertain = true;
      }
    })
    return isUncertain;
  }


  // Need a proper solution
  private getRangeFromFormula(formula: string) {
    let rangeAddress = new Array<string>();
    if (formula == "") {
      return;
    }

    formula = formula.replace('(', '').replace(')', '').replace('=', '');

    if (formula.includes("SUM")) {
      let i = formula.indexOf("SUM");
      formula = formula.slice(i + 3);

      if (formula.includes(',')) {
        rangeAddress = formula.split(',');
      } else if (formula.includes(':')) {
        rangeAddress.push(formula);
      }
    }

    if (formula.includes("AVERAGE")) {
      let i = formula.indexOf("AVERAGE");
      formula = formula.slice(i + 7);

      if (formula.includes(',')) {
        rangeAddress = formula.split(',');

      } else if (formula.includes(':')) {
        rangeAddress.push(formula);
      }
    }

    if (formula.includes("-")) {
      rangeAddress = formula.split('-');
    }

    return rangeAddress;
  }

  private inCellsDegree(cells: CellProperties[], i: number) {

    cells.forEach((cell: CellProperties) => {
      let j = i;
      cell.degreeToFocus = j;
      if (cell.inputCells.length > 0) {
        j = i + 1;
        this.inCellsDegree(cell.inputCells, j);
      }
      j = i;
    });
  }

  private outCellsDegree(cells: CellProperties[], i: number) {

    cells.forEach((cell: CellProperties) => {
      let j = i;
      cell.degreeToFocus = j;
      if (cell.outputCells.length > 0) {
        j = i + 1;
        this.outCellsDegree(cell.outputCells, j);
      }
      j = i;
    });
  }
}