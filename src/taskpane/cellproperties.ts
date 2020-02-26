import CellOperations from "./celloperations";
import SheetProperties from "./sheetproperties";

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
  public likelihood: number = 10;
  public isInputRelationship: boolean;
  public isOutputRelationship: boolean;
  public spreadRange: string;
  public variance: number = 0;
  public samples: number[];
  public isLineChart: boolean = false;
  // for impact and likelihood
  public rect: Excel.Shape;
  public rectColor: string;
  public rectTransparency: number;

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
  }

  async getRangeProperties(referenceCell: CellProperties, cells: CellProperties[]) {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange(true);
      range.load(['formulas', 'values']);
      await context.sync();
      this.performLazyUpdate(referenceCell, cells, range.values, range.formulas);
    });
  }

  private performLazyUpdate(referenceCell: CellProperties, cells: CellProperties[], newValues: any[][], newFormulas: any[][]) {

    if (referenceCell == null) {
      return;
    }

    let indices = this.convertIdToIndices(referenceCell.id);
    let rowIndex = indices.rowIndex;
    let colIndex = indices.colIndex;
    let oldValue = referenceCell.value;

    if (referenceCell.value == newValues[rowIndex][colIndex] && referenceCell.formula == newFormulas[rowIndex][colIndex]) {

      if (referenceCell.variance == newValues[rowIndex][colIndex + 1]) {
        // perform no updates
        return;
      } else {
        console.log('Variance has changed');
        // Check if Spread is selected
        // recalculate samples of spread
        // check if cheat sheet exist
        // write them in the CheatSheet new range
        // draw the new graph with a different color
      }
    }

    let newValue = newValues[rowIndex][colIndex];

    SheetProperties.temp = newValue - oldValue;

    // otherwise perform an update
    cells.forEach((cell: CellProperties) => {

      indices = this.convertIdToIndices(cell.id);
      rowIndex = indices.rowIndex;
      colIndex = indices.colIndex;

      cell.value = newValues[rowIndex][colIndex];
      cell.formula = newFormulas[rowIndex][colIndex];
      if (cell.formula == cell.value) {
        cell.formula = "";
      }
    })
  }

  private convertIdToIndices(id: string) {

    id = id.replace('R', '');
    let c = id.indexOf('C');
    let rowIndex = id.substring(0, c);
    let colIndex = id.substring(c + 1);

    return { rowIndex: rowIndex, colIndex: colIndex };
  }

  async getCellsProperties(cells = new Array<CellProperties>()) {

    await Excel.run(async (context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // range.load(["formulas", "values"]);

      for (let i = 0; i < 20; i++) {
        for (let j = 0; j < 18; j++) {

          let cell = sheet.getCell(i, j);
          cell.load(["formulas", "top", "left", "height", "width", "address", "values"]);

          await context.sync();

          if (cell.values[0][0] == "") {
            continue;
          }

          let cellProperties = new CellProperties();
          cellProperties.id = "R" + i + "C" + j;
          cellProperties.address = cell.address;
          cellProperties.value = cell.values[0][0];
          cellProperties.top = cell.top;
          cellProperties.left = cell.left;
          cellProperties.height = cell.height;
          cellProperties.width = cell.width;
          cellProperties.formula = cell.formulas[0][0];
          cellProperties.degreeToFocus = -1;

          if (cellProperties.formula == cellProperties.value) {
            cellProperties.formula = "";
          }

          cellProperties.inputCells = new Array<CellProperties>();
          cellProperties.outputCells = new Array<CellProperties>();
          cells.push(cellProperties);
        }
      }
      // context.sync();
    });
    return cells;
  }

  errorHandlerFunction(callback) {
    try {
      callback();
    } catch (error) {
      console.log(error);
    }
  }

  getRelationshipOfCells(cells: CellProperties[]) {

    cells.forEach((cell: CellProperties) => {
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

          const cellsFromRange = this.getCellsFromRangeAddress(cells, cellRangeAddresses);


          cellsFromRange.forEach((cellInRange: CellProperties) => {
            cell.inputCells.push(cellInRange);
            cellInRange.outputCells.push(cell);
          })
        })
      }
    })
  }

  // can be optimised further
  private getCellsFromRangeAddress(cells: CellProperties[], cellRangeAddresses: string[]) {

    let cellsInRange = new Array<CellProperties>();

    for (let i = 0; i < cellRangeAddresses.length; i++) {
      for (let j = 0; j < cells.length; j++) {

        if (cells[j].address.includes(cellRangeAddresses[i])) {
          cellsInRange.push(cells[j]);
          break;
        }
      }
    }

    return cellsInRange;
  }

  getReferenceAndNeighbouringCells(cells: CellProperties[], focusCellAddress: string) {
    let focusCell = new CellProperties();

    cells.forEach((cell: CellProperties) => {
      if (cell.address == focusCellAddress) {
        focusCell = cell;
      }
    });

    focusCell.degreeToFocus = 0;

    this.inCellsDegree(focusCell.inputCells, 1);
    this.outCellsDegree(focusCell.outputCells, 1);
    return focusCell;
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