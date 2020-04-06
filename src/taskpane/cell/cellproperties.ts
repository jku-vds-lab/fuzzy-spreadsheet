/* global console, Excel */

export default class CellProperties {
  public id: string;
  public address: string;
  public value: number;
  public top: number;
  public left: number;
  public height: number;
  public width: number;
  public formula: string;
  public fontColor: string;
  public isFocus: boolean;
  public isUncertain: boolean = false;
  public degreeToFocus: number;
  public inputCells: CellProperties[];
  public outputCells: CellProperties[];
  public impact: number = 0;
  public likelihood: number = 1;
  public spreadRange: string;
  public stdev: number = 0;
  public computedMean: number = 0;
  public computedStdDev: number = 0;
  public samples: number[] = null;
  public samplesLikelihood: number[];
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
  public binBlueColors: string[];
  public binOrangeColors: string[];

  private cells: CellProperties[];
  private rowStart: number = 0;
  private rowEnd: number = 20;

  private colStart: number = 0;
  private colEnd: number = 18;

  // what if info
  public isTextbox: boolean = false;
  public updatedValue: number = 0;


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

  async getCells() {

    this.cells = new Array<CellProperties>();
    let cellRanges = new Array<Excel.Range>();
    let fontColors = new Array<Excel.RangeFont>();

    await Excel.run(async (context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();

      for (let i = this.rowStart; i < this.rowEnd; i++) {
        for (let j = this.colStart; j < this.colEnd; j++) {

          let cell = sheet.getCell(i, j);
          cellRanges.push(cell.load(["top", "left", "address", 'formulas', 'values']));
          fontColors.push(cell.format.font.load('color'));
        }
      }
      await context.sync();

      this.updateCellsValues(cellRanges, fontColors);

    });
    return this.cells;
  }

  async getCellsFormulasValues() {

    let cellRanges = new Array<Excel.Range>();
    let newValues = new Array<Array<any>>();
    let newFormulas = new Array<Array<any>>();
    try {
      await Excel.run(async (context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();

        for (let i = this.rowStart; i < this.rowEnd; i++) {
          for (let j = this.colStart; j < this.colEnd; j++) {

            let cell = sheet.getCell(i, j);
            cellRanges.push(cell.load(['formulas', 'values']));
          }
        }
        await context.sync();
      });

      cellRanges.forEach((cellRange: Excel.Range) => {
        newValues.push(cellRange.values);
        newFormulas.push(cellRange.formulas);
      })
    } catch (error) {
      console.log(error);
    }
    return { values: newValues, formulas: newFormulas };
  }


  updateCellsValues(cellRanges: Excel.Range[], fontColors: Excel.RangeFont[]) {

    let index = 0;
    for (let i = this.rowStart; i < this.rowEnd; i++) {
      for (let j = this.colStart; j < this.colEnd; j++) {

        if (cellRanges[index].values[0][0] == "") {
          index++;
          continue;
        }

        let cellProperties = new CellProperties();
        cellProperties.id = "R" + i + "C" + j;
        cellProperties.address = cellRanges[index].address;
        cellProperties.value = cellRanges[index].values[0][0];
        cellProperties.top = cellRanges[index].top;
        cellProperties.left = cellRanges[index].left;
        cellProperties.height = 15;
        cellProperties.width = 75.5;
        cellProperties.formula = cellRanges[index].formulas[0][0];
        cellProperties.degreeToFocus = -1;
        cellProperties.fontColor = fontColors[index].color;

        if (cellProperties.formula == cellProperties.value.toString()) {
          cellProperties.formula = "";
        }

        cellProperties.inputCells = new Array<CellProperties>();
        cellProperties.outputCells = new Array<CellProperties>();
        this.cells.push(cellProperties);
        index++;
      }
    }
  }

  public writeCellsToSheet(cells: CellProperties[]) {

    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let cellRanges = new Array<Excel.Range>();
      let cellValues = new Array<number>();
      let cellFormulas = new Array<any>();

      cells.forEach((cell: CellProperties) => {
        let range = sheet.getRange(cell.address);
        cellRanges.push(range.load(['values', 'formulas']));
        cellValues.push(cell.value);

        let formula = cell.formula;
        if (formula == "") {
          formula = cell.value.toString();
        }
        cellFormulas.push(formula);
      })

      await context.sync();

      let i = 0;

      cellRanges.forEach((cellRange: Excel.Range) => {
        cellRange.values = [[cellValues[i]]];
        cellRange.formulas = [[cellFormulas[i]]];
        i++;
      })
    });
  }

  public addVarianceAndLikelihoodInfo(cells: CellProperties[]) {

    try {
      for (let i = 0; i < this.cells.length; i++) {
        cells[i].stdev = 0;
        cells[i].likelihood = 1;

        if (cells[i].isUncertain) {

          cells[i].stdev = this.cells[i + 1].value;
          cells[i].likelihood = this.cells[i + 2].value;
        }
      }
    } catch (error) {
      console.log(error);
    }
  }


  errorHandlerFunction(callback: any) {
    try {
      callback();
    } catch (error) {
      console.log(error);
    }
  }

  getRelationshipOfCells(cells: CellProperties[] = this.cells) {

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