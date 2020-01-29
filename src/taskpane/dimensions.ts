/* global console, Excel */
import CellProperties from './cellproperties';
import ShapeProperties from './shapeproperties';

export default class Dimensions {
  chartType: string;
  cells: CellProperties[];
  shapes: ShapeProperties[];
  MyDimension() { }
  getCells() {
    return this.cells;
  }
  getShapes() {
    return this.shapes;
  }

  static async getAddressFromRange(rowIndex: number, colIndex: number, rowCount: number, colCount: number) {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      for (let i = rowIndex; i < rowIndex + rowCount; i++) {
        for (let j = colIndex; j < colIndex + colCount; j++) {
          let cell = sheet.getCell(i, j);
          console.log("Fetch: Row and Column: " + i + " " + j);
          cell.load(["address"]);
          await context.sync();
          // console.log("Address: " + cell.address);
          // await context.sync();
        }
      }
    });
  }
  static getCellsFromFormula(formula: string) {
    let inCellAddresses = new Array<string>();
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(formula);
      range.load(["columnIndex", "rowIndex", "columnCount", "rowCount"]);
      await context.sync();
      const rowIndex = range.rowIndex;
      const colIndex = range.columnIndex;
      const rowCount = range.rowCount;
      const colCount = range.columnCount;
      console.log("Formula: " + formula);
      console.log("Index: Rows and Columns: " + rowIndex + " " + colIndex);
      console.log("Count: Rows and Columns: " + rowCount + " " + colCount);
      await Dimensions.getAddressFromRange(rowIndex, colIndex, rowCount, colCount);
      // await context.sync();
      console.log("------------------------------");
    });
    return inCellAddresses;
  }
  static scanRangeForRelations() {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // const range = sheet.getUsedRange(true);
      const range = sheet.getRange("I6:I9");
      range.load(["columnIndex", "rowIndex", "columnCount", "rowCount"]);
      await context.sync();
      const rowIndex = range.rowIndex;
      const colIndex = range.columnIndex;
      const rowCount = range.rowCount;
      const colCount = range.columnCount;
      let cells = new Array<CellProperties>();
      for (let i = rowIndex; i < rowIndex + rowCount; i++) {
        for (let j = colIndex; j < colIndex + colCount; j++) {
          let cell = sheet.getCell(i, j);
          cell.load(["formulas", "top", "left", "height", "width", "address", "values"]);
          await context.sync();
          // eslint-disable-next-line no-empty
          if (cell.values[0][0] == "") {
          } else {
            let cellProp = new CellProperties();
            cellProp.cell = cell.address;
            cellProp.value = cell.values[0][0];
            cellProp.top = cell.top;
            cellProp.left = cell.left;
            cellProp.height = cell.height;
            cellProp.width = cell.width;
            cellProp.formula = cell.formulas[0][0];
            if (cellProp.formula == cellProp.value) {
              cellProp.formula = "";
            }
            cellProp.inCells = new Array<CellProperties>();
            cellProp.outCells = new Array<CellProperties>();
            cells.push(cellProp);

          }
        }
      }
      // define input/output cells now
      Dimensions.insertRelations(cells);
      await context.sync();
    });
  }
  static insertRelations(cells: CellProperties[]) {
    cells.forEach((cell: CellProperties) => {
      // eslint-disable-next-line no-empty
      if (cell.formula == "") {
      } else {
        let rangeAddress = Dimensions.getRangeFromFormula(cell.formula);
        // let inCellAddresses: string[] = new Array<string>();
        Dimensions.getCellsFromFormula(rangeAddress);
        // console.log("Checking for matches");
        // console.log("First: " + inCellAddresses[0]);
        // for (let address in inCellAddresses) {
        //   console.log("In Cell Address: ", address);
        // }
        // inCellAddresses.forEach((address: string) => {
        //   cells.forEach((c: CellProperties) => {
        //     console.log("Cell Address: ", c.cell);
        //     console.log("In Cell Address: ", address);
        //     if (c.cell == address) {
        //       console.log("Found a match");
        //       // c.outCells.push(cell);
        //       // cell.inCells.push(c);
        //       return;
        //     }
        //   });
        // });
      }
    });
  }
  async scanRange(cellAddresses: string[], focusCell: string) {
    this.cells = new Array<CellProperties>();
    for (let i = 0; i < cellAddresses.length; i++) {
      let degreeToFocus = 2;
      if (i == 0 || i == 4) {
        degreeToFocus = 1;
      }
      this.cells.push(await new CellProperties().getCellProperties(cellAddresses[i], focusCell, degreeToFocus));
    }
  }
  addSpreadInfo() {
    Excel.run(async (context) => {
      const cheatsheet = context.workbook.worksheets.add("CheatSheet");
      let data: number[][] = new Array<Array<number>>();
      let means = [32, 13, 7, 12, 26.6, 0.6, 1, 9, 9, 7]; // make it dynamic
      let stdDev = [6.38, 2.5, 2.9, 1.8, 4.8, 0.2, 0.4, 2.7, 2.2, 1.34]; // make it dynamic
      for (let i = 0; i < 47; i++) {
        let row = new Array<number>();
        for (let j = 0; j < 10; j++) {
          var normalVal = context.workbook.functions.norm_Dist(i + 1, means[j], stdDev[j], false);
          normalVal.load("value");
          await context.sync();
          row.push(normalVal.value);
        }
        data.push(row);
      }
      var range = cheatsheet.getRange("A1:J47");
      range.values = data;
      await context.sync();
    });
  }
  addImpactInfo(cells: CellProperties[]) {
    this.cells = cells;
    let color = "green";
    let transparency = 0;
    let height = 5;
    let width = 5;
    let firstDegreeDivisor = -1;
    let secondDegreeDivisor = -1;
    this.shapes = new Array<ShapeProperties>();
    // Finding the firstDegreeDivisor
    this.cells.forEach((cell: CellProperties) => {
      let val = cell.value;
      if (val < 0) {
        val = -cell.value;
      }
      if (cell.degreeToFocus == 1 && val > firstDegreeDivisor) {
        firstDegreeDivisor = val;
      }
    });
    // Finding the secondDegreeDivisor & assigning shape properties
    this.cells.forEach((cell: CellProperties) => {
      let val = cell.value;
      if (val < 0) {
        val = -cell.value;
      }
      if (cell.value < 0) {
        color = "red";
      }
      if (cell.degreeToFocus == 1) {
        secondDegreeDivisor = val;
        transparency = 1 - val / firstDegreeDivisor;
      } else if (cell.degreeToFocus == 2) {
        transparency = 1 - val / secondDegreeDivisor;
      }
      this.shapes.push(
        new ShapeProperties().getShapeProperties(Excel.GeometricShapeType.rectangle, color, transparency, height, width)
      );
    });
  }
  async addLikelihoodInfo(cells: CellProperties[], shapes: ShapeProperties[], likelihoodAddresses: string[]) {
    this.cells = cells;
    this.shapes = shapes;
    let likelihoodCell: number;
    if (this.shapes.length > 0) {
      if (this.shapes.length != likelihoodAddresses.length) {
        return;
      }
      for (let i = 0; i < this.shapes.length; i++) {
        likelihoodCell = await new CellProperties().getCellValue(likelihoodAddresses[i]);
        this.shapes[i].height = likelihoodCell / 10;
        this.shapes[i].width = likelihoodCell / 10;
      }
    }
  }
  static getRangeFromFormula(formula: string) {
    let rangeAddress = "";
    if (formula == "") {
      return;
    }
    if (formula.includes("SUM")) {
      let i = formula.indexOf("SUM");
      rangeAddress = formula.slice(i + 3);
      // rangeAddress = formula.replace("= SUM", "");
    }
    if (formula.includes("MEDIAN")) {
      let i = formula.indexOf("MEDIAN");
      rangeAddress = formula.slice(i + 6);
      // rangeAddress = formula.replace("= MEDIAN", "");
    }
    return rangeAddress;
  }
  addRelationshipInfo(cells: CellProperties[]) {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      for (let i = 0; i < cells.length; i++) {
        cells[i].outputCells = new Array<string>();
        cells[i].inputCells = new Array<Excel.Range>();
        let rangeAddress: string = Dimensions.getRangeFromFormula(cells[i].formula);
        if (rangeAddress.includes(",")) {
          let splits = rangeAddress.split(",");
          for (var split in splits) {
            cells[i].inputCells.push(sheet.getRange(split));
            // cells[i].inputCells.format.fill.color = "orange";
          }
        }
        cells[i].inputCells.forEach((c: Excel.Range) => {
          // c.format.fill.color("orange");
          c.format.fill.color = "orange";
        });
        for (let j = 0; j < cells.length; j++) {
          if (cells[i].cell == cells[j].cell) {
            continue;
          }
          rangeAddress = Dimensions.getRangeFromFormula(cells[j].formula);
          if (rangeAddress.includes(",")) {
            continue;
          } // not checking it at the moment, but also needs to be included
          let range = sheet.getRange(rangeAddress);
          let checkIntersection = range.getIntersectionOrNullObject(cells[i].cell);
          checkIntersection.load("address");
          await context.sync().then(function () {
            if (checkIntersection.address) {
              cells[i].outputCells.push(cells[j].cell);
            }
          });
        }
        console.log("Output Cells for  " + cells[i].cell, cells[i].outputCells);
      }
      await context.sync().then(function () { });
    }).catch(function (error) {
      console.log("Error: " + JSON.stringify(error.debugInfo));
    });
  }
}