import Dimensions from './testfile';
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = addImpact;
  document.getElementById("impact").onclick = run;

  //   document.getElementById("run").onclick = run;
  //   document.getElementById("run").onclick = run;

  // $("#remove-impact").click(() =>
  //   tryCatch(function () {
  //     addLikelihood(false);
  //   })
  // );
  // $("#remove-distributions").click(() => tryCatch(removeDistributions));
  // $("#remove-likelihood").click(() => tryCatch(removeLikelihood));
  // $("#protect-sheet").click(() =>
  //   tryCatch(MyDimensions.scanRangeForRelations));
  // $("#unprotect-sheet").click(() => tryCatch(unprotectSheet));
  // };

  async function run() {
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }
  class CellProperties {
    public cell: string;
    public value: number;
    public top: number;
    public left: number;
    public height: number;
    public width: number;
    public isFocus: boolean;
    public degreeToFocus: number;
    public formula: any;
    public inputCells: Excel.Range[];
    public outputCells: string[];
    public inCells: CellProperties[];
    public outCells: CellProperties[];
    CellProperties() {
      this.cell = "";
      this.value = 0;
      this.top = 0;
      this.left = 0;
      this.height = 0;
      this.width = 0;
      this.isFocus = false;
      this.degreeToFocus = 0;
      this.formula = "";
      this.outputCells = new Array<string>();
    }
    getCellProperties(cellAddress: string, focusCell: string, degreeToFocus: number) {
      this.cell = cellAddress;
      this.isFocus = false;
      if (cellAddress == focusCell) {
        this.isFocus = true;
        this.degreeToFocus = 0;
      }
      Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Probability");
        const cell = sheet.getRange(cellAddress);
        cell.load(["values", "top", "left", "height", "width", "formulas"]);
        await context.sync();
        this.value = cell.values[0][0]; // gets the current cell value
        this.top = cell.top;
        this.left = cell.left;
        this.height = cell.height;
        this.width = cell.width;
        this.degreeToFocus = degreeToFocus;
        this.formula = cell.formulas[0][0]; // gets the formula of the current cell
        await context.sync();
      });
      return this;
    }
    async getCellValue(cellAddress: string) {
      let value: number = 0;
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Probability");
        const cell = sheet.getRange(cellAddress);
        cell.load("values");
        return context.sync().then(function () {
          value = cell.values[0][0]; // gets the current cell value
        });
      });
      return value;
    }
  }
  class ShapeProperties {
    shapeType: string;
    color: string;
    transparency: number;
    height: number;
    width: number;
    getShapeProperties(shapeType: string, color: string, transparency: number, height: number, width: number) {
      this.shapeType = shapeType;
      this.color = color;
      this.transparency = transparency;
      this.height = height;
      this.width = width;
      return this;
    }
  }
  // C:\Windows\SysWOW64\F12
  class MyDimensions {
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

    static wtf() {
      console.log("WTF");
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
        await MyDimensions.getAddressFromRange(rowIndex, colIndex, rowCount, colCount);
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
        MyDimensions.insertRelations(cells);
        await context.sync();
      });
    }
    static insertRelations(cells: CellProperties[]) {
      cells.forEach((cell: CellProperties) => {
        // eslint-disable-next-line no-empty
        if (cell.formula == "") {
        } else {
          let rangeAddress = MyDimensions.getRangeFromFormula(cell.formula);
          // let inCellAddresses: string[] = new Array<string>();
          MyDimensions.getCellsFromFormula(rangeAddress);
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
    scanRange(cellAddresses: string[], focusCell: string) {
      this.cells = new Array<CellProperties>();
      for (let i = 0; i < cellAddresses.length; i++) {
        let degreeToFocus = 2;
        if (i == 0 || i == 4) {
          degreeToFocus = 1;
        }
        this.cells.push(new CellProperties().getCellProperties(cellAddresses[i], focusCell, degreeToFocus));
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
        if (this.shapes.length != likehoodAddresses.length) {
          return;
        }
        for (let i = 0; i < this.shapes.length; i++) {
          likelihoodCell = await new CellProperties().getCellValue(likehoodAddresses[i]);
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
          let rangeAddress: string = MyDimensions.getRangeFromFormula(cells[i].formula);
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
            rangeAddress = MyDimensions.getRangeFromFormula(cells[j].formula);
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

  let dim = new MyDimensions();
  let cellAddresses = ["I6", "I7", "I8", "I9", "I11", "I12", "I13", "I14",
    "I15", "I16"];
  let likehoodAddresses = ["J6", "J7", "J8", "J9", "J11", "J12", "J13", "J14",
    "J15", "J16"];
  dim.scanRange(cellAddresses, "I18");
  let cells = dim.getCells();
  let shapes = dim.getShapes();

  // $("#relationship").click(() => tryCatch(addRelationship));
  // $("#impact").click(() => tryCatch(addImpact));
  // $("#likelihood").click(() => tryCatch(addLikelihood));
  // $("#line-spread").click(() =>
  //   tryCatch(function () {
  //     addSpread(true);
  //   })
  // );
  // $("#column-spread").click(() =>
  //   tryCatch(function () {
  //     addSpread(false);
  //   })
  // );
  async function addRelationship() {
    let cells = dim.getCells();
    dim.addRelationshipInfo(cells);
  }
  async function addSpread(isLine = false) {
    let cells = dim.getCells();
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      const cheatSheet = context.workbook.worksheets.getItem("CheatSheet");
      // make it dynamic
      let ranges: string[] = [
        "A1:A47",
        "B1:B47",
        "C1:C47",
        "D1:D47",
        "E1:E47",
        "F1:F47",
        "G1:G47",
        "H1:H47",
        "I1:I47",
        "J1:J47"
      ];
      for (let i = 0; i < ranges.length; i++) {
        const dataRange = cheatSheet.getRange(ranges[i]);
        let chart: Excel.Chart;
        if (isLine) {
          chart = sheet.charts.add("Line", dataRange, Excel.ChartSeriesBy.auto);
        } else {
          chart = sheet.charts.add("ColumnClustered", dataRange, Excel.ChartSeriesBy.auto);
        }
        chart.setPosition(cells[i].cell, cells[i].cell);
        chart.left = cells[i].left + 0.2 * cells[i].width;
        chart.title.visible = false;
        chart.legend.visible = false;
        chart.axes.valueAxis.minimum = 0;
        chart.axes.valueAxis.maximum = 0.21;
        chart.dataLabels.showValue = false;
        chart.axes.valueAxis.visible = false;
        chart.axes.categoryAxis.visible = false;
        chart.axes.valueAxis.majorGridlines.visible = false;
        chart.plotArea.top = 0;
        chart.plotArea.left = 0;
        chart.plotArea.width = cells[i].width - 0.4 * cells[i].width;
        chart.plotArea.height = 100;
        chart.format.fill.clear();
        chart.format.border.clear();
      }
      await context.sync();
    });
  }
  // Not possible without impact yet
  async function addLikelihood(isImpact: boolean = true) {
    let shapes = dim.getShapes();
    await dim.addLikelihoodInfo(cells, shapes, likehoodAddresses);
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      for (let i = 0; i < shapes.length; i++) {
        var shape = sheet.shapes.getItem("Impact" + i);
        console.log(shape);
        shape.load(["geometricShapeType", "width", "height"]);
        await context.sync();
        console.log("Geometric Shape Type: " + shape.geometricShapeType);
        if (shape.geometricShapeType == Excel.GeometricShapeType.rectangle) {
          console.log("Rectangle Found");
          shape.height = shapes[i].height;
          shape.width = shapes[i].width;
        }
      }
      createLikelihoodLegend().then(function () { });
      await context.sync();
    });
  }
  async function addImpact() {
    // console.log("Here I am ");
    // let dim = new MyDimensions();
    // let cellAddresses = ["I6", "I7", "I8", "I9", "I11", "I12", "I13", "I14", "I15", "I16"];
    // dim.scanRange(cellAddresses, "I18");
    // let cells = dim.getCells();
    // dim.addImpactInfo(cells);
    // let shapes = dim.getShapes();

    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */

        console.log("Here I am ");
        debugger;
        Dimensions.myFunc();
        // let cellAddresses = ["I6", "I7", "I8", "I9", "I11", "I12", "I13", "I14", "I15", "I16"];
        // dim.scanRange(cellAddresses, "I18");
        // let cells = dim.getCells();
        // dim.addImpactInfo(cells);
        // let shapes = dim.getShapes();
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }

    // Ensure cells and shapes are the same length
    // await Excel.run(async (context) => {
    //   const sheet = context.workbook.worksheets.getItem("Probability");
    //   for (let i = 0; i < cells.length; i++) {
    //     var impact = sheet.shapes.addGeometricShape("Rectangle"); // shapes[i].shapeType
    //     impact.name = "Impact" + i;
    //     impact.height = shapes[i].height;
    //     impact.width = shapes[i].width;
    //     impact.left = cells[i].left + 2;
    //     impact.top = cells[i].top + cells[i].height / 4;
    //     impact.rotation = 0;
    //     impact.fill.transparency = shapes[i].transparency;
    //     impact.lineFormat.weight = 0;
    //     impact.lineFormat.color = shapes[i].color;
    //     impact.fill.setSolidColor(shapes[i].color);
    //   }
    //   createImpactLegend().then(function () { });
    //   await context.sync();
    // });
  }
  async function protectSheet() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      sheet.load("protection/protected");
      await context.sync().then(function () {
        if (!sheet.protection.protected) {
          sheet.protection.protect();
        }
      });
    });
  }
  async function unprotectSheet() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      sheet.load("protection/protected");
      await context.sync().then(function () {
        if (sheet.protection.protected) {
          sheet.protection.unprotect();
        }
      });
    });
  }
  async function removeLikelihood() {
    // To be fixed
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      const count = sheet.shapes.getCount();
      await context.sync();
      for (let i = 0; i < count.value; i++) {
        var shape = sheet.shapes.getItemAt(i);
        shape.load(["geometricShapeType", "width", "height"]);
        await context.sync();
        if (shape.geometricShapeType == Excel.GeometricShapeType.rectangle) {
          shape.width = 7;
          shape.height = 7;
        }
      }
      await context.sync();
    });
  }
  async function createLikelihoodLegend() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      const textRange = ["    < 50", "    <= 80", "    <= 100"];
      const sizeRange = [5, 7, 9];
      let color = "gray";
      for (let i = 0; i < 3; i++) {
        var legend = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
        var cell = sheet.getCell(i + 22, 4);
        cell.load("top");
        cell.load("left");
        cell.load("height");
        cell.load("values");
        await context.sync();
        legend.height = sizeRange[i];
        legend.width = sizeRange[i];
        legend.left = cell.left + 2;
        legend.top = cell.top + cell.height / 4;
        legend.lineFormat.weight = 0;
        legend.lineFormat.color = color;
        legend.fill.setSolidColor(color);
        cell.values = [[textRange[i]]];
      }
      await context.sync();
    });
  }
  async function createImpactLegend() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      const textRange = ["    > 20", "    >= 9 & < 20", "    < 9", "    < 9", "    >= 9 & < 20", "    > 20"];
      const transparencyRange = [0, 0.4, 0.7, 0.7, 0.4, 0];
      let color = "green";
      for (let i = 0; i < 6; i++) {
        if (i == 3) {
          color = "red";
        }
        var legend = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
        var cell = sheet.getCell(i + 22, 2);

        cell.load("top");
        cell.load("left");
        cell.load("height");
        cell.load("values");
        await context.sync();
        legend.height = 7;
        legend.width = 7;
        legend.left = cell.left + 2;
        legend.top = cell.top + cell.height / 4;
        legend.lineFormat.weight = 0;
        legend.lineFormat.color = color;
        legend.fill.setSolidColor(color);
        legend.fill.transparency = transparencyRange[i];
        cell.values = [[textRange[i]]];
      }
      await context.sync();
    });
  }
  async function removeAll() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      var shapes = sheet.shapes;
      shapes.load("items/$none");
      return context.sync().then(function () {
        shapes.items.forEach(function (shape) {
          shape.delete();
        });
        return context.sync();
      });
    });
  }
  async function removeDistributions() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      var charts = sheet.charts;
      charts.load("items/$none");
      return context.sync().then(function () {
        charts.items.forEach(function (chart) {
          chart.delete();
        });
        return context.sync();
      });
    });
  }
  /** Default helper for invoking an action and handling errors. */
  async function tryCatch(callback) {
    try {
      await callback();
    } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
    }
  }
}