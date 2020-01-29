import CellOperations from './operations';
// C:\Windows\SysWOW64\F12

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */
Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("impact").onclick = run;
}

// async function addImpact() {
//   try {
//     // Ensure cells and shapes are the same length
//     await Excel.run(async (context) => {
//       let dim = new CellOperations();
//       let cellAddresses = ["I6", "I7", "I8", "I9", "I11", "I12", "I13", "I14",
//         "I15", "I16"];
//       await dim.scanRange(cellAddresses, "I18");
//       let cells = dim.getCells();
//       dim.addImpactInfo(cells);
//       let shapes = dim.getShapes();
//       const sheet = context.workbook.worksheets.getItem("Probability");
//       for (let i = 0; i < cells.length; i++) {
//         var impact = sheet.shapes.addGeometricShape("Rectangle"); // shapes[i].shapeType
//         impact.name = "Impact" + i;
//         impact.height = shapes[i].height;
//         impact.width = shapes[i].width;
//         impact.left = cells[i].left + 2;
//         impact.top = cells[i].top + cells[i].height / 4;
//         impact.rotation = 0;
//         impact.fill.transparency = shapes[i].transparency;
//         impact.lineFormat.weight = 0;
//         impact.lineFormat.color = shapes[i].color;
//         impact.fill.setSolidColor(shapes[i].color);
//       }
//       // createImpactLegend().then(function () { });
//       await context.sync();
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }
async function run() {
  try {
    let x = new CellOperations();
    let cells = await x.scanCellsInRange();
    await x.insertRelations(cells);
    console.log("Cells");
    console.log(cells);
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

// letlikehoodAddresses = ["J6", "J7", "J8", "J9", "J11", "J12", "J13", "J14",
// "J15", "J16"];

// async function addRelationship() {
//   let cells = dim.getCells();
//   dim.addRelationshipInfo(cells);
// }
// async function addSpread(isLine = false) {
//   let cells = dim.getCells();
//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getItem("Probability");
//     const cheatSheet = context.workbook.worksheets.getItem("CheatSheet");
//     // make it dynamic
//     let ranges: string[] = [
//       "A1:A47",
//       "B1:B47",
//       "C1:C47",
//       "D1:D47",
//       "E1:E47",
//       "F1:F47",
//       "G1:G47",
//       "H1:H47",
//       "I1:I47",
//       "J1:J47"
//     ];
//     for (let i = 0; i < ranges.length; i++) {
//       const dataRange = cheatSheet.getRange(ranges[i]);
//       let chart: Excel.Chart;
//       if (isLine) {
//         chart = sheet.charts.add("Line", dataRange, Excel.ChartSeriesBy.auto);
//       } else {
//         chart = sheet.charts.add("ColumnClustered", dataRange, Excel.ChartSeriesBy.auto);
//       }
//       chart.setPosition(cells[i].cell, cells[i].cell);
//       chart.left = cells[i].left + 0.2 * cells[i].width;
//       chart.title.visible = false;
//       chart.legend.visible = false;
//       chart.axes.valueAxis.minimum = 0;
//       chart.axes.valueAxis.maximum = 0.21;
//       chart.dataLabels.showValue = false;
//       chart.axes.valueAxis.visible = false;
//       chart.axes.categoryAxis.visible = false;
//       chart.axes.valueAxis.majorGridlines.visible = false;
//       chart.plotArea.top = 0;
//       chart.plotArea.left = 0;
//       chart.plotArea.width = cells[i].width - 0.4 * cells[i].width;
//       chart.plotArea.height = 100;
//       chart.format.fill.clear();
//       chart.format.border.clear();
//     }
//     await context.sync();
//   });
// }
// // Not possible without impact yet
// async function addLikelihood(isImpact: boolean = true) {
//   let shapes = dim.getShapes();
//   await dim.addLikelihoodInfo(cells, shapes, likehoodAddresses);
//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getItem("Probability");
//     for (let i = 0; i < shapes.length; i++) {
//       var shape = sheet.shapes.getItem("Impact" + i);
//       console.log(shape);
//       shape.load(["geometricShapeType", "width", "height"]);
//       await context.sync();
//       console.log("Geometric Shape Type: " + shape.geometricShapeType);
//       if (shape.geometricShapeType == Excel.GeometricShapeType.rectangle) {
//         console.log("Rectangle Found");
//         shape.height = shapes[i].height;
//         shape.width = shapes[i].width;
//       }
//     }
//     createLikelihoodLegend().then(function () { });
//     await context.sync();
//   });
// }

// async function protectSheet() {
//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getItem("Probability");
//     sheet.load("protection/protected");
//     await context.sync().then(function () {
//       if (!sheet.protection.protected) {
//         sheet.protection.protect();
//       }
//     });
//   });
// }
// async function unprotectSheet() {
//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getItem("Probability");
//     sheet.load("protection/protected");
//     await context.sync().then(function () {
//       if (sheet.protection.protected) {
//         sheet.protection.unprotect();
//       }
//     });
//   });
// }
// async function removeLikelihood() {
//   // To be fixed
//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getItem("Probability");
//     const count = sheet.shapes.getCount();
//     await context.sync();
//     for (let i = 0; i < count.value; i++) {
//       var shape = sheet.shapes.getItemAt(i);
//       shape.load(["geometricShapeType", "width", "height"]);
//       await context.sync();
//       if (shape.geometricShapeType == Excel.GeometricShapeType.rectangle) {
//         shape.width = 7;
//         shape.height = 7;
//       }
//     }
//     await context.sync();
//   });
// }
// async function createLikelihoodLegend() {
//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getItem("Probability");
//     const textRange = ["    < 50", "    <= 80", "    <= 100"];
//     const sizeRange = [5, 7, 9];
//     let color = "gray";
//     for (let i = 0; i < 3; i++) {
//       var legend = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
//       var cell = sheet.getCell(i + 22, 4);
//       cell.load("top");
//       cell.load("left");
//       cell.load("height");
//       cell.load("values");
//       await context.sync();
//       legend.height = sizeRange[i];
//       legend.width = sizeRange[i];
//       legend.left = cell.left + 2;
//       legend.top = cell.top + cell.height / 4;
//       legend.lineFormat.weight = 0;
//       legend.lineFormat.color = color;
//       legend.fill.setSolidColor(color);
//       cell.values = [[textRange[i]]];
//     }
//     await context.sync();
//   });
// }
// async function createImpactLegend() {
//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getItem("Probability");
//     const textRange = ["    > 20", "    >= 9 & < 20", "    < 9", "    < 9", "    >= 9 & < 20", "    > 20"];
//     const transparencyRange = [0, 0.4, 0.7, 0.7, 0.4, 0];
//     let color = "green";
//     for (let i = 0; i < 6; i++) {
//       if (i == 3) {
//         color = "red";
//       }
//       var legend = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
//       var cell = sheet.getCell(i + 22, 2);

//       cell.load("top");
//       cell.load("left");
//       cell.load("height");
//       cell.load("values");
//       await context.sync();
//       legend.height = 7;
//       legend.width = 7;
//       legend.left = cell.left + 2;
//       legend.top = cell.top + cell.height / 4;
//       legend.lineFormat.weight = 0;
//       legend.lineFormat.color = color;
//       legend.fill.setSolidColor(color);
//       legend.fill.transparency = transparencyRange[i];
//       cell.values = [[textRange[i]]];
//     }
//     await context.sync();
//   });
// }
// async function removeAll() {
//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getItem("Probability");
//     var shapes = sheet.shapes;
//     shapes.load("items/$none");
//     return context.sync().then(function () {
//       shapes.items.forEach(function (shape) {
//         shape.delete();
//       });
//       return context.sync();
//     });
//   });
// }
// async function removeDistributions() {
//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getItem("Probability");
//     var charts = sheet.charts;
//     charts.load("items/$none");
//     return context.sync().then(function () {
//       charts.items.forEach(function (chart) {
//         chart.delete();
//       });
//       return context.sync();
//     });
//   });
// }
