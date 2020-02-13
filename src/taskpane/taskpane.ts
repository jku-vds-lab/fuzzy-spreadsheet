import CellOperations from './celloperations';
import CellProperties from './cellproperties';
// C:\Windows\SysWOW64\F12

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */
Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("focusCell").onclick = markAsFocusCell;
  document.getElementById("impact").onclick = impact;
  document.getElementById("likelihood").onclick = likelihood;
  document.getElementById("spread").onclick = spread;
  document.getElementById("relationship").onclick = showArrows;
  document.getElementById("removeAll").onclick = removeAll;
  document.getElementById("neighbour").onchange = callMe;
}

function callMe() {

  var e = document.getElementById("list") as HTMLSelectElement;
  console.log(e.options[e.selectedIndex].value);
}

var cellOp = new CellOperations();
var cellProp = new CellProperties();
var cells: CellProperties[];
var focusCell: CellProperties;

async function markAsFocusCell() {
  try {

    let range;

    Excel.run(async context => {

      range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";
      await context.sync();
    });

    cellOp = new CellOperations();
    cellProp = new CellProperties();
    cells = await cellProp.getCellsProperties();

    await cellProp.getRelationshipOfCells(cells);
    focusCell = cellProp.getNeighbouringCells(cells, range.address);
    cellOp.setCells(cells);
    cellProp.checkUncertainty(cells);
    // console.log("Cells: ", cells);
  } catch (error) {
    console.error(error);
  }
}

async function impact() {
  try {
    await cellOp.addImpact(focusCell);
  } catch (error) {
    console.error(error);
  }
}

async function likelihood() {
  try {
    await cellOp.addLikelihood();
  } catch (error) {
    console.error(error);
  }
}

async function spread() {
  try {

    cellOp.createNormalDistributions(cells);

    // cellOp.addSpread(focusCell);
  } catch (error) {
    console.error(error);
  }
}

async function removeAll() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Probability");
    const range = sheet.getUsedRange(true);
    range.format.font.color = "black";
    if (focusCell != null) {
      if (focusCell.address != null) {
        const cell = sheet.getRange(focusCell.address);
        cell.format.fill.clear();
      }
    }

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

function showArrows() {
  try {
    blurBackground();
    cellOp.addInArrows(focusCell, focusCell.inputCells);
    cellOp.addOutArrows(focusCell, focusCell.outputCells);
  } catch (error) {
    console.error(error);
  }
}

function blurBackground() {
  try {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      const range = sheet.getUsedRange(true);
      range.format.font.color = "lightgrey";

      let specialRange = sheet.getRange(focusCell.address);
      specialRange.format.font.color = "black";

      focusCell.inputCells.forEach((cell: CellProperties) => {
        specialRange = sheet.getRange(cell.address);
        specialRange.format.font.color = "black";
      })

      focusCell.outputCells.forEach((cell: CellProperties) => {
        specialRange = sheet.getRange(cell.address);
        specialRange.format.font.color = "black";
      })
      return context.sync();
    })
  } catch (error) {
    console.error(error);
  }
}

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