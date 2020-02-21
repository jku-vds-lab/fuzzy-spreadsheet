import CellOperations from './celloperations';
import CellProperties from './cellproperties';
import SheetProperties from './sheetproperties';
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
  document.getElementById("inputRelationship").onclick = showInputRelationship;
  document.getElementById("outputRelationship").onclick = showOutputRelationship;
  document.getElementById("removeAll").onclick = removeAll;
  document.getElementById("first").onchange = first;
  document.getElementById("second").onchange = second;
  document.getElementById("third").onchange = third;
}




var cellOp = new CellOperations();
var cellProp = new CellProperties();
var cells: CellProperties[];
var focusCell: CellProperties;

async function markAsFocusCell() {
  try {


    let range: Excel.Range;
    Excel.run(async context => {

      range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "lightgrey";
      await context.sync();


    });

    disableInputs();

    cellOp = new CellOperations();
    cellProp = new CellProperties();
    console.log('Getting properties of cells');
    cells = await cellProp.getCellsProperties(); // needs to be optimised
    console.log('Getting relationship of cells');
    await cellProp.getRelationshipOfCells(cells); // already optimised
    console.log('Getting neighbouring cells');
    focusCell = cellProp.getFocusAndNeighbouringCells(cells, range.address);
    console.log('Checking Uncertain cells');
    cellProp.checkUncertainty(cells); // alreadt optimised
    cellOp.setCells(cells);
    SheetProperties.isFocusCell = true;
    console.log("Cells: ", cells);
    enableInputs();
  } catch (error) {
    console.error(error);
    enableInputs();
  }
}

function disableInputs() {

  document.getElementById('loading').hidden = false;
  (<HTMLInputElement>document.getElementById("impact")).disabled = true;
  (<HTMLInputElement>document.getElementById("likelihood")).disabled = true;
  (<HTMLInputElement>document.getElementById("spread")).disabled = true;
  (<HTMLInputElement>document.getElementById("inputRelationship")).disabled = true;
  (<HTMLInputElement>document.getElementById("outputRelationship")).disabled = true;
  (<HTMLInputElement>document.getElementById("removeAll")).disabled = true;
  (<HTMLInputElement>document.getElementById("first")).disabled = true;
  (<HTMLInputElement>document.getElementById("second")).disabled = true;
  (<HTMLInputElement>document.getElementById("third")).disabled = true;

}

function enableInputs() {
  document.getElementById('loading').hidden = true;
  (<HTMLInputElement>document.getElementById("impact")).disabled = false;
  (<HTMLInputElement>document.getElementById("likelihood")).disabled = false;
  (<HTMLInputElement>document.getElementById("spread")).disabled = false;
  (<HTMLInputElement>document.getElementById("inputRelationship")).disabled = false;
  (<HTMLInputElement>document.getElementById("outputRelationship")).disabled = false;
  (<HTMLInputElement>document.getElementById("removeAll")).disabled = false;
  (<HTMLInputElement>document.getElementById("first")).disabled = false;
  (<HTMLInputElement>document.getElementById("second")).disabled = false;
  (<HTMLInputElement>document.getElementById("third")).disabled = false;
}

async function impact() {
  try {
    var element = <HTMLInputElement>document.getElementById("impact");
    if (element.checked) {
      SheetProperties.isImpact = true;
      await cellOp.addImpact(focusCell);
    } else {
      removeImpact();
    }
  } catch (error) {
    console.error(error);
  }
}

function removeImpact() {
  try {
    //remove impact
    focusCell.inputCells.forEach((cell: CellProperties) => {
      // remove focus cells

    })
  } catch (error) {
    console.error(error);
  }
}

async function likelihood() {
  try {
    await cellOp.addLikelihood(focusCell);
    SheetProperties.isLikelihood = true;
  } catch (error) {
    console.error(error);
  }
}

async function spread() {
  try {
    await cellOp.addSpread(focusCell);
    SheetProperties.isSpread = true;
  } catch (error) {
    console.error(error);
  }
}

async function removeAll() {
  SheetProperties.isFocusCell = false;
  SheetProperties.isImpact = false;
  SheetProperties.isLikelihood = false;
  SheetProperties.isSpread = false;
  SheetProperties.isInputRelationship = false;
  SheetProperties.isOutputRelationship = false;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
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

function showInputRelationship() {
  try {
    blurBackground();
    cellOp.addInArrows(focusCell, focusCell.inputCells);
    SheetProperties.isInputRelationship = true;
  } catch (error) {
    console.error(error);
  }
}

function showOutputRelationship() {
  try {
    blurBackground();
    cellOp.addOutArrows(focusCell, focusCell.outputCells);
    SheetProperties.isOutputRelationship = true;
  } catch (error) {
    console.error(error);
  }
}

function blurBackground() {
  try {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange(true);
      range.format.font.color = "grey";

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
//     const sheet = context.workbook.worksheets.getActiveWorksheet();
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

function protectSheet() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("protection/protected");
    return context.sync().then(function () {
      if (!sheet.protection.protected) {
        console.log("Protecting the sheet");
        sheet.protection.protect();
      }
    });
  });
}
function unprotectSheet() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("protection/protected");
    return context.sync().then(function () {
      if (sheet.protection.protected) {
        console.log("Unprotecting the sheet");
        sheet.protection.unprotect();
      }
    });
  });
}
// async function removeLikelihood() {
//   // To be fixed
//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getActiveWorksheet();
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

var eventResult;

Excel.run(function (context) {
  var worksheet = context.workbook.worksheets.getActiveWorksheet();
  eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

  return context.sync()
    .then(function () {
      console.log(eventResult);
    });
}).catch(errorHandlerFunction);

function handleSelectionChange(event) {
  return Excel.run(function (context) {
    return context.sync()
      .then(function () {
        if (SheetProperties.isFocusCell) {
          cellOp.showPopUpWindow(event.address);
        }
        console.log("Address of current selection: ", event);
      });
  }).catch(errorHandlerFunction);
}

function remove() {
  return Excel.run(eventResult.context, function (context) {
    eventResult.remove();

    return context.sync()
      .then(function () {
        eventResult = null;
        console.log("Event handler successfully removed.");
      });
  }).catch(errorHandlerFunction);
}

function errorHandlerFunction(callback) {
  try {
    callback();
  } catch (error) {
    console.log(error);
  }
}

function first() {
  console.log('first');
}


function second() {
  console.log('second');
}


function third() {
  console.log('third');
}