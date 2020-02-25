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
  document.getElementById("parseSheet").onclick = parseSheet;
  document.getElementById("referenceCell").onclick = markAsReferenceCell;
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

var cellOp: CellOperations;
var cellProp = new CellProperties();
var cells: CellProperties[];
var referenceCell: CellProperties;
var isSheetParsed = false;

Excel.run(function (context) {
  var worksheet = context.workbook.worksheets.getActiveWorksheet();
  eventResult = worksheet.onChanged.add(parseSheet); // improve logic for sheet parsing

  return context.sync()
    .then(function () {
      console.log(eventResult);
    });
}).catch(errorHandlerFunction);


async function parseSheet() {

  isSheetParsed = true;

  try {
    disableInputs();
    console.log("Start parsing the sheet");

    cellProp = new CellProperties();
    cells = await cellProp.getCellsProperties(); // needs to be optimised
    cellProp.getRelationshipOfCells(cells);

    console.log('Done parsing the sheet');
    enableInputs();
  } catch (error) {
    console.log(error);
    enableInputs();
  }
}

async function markAsReferenceCell() {
  try {

    if (!isSheetParsed) {
      await parseSheet();
    }

    if (SheetProperties.isReferenceCell) {
      removeShapesFromReferenceCell();
    }

    let range: Excel.Range;
    Excel.run(async context => {

      range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "lightgrey";
      await context.sync();

      disableInputs();
      console.log('Marking a reference cell');

      referenceCell = cellProp.getReferenceAndNeighbouringCells(cells, range.address);
      cellProp.checkUncertainty(cells);
      cellOp = new CellOperations(cells, referenceCell, 1);
      SheetProperties.isReferenceCell = true;

      console.log('Done Marking a reference cell');
      enableInputs();
      displayOptions();
    });

  } catch (error) {
    console.error(error);
    enableInputs();
  }
}

function displayOptions() {
  if (SheetProperties.isImpact) {
    impact();
  }
  if (SheetProperties.isLikelihood) {
    likelihood();
  }
  if (SheetProperties.isSpread) {
    spread();
  }
  if (SheetProperties.isInputRelationship) {
    showInputRelationship();
  }
  if (SheetProperties.isOutputRelationship) {
    showOutputRelationship();
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
      cellOp.showImpact();
    } else {
      SheetProperties.isImpact = false;
      await cellOp.removeImpact();
    }
  } catch (error) {
    console.error(error);
  }
}


async function likelihood() {
  try {
    var element = <HTMLInputElement>document.getElementById("likelihood");

    if (element.checked) {
      SheetProperties.isLikelihood = true;
      cellOp.showLikelihood();
    } else {
      SheetProperties.isLikelihood = false;
      await cellOp.removeLikelihood();
    }
  } catch (error) {
    console.error(error);
  }
}

async function spread() {
  try {

    if (!SheetProperties.isCheatSheetExist) {
      await cellOp.createCheatSheet(); // but create it just once
    }

    var element = <HTMLInputElement>document.getElementById("spread");

    if (element.checked) {
      // eslint-disable-next-line require-atomic-updates
      SheetProperties.isSpread = true;
      cellOp.showSpread();
    } else {
      // eslint-disable-next-line require-atomic-updates
      SheetProperties.isSpread = false;
      await cellOp.removeSpread();
    }
  } catch (error) {
    console.error(error);
  }
}

async function removeShapesFromReferenceCell() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange(true);
    range.format.font.color = "black";
    if (referenceCell != null) {
      if (referenceCell.address != null) {
        const cell = sheet.getRange(referenceCell.address);
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

async function removeAll() {

  var element1 = <HTMLInputElement>document.getElementById("impact");
  var element2 = <HTMLInputElement>document.getElementById("likelihood");
  var element3 = <HTMLInputElement>document.getElementById("spread");
  var element4 = <HTMLInputElement>document.getElementById("inputRelationship");
  var element5 = <HTMLInputElement>document.getElementById("outputRelationship");

  element1.checked = false;
  element2.checked = false;
  element3.checked = false;
  element4.checked = false;
  element5.checked = false;
  await removeShapesFromReferenceCell();

}

function showInputRelationship() {
  try {
    var element = <HTMLInputElement>document.getElementById("inputRelationship");

    if (element.checked) {
      blurBackground();
      SheetProperties.isInputRelationship = true;
      cellOp.showInputRelationship();
    } else {
      SheetProperties.isInputRelationship = false;
      unblurBackground();
      cellOp.removeInputRelationship();
    }
  } catch (error) {
    console.error(error);
  }
}

function showOutputRelationship() {
  try {
    var element = <HTMLInputElement>document.getElementById("outputRelationship");

    if (element.checked) {
      blurBackground();
      SheetProperties.isOutputRelationship = true;
      cellOp.showOutputRelationship();
    } else {
      SheetProperties.isOutputRelationship = false;
      unblurBackground();
      cellOp.removeOutputRelationship();
    }
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

      let specialRange = sheet.getRange(referenceCell.address);
      specialRange.format.font.color = "black";

      referenceCell.inputCells.forEach((cell: CellProperties) => {
        specialRange = sheet.getRange(cell.address);
        specialRange.format.font.color = "black";
      })

      referenceCell.outputCells.forEach((cell: CellProperties) => {
        specialRange = sheet.getRange(cell.address);
        specialRange.format.font.color = "black";
      })
      return context.sync();
    })
  } catch (error) {
    console.error(error);
  }
}

function unblurBackground() {

  Excel.run(function (context) {

    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange(true);
    range.format.font.color = "black";

    return context.sync();
  })
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
        if (SheetProperties.isReferenceCell) {
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