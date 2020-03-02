import CellOperations from './celloperations';
import CellProperties from './cellproperties';
import SheetProperties from './sheetproperties';
import WhatIf from './operations/whatif';
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
  document.getElementById("relationship").onclick = showRelationship;
  document.getElementById("inputRelationship").onclick = showInputRelationship;
  document.getElementById("outputRelationship").onclick = showOutputRelationship;
  document.getElementById("removeAll").onclick = removeAll;
  document.getElementById("first").onchange = first;
  document.getElementById("second").onchange = second;
  document.getElementById("third").onchange = third;
  document.getElementById("useNewValues").onclick = useNewValues;
  document.getElementById("dismissValues").onclick = dismissValues;
}

Excel.run(function (context) {

  var worksheet = context.workbook.worksheets.getActiveWorksheet();
  eventResult = worksheet.onChanged.add(handleDataChanged);

  return context.sync()
    .then(function () {
      console.log(eventResult);
      console.log('Got the range properties');

    });
}).catch(errorHandlerFunction);


function useNewValues() {
  SheetProperties.cellProp.updateNewValues(SheetProperties.newValues, SheetProperties.newFormulas, true);
}

async function dismissValues() {
  // Error so far
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    const values = new Array<any>();

    SheetProperties.cells.forEach((cell: CellProperties) => {
      values.push(cell.value);
    })

    range.values = [values];
    await context.sync();
  });
}

async function handleDataChanged() {


  console.log('Registered data changed');

  if (SheetProperties.referenceCell == null) {
    console.log('Returning because reference cell is null');
    return;
  }

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange(true);
    range.load(['formulas', 'values']);
    await context.sync();
    SheetProperties.newValues = range.values;
    SheetProperties.newFormulas = range.formulas;
  });

  let newCells = SheetProperties.cellProp.updateNewValues(SheetProperties.newValues, SheetProperties.newFormulas);

  newCells.forEach((nC: CellProperties) => {
    if (nC.id == SheetProperties.referenceCell.id) {
      console.log('New cell id:' + nC.value);
    }
  })

  console.log('Updated values');

  const whatif = new WhatIf();
  whatif.setNewCells(newCells);

  console.log('Calculating updated number');

  await whatif.calculateUpdatedNumber(SheetProperties.referenceCell);

  if (!SheetProperties.referenceCell.whatIf) {
    console.log('Returning because what if is null');
    return;
  }

  const updatedValue = SheetProperties.referenceCell.whatIf.value;

  if (updatedValue == 0) {
    console.log('No update in value');
  } else {
    console.log("CHANGE: " + updatedValue);
    SheetProperties.cellOp.deleteUpdateshapes();
    SheetProperties.cellOp.addTextBoxOnUpdate(updatedValue);
  }

  if (SheetProperties.isSpread) {
    console.log('Computing new spread');
    await whatif.drawChangedSpread(SheetProperties.referenceCell, SheetProperties.referenceCell.variance);
  }
}


async function parseSheet() {

  SheetProperties.isSheetParsed = true;

  try {
    disableInputs();
    console.log("Start parsing the sheet");

    SheetProperties.cellProp = new CellProperties();
    // eslint-disable-next-line require-atomic-updates
    SheetProperties.cells = await SheetProperties.cellProp.getCells();
    console.log('Cells', SheetProperties.cells);
    SheetProperties.cellProp.getRelationshipOfCells();

    console.log('Done parsing the sheet');
    enableInputs();
  } catch (error) {
    console.log(error);
    enableInputs();
  }
}

async function markAsReferenceCell() {
  try {

    if (!SheetProperties.isSheetParsed) {
      await parseSheet();
    }

    if (SheetProperties.isReferenceCell) {
      removeShapesFromReferenceCell();
    }

    clearPreviousReferenceCell();

    let range: Excel.Range;
    Excel.run(async context => {

      range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "lightgrey";
      await context.sync();

      disableInputs();
      console.log('Marking a reference cell');

      SheetProperties.referenceCell = SheetProperties.cellProp.getReferenceAndNeighbouringCells(range.address);
      SheetProperties.cellProp.checkUncertainty(SheetProperties.cells);
      SheetProperties.cellOp = new CellOperations(SheetProperties.cells, SheetProperties.referenceCell, 1);
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
  (<HTMLInputElement>document.getElementById("relationship")).disabled = true;
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
  (<HTMLInputElement>document.getElementById("relationship")).disabled = false;
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
      SheetProperties.cellOp.showImpact(SheetProperties.degreeOfNeighbourhood);
    } else {
      SheetProperties.isImpact = false;
      await SheetProperties.cellOp.removeImpact(SheetProperties.degreeOfNeighbourhood);
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
      SheetProperties.cellOp.showLikelihood(SheetProperties.degreeOfNeighbourhood); // should be available on click
    } else {
      SheetProperties.isLikelihood = false;
      await SheetProperties.cellOp.removeLikelihood(SheetProperties.degreeOfNeighbourhood);
    }
  } catch (error) {
    console.error(error);
  }
}

async function spread() {
  try {

    var element = <HTMLInputElement>document.getElementById("spread");

    if (element.checked) {
      // eslint-disable-next-line require-atomic-updates
      SheetProperties.isSpread = true;
      await SheetProperties.cellOp.createNewSheet();
      await SheetProperties.cellOp.showSpread(SheetProperties.degreeOfNeighbourhood);
    } else {
      // eslint-disable-next-line require-atomic-updates
      SheetProperties.isSpread = false;
      await SheetProperties.cellOp.removeSpread();
    }
  } catch (error) {
    console.error(error);
  }
}

async function removeShapesFromReferenceCell() {

  SheetProperties.cells.forEach((cell: CellProperties) => {
    cell.isImpact = false;
    cell.isInputRelationship = false;
    cell.isOutputRelationship = false;
    cell.isLikelihood = false;
    cell.isSpread = false;
  })

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    var shapes = sheet.shapes;
    var charts = sheet.charts;
    shapes.load("items/$none");
    charts.load("items/$none");
    return context.sync().then(function () {
      shapes.items.forEach(function (shape) {
        shape.delete();
      });
      charts.items.forEach(function (chart) {
        chart.delete();
      });

      return context.sync();
    });
  });
}

async function clearPreviousReferenceCell() {

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    if (SheetProperties.referenceCell != null) {
      if (SheetProperties.referenceCell.address != null) {
        const cell = sheet.getRange(SheetProperties.referenceCell.address);
        cell.format.fill.clear();
      }
    }
    await context.sync();
  })
}

async function removeAll() {

  clearPreviousReferenceCell();

  var element1 = <HTMLInputElement>document.getElementById("impact");
  var element2 = <HTMLInputElement>document.getElementById("likelihood");
  var element3 = <HTMLInputElement>document.getElementById("spread");
  var element4 = <HTMLInputElement>document.getElementById("inputRelationship");
  var element5 = <HTMLInputElement>document.getElementById("outputRelationship");
  var element6 = <HTMLInputElement>document.getElementById("relationship");

  element1.checked = false;
  element2.checked = false;
  element3.checked = false;
  element4.checked = false;
  element5.checked = false;
  element6.checked = false;
  await removeShapesFromReferenceCell();

}

function showRelationship() {

  var element = <HTMLInputElement>document.getElementById("relationship");
  var element1 = <HTMLInputElement>document.getElementById("inputRelationship");
  var element2 = <HTMLInputElement>document.getElementById("outputRelationship");

  if (element.checked) {

    console.log('Relationship');
    console.log('SheetProperties.isInputRelationship' + SheetProperties.isInputRelationship);
    console.log('SheetProperties.isOutputRelationship' + SheetProperties.isOutputRelationship);

    if (SheetProperties.isInputRelationship == false) {
      console.log('Input Relationship');
      element1.checked = true;
      showInputRelationship();
    }

    if (SheetProperties.isOutputRelationship == false) {
      console.log('Output Relationship');
      element2.checked = true;
      showOutputRelationship();
    }
  } else {
    console.log('Unchecked');
    if (SheetProperties.isInputRelationship == true) {
      element1.checked = false;
      showInputRelationship();
    }

    if (SheetProperties.isOutputRelationship == true) {
      element2.checked = false;
      showOutputRelationship();
    }
  }

}

function showInputRelationship() {
  try {
    var element = <HTMLInputElement>document.getElementById("inputRelationship");

    if (element.checked) {
      SheetProperties.isInputRelationship = true;
      SheetProperties.cellOp.showInputRelationship(SheetProperties.degreeOfNeighbourhood);
    } else {
      SheetProperties.isInputRelationship = false;
      SheetProperties.cellOp.removeInputRelationship();
    }
  } catch (error) {
    console.error(error);
  }
}

function showOutputRelationship() {
  try {
    var element = <HTMLInputElement>document.getElementById("outputRelationship");

    if (element.checked) {
      SheetProperties.isOutputRelationship = true;
      SheetProperties.cellOp.showOutputRelationship(SheetProperties.degreeOfNeighbourhood);
    } else {
      SheetProperties.isOutputRelationship = false;
      SheetProperties.cellOp.removeOutputRelationship();
    }
  } catch (error) {
    console.error(error);
  }
}

var eventResult;

Excel.run(function (context) {
  var worksheet = context.workbook.worksheets.getActiveWorksheet();
  eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

  return context.sync()
    .then(function () {
      console.log(eventResult);
    });
}).catch(errorHandlerFunction);

async function handleSelectionChange(event) {
  if (SheetProperties.isReferenceCell) {
    await SheetProperties.cellOp.showPopUpWindow(event.address);
  }
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
  SheetProperties.degreeOfNeighbourhood = 1;
  removeShapesFromReferenceCell();
  displayOptions();
}


function second() {
  SheetProperties.degreeOfNeighbourhood = 2;
  removeShapesFromReferenceCell();
  displayOptions();
}


function third() {
  SheetProperties.degreeOfNeighbourhood = 3;
  removeShapesFromReferenceCell();
  displayOptions();
}