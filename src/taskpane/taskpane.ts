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
  document.getElementById("inputRelationship").onclick = showInputRelationship;
  document.getElementById("outputRelationship").onclick = showOutputRelationship;
  document.getElementById("first").onchange = first;
  document.getElementById("second").onchange = second;
  document.getElementById("third").onchange = third;
  document.getElementById("startWhatIf").onchange = startWhatIf;
  document.getElementById("useNewValues").onclick = useNewValues;
  document.getElementById("dismissValues").onclick = dismissValues;
}


async function parseSheet() {

  SheetProperties.isSheetParsed = true;

  try {
    hideOptions();
    console.log("Start parsing the sheet");

    SheetProperties.cellProp = new CellProperties();
    // eslint-disable-next-line require-atomic-updates
    SheetProperties.cells = await SheetProperties.cellProp.getCells();
    SheetProperties.cellProp.getRelationshipOfCells();

    console.log('Done parsing the sheet');
    showReferenceCellOption();
  } catch (error) {
    console.log(error);
    showReferenceCellOption();
  }
}

async function markAsReferenceCell() {
  try {

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

      console.log('Marking a reference cell');

      SheetProperties.referenceCell = SheetProperties.cellProp.getReferenceAndNeighbouringCells(range.address);
      SheetProperties.cellProp.checkUncertainty(SheetProperties.cells);
      SheetProperties.cellOp = new CellOperations(SheetProperties.cells, SheetProperties.referenceCell, 1);
      SheetProperties.isReferenceCell = true;

      console.log('Done Marking a reference cell');
      showVisualizationOption();
    });

  } catch (error) {
    console.error(error);
    showVisualizationOption();
  }
}

function showInputRelationship() {
  try {
    console.log('Degree of neighbourhood: ' + SheetProperties.degreeOfNeighbourhood);

    var element = <HTMLInputElement>document.getElementById("inputRelationship");

    if (element.checked) {
      showAllOptions();
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
      showAllOptions();
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

function first() {

  SheetProperties.degreeOfNeighbourhood = 1;
  console.log('First');
  removeShapesFromReferenceCell();
  displayOptions();
}


function second() {
  SheetProperties.degreeOfNeighbourhood = 2;
  console.log('Second');
  removeShapesFromReferenceCell();
  displayOptions();
}


function third() {
  SheetProperties.degreeOfNeighbourhood = 3;
  console.log('Third');
  removeShapesFromReferenceCell();
  displayOptions();
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
      SheetProperties.isSpread = true;
      SheetProperties.cellOp.showSpread(SheetProperties.degreeOfNeighbourhood);
    } else {
      // eslint-disable-next-line require-atomic-updates
      SheetProperties.isSpread = false;
      await SheetProperties.cellOp.removeSpread();
    }
  } catch (error) {
    console.error(error);
  }
}

async function startWhatIf() {
  (<HTMLInputElement>document.getElementById("startWhatIf")).disabled = true;
  document.getElementById('useNewValues').hidden = true;
  document.getElementById('dismissValues').hidden = true;
  await whatIfProcess();
  document.getElementById('useNewValues').hidden = false;
  document.getElementById('dismissValues').hidden = false;
}


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

async function whatIfProcess() {

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
  whatif.setNewCells(newCells, SheetProperties.referenceCell);

  console.log('Computing new spread');
  await whatif.drawChangedSpread(SheetProperties.referenceCell, SheetProperties.degreeOfNeighbourhood);

  // console.log('Calculating updated number');

  // await whatif.calculateUpdatedNumber();

  // if (!SheetProperties.referenceCell.whatIf) {
  //   console.log('Returning because what if is null');
  //   return;
  // }

  // const updatedValue = SheetProperties.referenceCell.whatIf.value;

  // if (updatedValue == 0) {
  //   console.log('No update in value');
  // } else {
  //   console.log("CHANGE: " + updatedValue);
  //   SheetProperties.cellOp.deleteUpdateshapes();
  //   // SheetProperties.cellOp.addTextBoxOnUpdate(updatedValue);
  // }

  // if (SheetProperties.isSpread) {
  //   console.log('Computing new spread');
  //   await whatif.drawChangedSpread(SheetProperties.referenceCell, SheetProperties.referenceCell.variance);
  // }
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

function hideOptions() {

  document.getElementById('referenceCell').hidden = true;
  document.getElementById('relationship').hidden = true;
  document.getElementById('neighborhood').hidden = true;
  document.getElementById('impact').hidden = true;
  document.getElementById('likelihood').hidden = true;
  document.getElementById('spread').hidden = true;
  document.getElementById('startWhatIf').hidden = true;
  document.getElementById('useNewValues').hidden = true;
  document.getElementById('dismissValues').hidden = true;
}

function showReferenceCellOption() {
  document.getElementById('referenceCell').hidden = false;
}

function showVisualizationOption() {

  document.getElementById('relationship').hidden = false;
  document.getElementById('neighborhood').hidden = false;
  document.getElementById('impact').hidden = false;
  document.getElementById('likelihood').hidden = false;
  document.getElementById('spread').hidden = false;
  document.getElementById('startWhatIf').hidden = false;
  (<HTMLInputElement>document.getElementById("neighborhood")).disabled = true;
  (<HTMLInputElement>document.getElementById("impact")).disabled = true;
  (<HTMLInputElement>document.getElementById("likelihood")).disabled = true;
  (<HTMLInputElement>document.getElementById("spread")).disabled = false;
  (<HTMLInputElement>document.getElementById("startWhatIf")).disabled = false;
}


function showAllOptions() {

  document.getElementById('relationship').hidden = false;
  document.getElementById('neighborhood').hidden = false;
  document.getElementById('impact').hidden = false;
  document.getElementById('likelihood').hidden = false;
  document.getElementById('spread').hidden = false;
  document.getElementById('startWhatIf').hidden = false;
  (<HTMLInputElement>document.getElementById("neighborhood")).disabled = false;
  (<HTMLInputElement>document.getElementById("impact")).disabled = false;
  (<HTMLInputElement>document.getElementById("likelihood")).disabled = false;
  (<HTMLInputElement>document.getElementById("spread")).disabled = false;
  (<HTMLInputElement>document.getElementById("startWhatIf")).disabled = false;
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

