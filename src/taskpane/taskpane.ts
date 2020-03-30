import CellOperations from './celloperations';
import CellProperties from './cellproperties';
import SheetProperties from './sheetproperties';
import WhatIf from './operations/whatif';
import * as d3 from 'd3';
import * as jStat from 'jstat';
import { max, histogram, min } from 'd3';
import { range, dotMultiply, Matrix } from 'mathjs';
import { Bernoulli } from 'discrete-sampling';
import Likelihood from './operations/likelihood';
import Bins from './operations/bins';
import { add } from 'src/functions/functions';

// show cell info in the taskpane

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */
// discrete samples and continuous samples

Office.initialize = () => {

  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("parseSheet").onclick = parseSheet;
  document.getElementById("referenceCell").onclick = markAsReferenceCell;
  document.getElementById("impact").onclick = impact;
  document.getElementById("likelihood").onclick = likelihood;
  document.getElementById("spread").onclick = spread;
  document.getElementById("relationship").onclick = relationshipIcons;
  document.getElementById("inputRelationship").onclick = inputRelationship;
  document.getElementById("outputRelationship").onclick = outputRelationship;
  document.getElementById("first").onchange = first;
  document.getElementById("second").onchange = second;
  document.getElementById("third").onchange = third;
  document.getElementById("startWhatIf").onclick = startWhatIf;
  document.getElementById("useNewValues").onclick = useNewValues;
  document.getElementById("dismissValues").onclick = dismissValues;
}

async function protectSheet() {
  await Excel.run(async (context) => {
    let workbook = context.workbook;
    workbook.load("protection/protected");

    await context.sync();

    if (!workbook.protection.protected) {
      console.log('Sheet is protected');
      workbook.protection.protect();
    }
  });
}

async function unprotectSheet() {
  await Excel.run(async (context) => {
    let workbook = context.workbook;
    workbook.protection.unprotect();
    console.log('Sheet is unprotected');
  });
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

function shapeActivated() {
  try {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.shapes.load('items/length');

      await context.sync();
      const length = sheet.shapes.items.length;
      console.log('Length of shapes: ', length);
      var activationResult = sheet.shapes.getItemAt(0).onActivated.add(sheetActivated);
      return context.sync()
        .then(function () {
          console.log("Activation Handler registered");
        });
    })
  } catch (error) {
    console.log(error);
  }
}

async function sheetActivated() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const length = sheet.shapes.items.length;
    console.log('Length of shapes: ' + length);
    console.log('Shape is active');
    await context.sync();
  })
}

async function markAsReferenceCell() {
  try {

    if (SheetProperties.isReferenceCell) {
      removeShapesFromReferenceCell();
      // changeFontColorsToOriginal();
    }

    clearPreviousReferenceCell();

    let range: Excel.Range;
    Excel.run(async context => {

      range = context.workbook.getSelectedRange();
      range.load("address");
      drawBorder();
      await context.sync();

      console.log('Marking a reference cell');

      SheetProperties.referenceCell = SheetProperties.cellProp.getReferenceAndNeighbouringCells(range.address);
      SheetProperties.cellProp.checkUncertainty(SheetProperties.cells);
      SheetProperties.cellOp = new CellOperations(SheetProperties.cells, SheetProperties.referenceCell, 1);
      SheetProperties.isReferenceCell = true;
      console.log('Done Marking a reference cell');

      showVisualizationOption();
      displayOptions();
      selectSomethingElse();
    });

  } catch (error) {
    console.error(error);
    showVisualizationOption();
  }
}

function drawBorder(address: string = null, isSetToOriginal: boolean = false) {

  try {
    Excel.run(async context => {
      let range: Excel.Range;
      let color: string = 'orange'

      if (address == null) {
        range = context.workbook.getSelectedRange();
      } else {
        range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
      }


      if (isSetToOriginal) {
        range.format.borders.getItem('EdgeTop').color = SheetProperties.originalTopBorder.color;
        range.format.borders.getItem('EdgeBottom').color = SheetProperties.originalBottomBorder.color;
        range.format.borders.getItem("EdgeLeft").color = SheetProperties.originalLeftBorder.color;
        range.format.borders.getItem('EdgeRight').color = SheetProperties.originalRightBorder.color;

        range.format.borders.getItem('EdgeTop').weight = SheetProperties.originalTopBorder.weight;
        range.format.borders.getItem('EdgeBottom').weight = SheetProperties.originalBottomBorder.weight;
        range.format.borders.getItem("EdgeLeft").weight = SheetProperties.originalLeftBorder.weight;
        range.format.borders.getItem('EdgeRight').weight = SheetProperties.originalRightBorder.weight;

        range.format.borders.getItem('EdgeTop').style = SheetProperties.originalTopBorder.style;
        range.format.borders.getItem('EdgeBottom').style = SheetProperties.originalBottomBorder.style;
        range.format.borders.getItem("EdgeLeft").style = SheetProperties.originalLeftBorder.style;
        range.format.borders.getItem('EdgeRight').style = SheetProperties.originalRightBorder.style;
      }
      else {
        SheetProperties.originalTopBorder = range.format.borders.getItem('EdgeTop');
        SheetProperties.originalBottomBorder = range.format.borders.getItem('EdgeBottom');
        SheetProperties.originalLeftBorder = range.format.borders.getItem('EdgeLeft');
        SheetProperties.originalRightBorder = range.format.borders.getItem('EdgeRight');

        SheetProperties.originalTopBorder.load(['color', 'weight', 'style']);
        SheetProperties.originalBottomBorder.load(['color', 'weight', 'style']);
        SheetProperties.originalLeftBorder.load(['color', 'weight', 'style']);
        SheetProperties.originalRightBorder.load(['color', 'weight', 'style']);

        range.format.borders.getItem('EdgeTop').color = color;
        range.format.borders.getItem('EdgeBottom').color = color;
        range.format.borders.getItem("EdgeLeft").color = color;
        range.format.borders.getItem('EdgeRight').color = color;

        range.format.borders.getItem('EdgeTop').weight = "Thick";
        range.format.borders.getItem('EdgeBottom').weight = "Thick";
        range.format.borders.getItem('EdgeLeft').weight = "Thick";
        range.format.borders.getItem('EdgeRight').weight = "Thick";
      }

      return context.sync().then(() => { }).catch((reason: any) => console.log(reason));
    })
  } catch (error) {
    console.log(error);
  }
}


function inputRelationship() {
  try {

    var element = <HTMLInputElement>document.getElementById("inputRelationship");

    if (element.checked) {
      showAllOptions();
      SheetProperties.isInputRelationship = true;
      showInputRelationForOptions();
      checkCellChanged();
    } else {
      SheetProperties.isInputRelationship = false;
      removeInputRelationFromOptions();
    }

  } catch (error) {
    console.error(error);
  }
  selectSomethingElse();
}

function outputRelationship() {
  try {
    var element = <HTMLInputElement>document.getElementById("outputRelationship");

    if (element.checked) {
      showAllOptions();
      SheetProperties.isOutputRelationship = true;
      showOutputRelationForOptions();
      checkCellChanged();
    } else {
      SheetProperties.isOutputRelationship = false;
      removeOutputRelationFromOptions();
    }

  } catch (error) {
    console.error(error);
  }
  selectSomethingElse();
}

function first() {

  SheetProperties.degreeOfNeighbourhood = 1;
  // removeShapesFromReferenceCell();
  displayOptions();
  selectSomethingElse();
}


function second() {
  SheetProperties.degreeOfNeighbourhood = 2;
  // removeShapesFromReferenceCell();
  displayOptions();
  selectSomethingElse();
}


function third() {
  SheetProperties.degreeOfNeighbourhood = 3;
  // removeShapesFromReferenceCell();
  displayOptions();
  selectSomethingElse();
}

function impact() {
  try {
    var element = <HTMLInputElement>document.getElementById("impact");

    if (element.checked) {
      SheetProperties.isImpact = true;
      if (SheetProperties.isInputRelationship) {
        SheetProperties.cellOp.showInputImpact(SheetProperties.degreeOfNeighbourhood);
      }

      if (SheetProperties.isOutputRelationship) {
        SheetProperties.cellOp.showOutputImpact(SheetProperties.degreeOfNeighbourhood);
      }
      checkCellChanged();
    } else {
      removeImpactPercentage();
      SheetProperties.isImpact = false;
      SheetProperties.cellOp.removeInputImpact(SheetProperties.degreeOfNeighbourhood);
      SheetProperties.cellOp.removeOutputImpact(SheetProperties.degreeOfNeighbourhood);
    }
    selectSomethingElse();
  } catch (error) {
    console.error(error);
  }
}

function removeImpactPercentage() {
  document.getElementById('impactPercentage').innerHTML = '';
}

function likelihood() {
  try {
    var element = <HTMLInputElement>document.getElementById("likelihood");

    if (element.checked) {
      SheetProperties.isLikelihood = true;
      if (SheetProperties.isInputRelationship) {
        SheetProperties.cellOp.showInputLikelihood(SheetProperties.degreeOfNeighbourhood);
      }

      if (SheetProperties.isOutputRelationship) {
        SheetProperties.cellOp.showOutputLikelihood(SheetProperties.degreeOfNeighbourhood);
      }
      checkCellChanged();
    } else {
      SheetProperties.isLikelihood = false;
      SheetProperties.cellOp.removeInputLikelihood(SheetProperties.degreeOfNeighbourhood);
      SheetProperties.cellOp.removeOutputLikelihood(SheetProperties.degreeOfNeighbourhood);
    }
    selectSomethingElse();
  } catch (error) {
    console.error(error);
  }
}

async function spread() {
  try {
    var element = <HTMLInputElement>document.getElementById("spread");
    await unprotectSheet();
    if (element.checked) {
      SheetProperties.isSpread = true;
      SheetProperties.cellOp.showSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
      checkCellChanged();
    } else {
      SheetProperties.isSpread = false;
      removeHtmlSpreadInfoForOriginalChart();
      removeHtmlSpreadInfoForNewChart();
      SheetProperties.cellOp.removeSpread(SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship, true);
      SheetProperties.cellOp.removeSpreadFromReferenceCell();
    }
    selectSomethingElse();
    await protectSheet();
  } catch (error) {
    console.error(error);
  }
}

function relationshipIcons() {

  var element = <HTMLInputElement>document.getElementById("relationship");

  if (element.checked) {

    SheetProperties.isRelationship = true;

    if (SheetProperties.isInputRelationship) {
      console.log('Input Relation for: ' + SheetProperties.degreeOfNeighbourhood);
      SheetProperties.cellOp.showInputRelationship(SheetProperties.degreeOfNeighbourhood);
    }

    if (SheetProperties.isOutputRelationship) {
      SheetProperties.cellOp.showOutputRelationship(SheetProperties.degreeOfNeighbourhood);
    }
  } else {
    SheetProperties.isRelationship = false;
    SheetProperties.cellOp.removeInputRelationship();
    SheetProperties.cellOp.removeOutputRelationship();
  }
  selectSomethingElse();
}

async function startWhatIf() {
  try {
    (<HTMLInputElement>document.getElementById("startWhatIf")).disabled = true;
    document.getElementById('useNewValues').hidden = true;
    document.getElementById('dismissValues').hidden = true;
    performWhatIf();
    document.getElementById('useNewValues').hidden = false;
    document.getElementById('dismissValues').hidden = false;
    (<HTMLInputElement>document.getElementById("startWhatIf")).disabled = false;
  } catch (error) {
    console.log(error);
  }
}
var eventResult;

function performWhatIf() {
  Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    console.log('Worksheet has changed');
    eventResult = worksheet.onChanged.add(processWhatIf); // onCalculated

    return context.sync()
      .then(function () {
        console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
      });
  }).catch((reason: any) => { console.log(reason) });
}


async function processWhatIf() {

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


  // let x = await SheetProperties.cellProp.getCellsFormulasValues();
  // // eslint-disable-next-line require-atomic-updates
  // SheetProperties.newValues = x.values;
  // // eslint-disable-next-line require-atomic-updates
  // SheetProperties.newFormulas = x.formulas;
  // console.log('Original' + SheetProperties.cells[0].value);
  // console.log('New' + SheetProperties.newValues[0]);

  // eslint-disable-next-line require-atomic-updates
  SheetProperties.newCells = SheetProperties.cellProp.updateNewValues(SheetProperties.newValues, SheetProperties.newFormulas);

  const whatif = new WhatIf(SheetProperties.newCells, SheetProperties.cells, SheetProperties.referenceCell);

  whatif.calculateChange();

  SheetProperties.cellOp.deleteUpdateshapes();

  whatif.showUpdateTextInCells(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);

  if (SheetProperties.isSpread) {
    whatif.showNewSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
  }
}


// To be fixed!!
async function useNewValues() {
  try {
    document.getElementById('useNewValues').hidden = true;
    document.getElementById('dismissValues').hidden = true;
    removeHandler();
    removeHtmlSpreadInfoForOriginalChart();
    removeHtmlSpreadInfoForNewChart();
    removeAllShapes();
    SheetProperties.newCells = null;
    SheetProperties.cellProp = new CellProperties();
    // eslint-disable-next-line require-atomic-updates
    SheetProperties.cells = await SheetProperties.cellProp.getCells();
    SheetProperties.cellProp.getRelationshipOfCells();
    // eslint-disable-next-line require-atomic-updates
    SheetProperties.referenceCell = SheetProperties.cellProp.getReferenceAndNeighbouringCells(SheetProperties.referenceCell.address);
    SheetProperties.cellProp.checkUncertainty(SheetProperties.cells);
    // eslint-disable-next-line require-atomic-updates
    SheetProperties.cellOp = new CellOperations(SheetProperties.cells, SheetProperties.referenceCell, 1);
    // eslint-disable-next-line require-atomic-updates
    SheetProperties.isReferenceCell = true;
    displayOptions();
  } catch (error) {
    console.log(error);
  }
}

function removeAllShapes() {

  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    var shapes = sheet.shapes;
    shapes.load("items/$none");
    return context.sync().then(function () {
      shapes.items.forEach(function (shape) {
        shape.delete();
      });
      return context.sync();
    });
  });

  // function setCellPropsToFalse() {

  // }
}

async function dismissValues() {

  try {
    document.getElementById('useNewValues').hidden = true;
    document.getElementById('dismissValues').hidden = true;
    console.log('Remove Event Handler');

    removeHandler();

    if (SheetProperties.isSpread) {
      const whatif = new WhatIf(SheetProperties.newCells, SheetProperties.cells, SheetProperties.referenceCell);
      whatif.deleteNewSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
      removeHtmlSpreadInfoForNewChart();
    }

    SheetProperties.cellOp.deleteUpdateshapes();

    SheetProperties.newCells = null;

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let cellRanges = new Array<Excel.Range>();
      let cellValues = new Array<number>();
      let cellFormulas = new Array<any>();

      SheetProperties.cells.forEach((cell: CellProperties) => {

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

  } catch (error) {
    console.log('Error: ', error);
  }
}

function removeHandler() {
  return Excel.run(eventResult.context, function (context) {
    eventResult.remove();

    return context.sync()
      .then(function () {
        eventResult = null;
        console.log("Event handler successfully removed.");
      });
  }).catch((reason: any) => { console.log(reason) });
}

function displayOptions() {

  try {

    if (SheetProperties.isImpact && SheetProperties.isLikelihood) {
      SheetProperties.cellOp.addLikelihoodInfo();
      impact();
    } else if (SheetProperties.isImpact) {
      impact();
    } else if (SheetProperties.isLikelihood) {
      likelihood();
    }

    if (SheetProperties.isRelationship) {
      relationshipIcons();
    }

    if (SheetProperties.isSpread) {
      spread();
    }

  } catch (error) {
    console.log(error);
  }
  selectSomethingElse();
}

function showInputRelationForOptions() {


  if (SheetProperties.isImpact && SheetProperties.isLikelihood) {

    SheetProperties.cellOp.addLikelihoodInfo();
    SheetProperties.cellOp.showInputImpact(SheetProperties.degreeOfNeighbourhood);
  } else if (SheetProperties.isImpact) {

    SheetProperties.cellOp.showInputImpact(SheetProperties.degreeOfNeighbourhood);
  } else if (SheetProperties.isLikelihood) {

    SheetProperties.cellOp.showInputLikelihood(SheetProperties.degreeOfNeighbourhood);
  }

  if (SheetProperties.isRelationship) {
    SheetProperties.cellOp.showInputRelationship(SheetProperties.degreeOfNeighbourhood);
  }

  if (SheetProperties.isSpread) {
    SheetProperties.cellOp.showSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
  }

  selectSomethingElse();
}

function showOutputRelationForOptions() {

  if (SheetProperties.isImpact) {
    SheetProperties.cellOp.showOutputImpact(SheetProperties.degreeOfNeighbourhood);
  }
  if (SheetProperties.isLikelihood) {
    SheetProperties.cellOp.showOutputLikelihood(SheetProperties.degreeOfNeighbourhood);
  }
  if (SheetProperties.isSpread) {
    SheetProperties.cellOp.showSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
  }
  if (SheetProperties.isRelationship) {
    SheetProperties.cellOp.showOutputRelationship(SheetProperties.degreeOfNeighbourhood);
  }
  selectSomethingElse();
}

function removeInputRelationFromOptions() {

  if (SheetProperties.isImpact) {
    SheetProperties.cellOp.removeInputImpact(SheetProperties.degreeOfNeighbourhood);
  }
  if (SheetProperties.isLikelihood) {
    SheetProperties.cellOp.removeInputLikelihood(SheetProperties.degreeOfNeighbourhood);
  }
  if (SheetProperties.isSpread) {
    SheetProperties.cellOp.removeSpread(SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship, false);
  }
  if (SheetProperties.isRelationship) {
    SheetProperties.cellOp.removeInputRelationship();
  }
  selectSomethingElse();
}

function removeOutputRelationFromOptions() {

  if (SheetProperties.isImpact) {
    SheetProperties.cellOp.removeOutputImpact(SheetProperties.degreeOfNeighbourhood);
  }
  if (SheetProperties.isLikelihood) {
    SheetProperties.cellOp.removeOutputLikelihood(SheetProperties.degreeOfNeighbourhood);
  }
  if (SheetProperties.isSpread) {
    SheetProperties.cellOp.removeSpread(SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship, false);
  }
  if (SheetProperties.isRelationship) {
    SheetProperties.cellOp.removeOutputRelationship();
  }
  selectSomethingElse();
}

function hideOptions(isReferenceCellHidden: boolean = true) {

  if (isReferenceCellHidden) {
    document.getElementById('referenceCell').hidden = true;
  }
  document.getElementById('relationshipDiv').hidden = true;
  document.getElementById('neighborhoodDiv').hidden = true;
  document.getElementById('impactDiv').hidden = true;
  document.getElementById('likelihoodDiv').hidden = true;
  document.getElementById('spreadDiv').hidden = true;
  document.getElementById('relationshipInfoDiv').hidden = true;
  document.getElementById('startWhatIf').hidden = true;
  document.getElementById('useNewValues').hidden = true;
  document.getElementById('dismissValues').hidden = true;
}

function showReferenceCellOption() {
  document.getElementById('referenceCell').hidden = false;
}

function showVisualizationOption() {

  document.getElementById('relationshipDiv').hidden = false;
  document.getElementById('neighborhoodDiv').hidden = false;
  document.getElementById('impactDiv').hidden = false;
  drawImpactLegend(-200);
  document.getElementById('likelihoodDiv').hidden = false;
  document.getElementById('spreadDiv').hidden = false;
  document.getElementById('relationshipInfoDiv').hidden = false;
  document.getElementById('startWhatIf').hidden = false;
  (<HTMLInputElement>document.getElementById("neighborhoodDiv")).disabled = true;
  (<HTMLInputElement>document.getElementById("impactDiv")).disabled = true;
  (<HTMLInputElement>document.getElementById("likelihoodDiv")).disabled = true;
  (<HTMLInputElement>document.getElementById("relationshipInfoDiv")).disabled = true;
  (<HTMLInputElement>document.getElementById("spreadDiv")).disabled = false;
  (<HTMLInputElement>document.getElementById("startWhatIf")).disabled = false;
}


function showAllOptions() {

  document.getElementById('relationshipDiv').hidden = false;
  document.getElementById('neighborhoodDiv').hidden = false;
  document.getElementById('impactDiv').hidden = false;
  document.getElementById('likelihoodDiv').hidden = false;
  document.getElementById('spreadDiv').hidden = false;
  document.getElementById('relationshipInfoDiv').hidden = false;
  document.getElementById('startWhatIf').hidden = false;
  (<HTMLInputElement>document.getElementById("neighborhoodDiv")).disabled = false;
  (<HTMLInputElement>document.getElementById("impactDiv")).disabled = false;
  (<HTMLInputElement>document.getElementById("likelihoodDiv")).disabled = false;
  (<HTMLInputElement>document.getElementById("spreadDiv")).disabled = false;
  (<HTMLInputElement>document.getElementById("startWhatIf")).disabled = false;
  (<HTMLInputElement>document.getElementById("relationshipInfoDiv")).disabled = false;
}

function changeFontColorsToOriginal() {
  Excel.run((context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    SheetProperties.cells.forEach((cell: CellProperties) => {
      let range = sheet.getRange(cell.address);
      range.format.font.color = cell.fontColor;
    });
    return context.sync()
  })
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

function clearPreviousReferenceCell() {

  if (SheetProperties.referenceCell != null) {
    if (SheetProperties.referenceCell.address != null) {
      drawBorder(SheetProperties.referenceCell.address, true);
    }
  }
}

function checkCellChanged() {
  Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    worksheet.onSelectionChanged.add(handleSelectionChange);

    return context.sync()
      .then(function () {
        console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
      });
  }).catch((reason: any) => { console.log(reason) });
}

function handleSelectionChange(event) {
  return Excel.run(function (context) {
    return context.sync()
      .then(function () {

        if (SheetProperties.cells == null) {
          console.log('Returning because cells is undefined');
          return;
        }
        SheetProperties.cells.forEach((cell: CellProperties, index: number) => {

          if (cell.address.includes(event.address)) {

            removeImpactPercentage();
            clearRelationshipInfoInTaskpane();

            if (cell.isImpact) {
              addImpactPercentage(cell);
              drawImpactLegend(cell.impact, cell.rectColor);
            }

            if (cell.isLikelihood) {
              addLikelihoodPercentage(cell);
            }

            if (cell.isInputRelationship) {
              highlightInputRelationshipInfo(cell);
            }

            if (cell.isOutputRelationship) {
              highlightOutputRelationshipInfo(cell);
            }

            if (cell.isSpread) {
              showSpreadInTaskPane(cell);
              document.getElementById("mean").innerHTML = "Mean: " + cell.computedMean.toFixed(2) + " & Std Dev: " + cell.computedStdDev.toFixed(2);

              if (SheetProperties.newCells == null) {
                return;
              }
              if (cell.address == SheetProperties.newCells[index].address) {

                if (cell.samples == SheetProperties.newCells[index].samples) {
                  document.getElementById("newDistribution").hidden = true;
                  d3.select("#whatIfChart").select('svg').remove();
                  document.getElementById("spaceHack").hidden = true;
                  return;
                }
                document.getElementById("spaceHack").hidden = false;
                document.getElementById("newDistribution").hidden = false;
                document.getElementById("newMean").innerHTML = "New Mean: " + SheetProperties.newCells[index].computedMean.toFixed(2) + " & Std Dev: " + SheetProperties.newCells[index].computedStdDev.toFixed(2);
                showSpreadInTaskPane(SheetProperties.newCells[index], '.what-if-chart', 'whatIfChart', '#ff9933', true);
              }
            }
            else {
              removeHtmlSpreadInfoForOriginalChart();
              removeHtmlSpreadInfoForNewChart();
            }
          }
        })
      });
  }).catch((reason: any) => { console.log(reason) });
}

function highlightInputRelationshipInfo(cell: CellProperties) {

  clearRelationshipInfoInTaskpane();

  if (!cell.isInputRelationship) {
    return;
  }

  if (cell.degreeToFocus == 1) {
    document.getElementById('diamond1').className = 'dotted';
    document.getElementById('number1').className = 'dotted';
  }

  if (SheetProperties.degreeOfNeighbourhood == 2) {
    if (cell.degreeToFocus > 1) {
      document.getElementById('diamond2').className = 'dotted';
      document.getElementById('number2').className = 'dotted';
    }
  }

  if (SheetProperties.degreeOfNeighbourhood == 3) {

    console.log('Degree to focus: ' + cell.degreeToFocus);
    if (cell.degreeToFocus == 2) {
      document.getElementById('diamond2').className = 'dotted';
      document.getElementById('number2').className = 'dotted';
    } else if (cell.degreeToFocus > 2) {
      document.getElementById('diamond3').className = 'dotted';
      document.getElementById('number3').className = 'dotted';
    }
  }
}

function highlightOutputRelationshipInfo(cell: CellProperties) {

  clearRelationshipInfoInTaskpane();

  if (!cell.isOutputRelationship) {
    return;
  }

  if (cell.degreeToFocus == 1) {
    document.getElementById('circle1').className = 'dotted';
    document.getElementById('number1').className = 'dotted';
  }

  if (SheetProperties.degreeOfNeighbourhood == 2) {
    if (cell.degreeToFocus > 1) {
      document.getElementById('circle2').className = 'dotted';
      document.getElementById('number2').className = 'dotted';
    }
  }

  if (SheetProperties.degreeOfNeighbourhood == 3) {
    if (cell.degreeToFocus == 2) {
      document.getElementById('circle2').className = 'dotted';
      document.getElementById('number2').className = 'dotted';
    } else {
      document.getElementById('circle3').className = 'dotted';
      document.getElementById('number3').className = 'dotted';
    }
  }
}

function clearRelationshipInfoInTaskpane() {
  document.getElementById('number1').className = 'none';
  document.getElementById('number2').className = 'none';
  document.getElementById('number3').className = 'none';

  document.getElementById('diamond1').className = 'none';
  document.getElementById('diamond2').className = 'none';
  document.getElementById('diamond3').className = 'none';

  document.getElementById('circle1').className = 'none';
  document.getElementById('circle2').className = 'none';
  document.getElementById('circle3').className = 'none';
}


function drawImpactLegend(impact: number = 0, color: string = 'green') {


  d3.select("#impactLegend").select('svg').remove();
  impact = Math.ceil(impact * 0.5);

  if (color == 'green') {
    impact = impact + 50;
  }

  const minDomain = -5;
  const maxDomain = 40;
  const binWidth = 3;

  let binsObj = new Bins(minDomain, maxDomain, binWidth);
  var colors = binsObj.generateRedGreenColors();

  var Svg = d3.select('#impactLegend').append("svg")
    .attr("width", 200)
    .attr("height", 20);

  Svg.selectAll("mydots")
    .data(colors)
    .enter()
    .append("rect")
    .attr("x", function (d, i) { return (i) * 2 })
    .attr("y", function (d, i) {
      if (i == impact) {
        return 5;
      }
      return 10;
    })
    .attr("width", function (d, i) {
      if (i == impact) {
        return 8;
      }
      return 2;
    })
    .attr("height", function (d, i) {
      if (i == impact) {
        return 15;
      }
      return 5;
    }
    )
    .style("fill", (d) => { return d });
}


function removeHtmlSpreadInfoForOriginalChart() {
  try {
    d3.select("#" + 'originalChart').select('svg').remove();
    d3.select("#" + 'lines').select('svg').remove();
    d3.select("#" + 'spreadLegend').select('svg').remove();
    document.getElementById("mean").innerHTML = "";
  } catch (error) {
    console.log(error);
  }
}

function removeHtmlSpreadInfoForNewChart() {
  try {
    d3.select("#" + 'whatIfChart').select('svg').remove();
    d3.select("#" + 'newLines').select('svg').remove();
    d3.select("#" + 'newSpreadLegend').select('svg').remove();
    document.getElementById("newMean").innerHTML = "";
    document.getElementById("newDistribution").hidden = true;
    document.getElementById("spaceHack").hidden = true;
  } catch (error) {
    console.log(error);
  }
}

function showSpreadInTaskPane(cell: CellProperties, divClass: string = '.g-chart', idToBeRemoved: string = 'originalChart', color: string = '#399bfc', isLegendOrange: boolean = false) {

  try {

    d3.select("#" + idToBeRemoved).select('svg').remove();
    d3.select("#" + 'lines').select('svg').remove();
    d3.select("#" + 'spreadLegend').select('svg').remove();
    d3.select("#" + 'newLines').select('svg').remove();
    d3.select("#" + 'newSpreadLegend').select('svg').remove();

    if (SheetProperties.newCells == null) {
      d3.select('#whatIfChart').select('svg').remove();
    }

    let data = cell.samples;

    if (data == null) {
      return;
    }

    var margin = { top: 10, right: 30, bottom: 20, left: 40 },
      width = 260 - margin.left - margin.right,
      height = 150 - margin.top - margin.bottom;

    // append the svg object to the body of the page
    var svg = d3.select(divClass)
      .append("svg")
      .attr("width", width + margin.left + margin.right)
      .attr("height", height + margin.top + margin.bottom)
      .append("g")
      .attr("transform",
        "translate(" + margin.left + "," + margin.top + ")");


    const minDomain = -5;
    const maxDomain = 40;
    const binWidth = 3;

    let binsObj = new Bins(minDomain, maxDomain, binWidth);
    let bins = binsObj.createBins(data);
    let ticks = binsObj.getTickValues();

    var x = d3.scaleLinear().domain([minDomain, maxDomain]).range([0, width]);

    svg.append("g")
      .attr("transform", "translate(0," + height + ")")
      .call(d3.axisBottom(x).tickValues(ticks));

    var y = d3.scaleLinear()
      .range([height, 0])
      .domain([0, 100]);

    svg.append("g")
      .call(d3.axisLeft(y).ticks(5));

    svg.selectAll("rect")
      .data(bins)
      .enter()
      .append("rect")
      .attr("transform", function (d) { return "translate(" + x(d.x0) + "," + y(d.length) + ")"; })
      .attr("width", function (d) {
        if (x(d.x0) == x(d.x1)) {
          return 1;
        }
        return x(d.x1) - x(d.x0) - 1;
      })
      .attr("height", function (d) { return height - y(d.length); })
      .style("fill", color);

    drawLinesBeneathChart(cell);
    drawLegend();

    if (isLegendOrange) {
      drawLinesBeneathChart(cell, isLegendOrange);
      drawLegend(isLegendOrange);
    }

  } catch (error) {
    console.log(error);
  }
}

function drawLinesBeneathChart(cell: CellProperties, isLegendOrange: boolean = false) {

  var colors = cell.binBlueColors;
  let div = '#lines';

  if (isLegendOrange) {
    div = '#newLines';
    colors = cell.binOrangeColors;
  }

  var legendSvg = d3.select(div)
    .append("svg")
    .attr("width", 260)
    .attr("height", 10);

  // create a list of keys
  var keys = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14];

  // Add one dot in the legend for each name.
  legendSvg.selectAll("mydots")
    .data(keys)
    .enter()
    .append("rect")
    .attr("width", 3)
    .attr("x", function (d, i) { return 40 + i * 13 })
    .attr("y", 0)
    .attr("width", 12)
    .attr("height", 10)
    .style("fill", (d) => { return colors[d] });
}

function drawLegend(isLegendOrange: boolean = false) {

  const minDomain = -5;
  const maxDomain = 40;
  const binWidth = 3;

  let binsObj = new Bins(minDomain, maxDomain, binWidth);
  var colors = binsObj.generateBlueColors();


  let div = '#spreadLegend';

  if (isLegendOrange) {
    div = '#newSpreadLegend';
    colors = binsObj.generateOrangeColors();
  }

  var Svg = d3.select(div).append("svg")
    .attr("width", 125)
    .attr("height", 10);

  var keys = [0, 3, 6, 9, 12, 14];

  Svg.selectAll("mydots")
    .data(keys)
    .enter()
    .append("rect")
    .attr("x", function (d, i) { return (i + 2) * 11 })
    .attr("y", 0)
    .attr("width", 10)
    .attr("height", 5)
    .style("fill", (d) => { return colors[d] });

  Svg.selectAll("mylabels")
    .data([0, 100])
    .enter()
    .append("text")
    .attr("x", function (d, i) { return i * 90 })
    .attr("y", 10)
    .text(function (d) { return d + '%' })
    .attr("text-anchor", "left")
    .style("alignment-baseline", "middle")
    .style("font-size", "12px");
}

function selectSomethingElse() {
  Excel.run(function (context) {

    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange(SheetProperties.referenceCell.address);
    range.select();

    return context.sync();
  })
}

function addImpactPercentage(cell: CellProperties) {

  var impactText = document.getElementById('impactPercentage');
  impactText.innerHTML = cell.impact + '%';
  impactText.style.position = 'relative';
  impactText.style.left = 5 + 'px';

}

function addLikelihoodPercentage(cell: CellProperties) {

  var likelihoodText = document.getElementById('likelihoodPercentage');
  likelihoodText.innerHTML = (cell.likelihood * 100).toFixed(2) + '%';
  likelihoodText.style.position = 'relative';
  likelihoodText.style.left = 5 + 'px';

}