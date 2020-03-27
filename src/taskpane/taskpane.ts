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
      changeFontColorsToOriginal();
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
      displayOptions();
      selectSomethingElse();
    });

  } catch (error) {
    console.error(error);
    showVisualizationOption();
  }
}

function inputRelationship() {
  try {

    var element = <HTMLInputElement>document.getElementById("inputRelationship");

    if (element.checked) {
      showAllOptions();
      SheetProperties.isInputRelationship = true;
      showInputRelationForOptions();
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
      SheetProperties.isImpact = false;
      SheetProperties.cellOp.removeInputImpact(SheetProperties.degreeOfNeighbourhood);
      SheetProperties.cellOp.removeOutputImpact(SheetProperties.degreeOfNeighbourhood);
    }
    selectSomethingElse();
  } catch (error) {
    console.error(error);
  }
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

    if (element.checked) {
      SheetProperties.isSpread = true;
      SheetProperties.cellOp.showSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
      checkCellChanged();
    } else {
      SheetProperties.isSpread = false;
      removeHtmlInfoForSpread();
      SheetProperties.cellOp.removeSpread(SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship, true);
      SheetProperties.cellOp.removeSpreadFromReferenceCell();
    }
    selectSomethingElse();
  } catch (error) {
    console.error(error);
  }
}

function relationshipIcons() {

  var element = <HTMLInputElement>document.getElementById("relationship");

  if (element.checked) {

    SheetProperties.isRelationship = true;

    if (SheetProperties.isInputRelationship) {
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

async function useNewValues() {

  console.log('Remove Event Handler');

  remove();
  SheetProperties.cellOp.deleteUpdateshapes();
  if (SheetProperties.isSpread) {
    const whatif = new WhatIf(SheetProperties.newCells, SheetProperties.cells, SheetProperties.referenceCell);
    whatif.deleteNewSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
  }

  await parseSheet();

}

async function dismissValues() {

  try {

    console.log('Remove Event Handler');

    remove();

    if (SheetProperties.isSpread) {
      const whatif = new WhatIf(SheetProperties.newCells, SheetProperties.cells, SheetProperties.referenceCell);
      whatif.deleteNewSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
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
    document.getElementById('useNewValues').hidden = true;
    document.getElementById('dismissValues').hidden = true;
  } catch (error) {
    console.log('Error: ', error);
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
  }).catch((reason: any) => { console.log(reason) });
}

function displayOptions() {

  let timeout = 1500;


  if (SheetProperties.isImpact && SheetProperties.isLikelihood) {
    SheetProperties.cellOp.addLikelihoodInfo();
    impact();
  } else if (SheetProperties.isImpact) {
    impact();
  } else if (SheetProperties.isLikelihood) {
    likelihood();
  }

  if (SheetProperties.isRelationship) {
    // eslint-disable-next-line no-undef
    setTimeout(() => {
      relationshipIcons();
    }, timeout);
  }

  if (SheetProperties.isSpread) {
    timeout = 3000;
    // eslint-disable-next-line no-undef
    setTimeout(() => {
      spread();
    }, timeout);
  }

  selectSomethingElse();
}

function showInputRelationForOptions() {


  let timeout = 0;

  if (SheetProperties.isImpact && SheetProperties.isLikelihood) {
    timeout = 1000;
    SheetProperties.cellOp.addLikelihoodInfo();
    SheetProperties.cellOp.showInputImpact(SheetProperties.degreeOfNeighbourhood);
  } else if (SheetProperties.isImpact) {
    timeout = 1000;
    SheetProperties.cellOp.showInputImpact(SheetProperties.degreeOfNeighbourhood);
  } else if (SheetProperties.isLikelihood) {
    timeout = 1000;
    SheetProperties.cellOp.showInputLikelihood(SheetProperties.degreeOfNeighbourhood);
  }

  if (SheetProperties.isRelationship) {
    // eslint-disable-next-line no-undef
    setTimeout(() => {
      SheetProperties.cellOp.showInputRelationship(SheetProperties.degreeOfNeighbourhood);
    }, timeout);

    if (timeout == 0) {
      timeout = 1000;
    } else {
      timeout = 2 * timeout;
    }
  }

  if (SheetProperties.isSpread) {
    // eslint-disable-next-line no-undef
    setTimeout(() => {
      SheetProperties.cellOp.showSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
    }, timeout);
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

            if (cell.isImpact) {
              addImpactPercentage(cell);
            }

            if (cell.isLikelihood) {
              addLikelihoodPercentage(cell);
            }

            if (cell.isSpread) {
              showSpreadInTaskPane(cell);
              document.getElementById("mean").innerHTML = "Mean: " + cell.computedMean.toFixed(2);
              document.getElementById("stdDev").innerHTML = "Std Dev: " + cell.computedStdDev.toFixed(2);

              if (SheetProperties.newCells == null) {
                return;
              }
              if (cell.address == SheetProperties.newCells[index].address) {

                if (cell.samples == SheetProperties.newCells[index].samples) {
                  document.getElementById("newDistribution").hidden = true;
                  d3.select("#whatIfChart").select('svg').remove();
                  return;
                }

                document.getElementById("newDistribution").hidden = false;
                document.getElementById("newMean").innerHTML = "New Mean: " + SheetProperties.newCells[index].computedMean.toFixed(2);
                document.getElementById("newStdDev").innerHTML = "New Std Dev: " + SheetProperties.newCells[index].computedStdDev.toFixed(2);
                showSpreadInTaskPane(SheetProperties.newCells[index], '.what-if-chart', 'whatIfChart', '#ff9933', true);
              }
            }
          }
        })
      });
  }).catch((reason: any) => { console.log(reason) });
}

function removeHtmlInfoForSpread() {
  try {
    d3.select("#" + 'originalChart').select('svg').remove();
    d3.select("#" + 'whatIfChart').select('svg').remove();
    d3.select("#" + 'lines').select('svg').remove();
    d3.select("#" + 'spreadLegend').select('svg').remove();
    d3.select("#" + 'newLines').select('svg').remove();
    d3.select("#" + 'newSpreadLegend').select('svg').remove();
    document.getElementById("newMean").innerHTML = "";
    document.getElementById("newStdDev").innerHTML = "";
    document.getElementById("newDistribution").hidden = true;
    document.getElementById("mean").innerHTML = "";
    document.getElementById("stdDev").innerHTML = "";
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
      width = 360 - margin.left - margin.right,
      height = 200 - margin.top - margin.bottom;

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
      .call(d3.axisLeft(y));

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
    .attr("width", 360)
    .attr("height", 30);

  // create a list of keys
  var keys = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14];

  // Add one dot in the legend for each name.
  legendSvg.selectAll("mydots")
    .data(keys)
    .enter()
    .append("rect")
    .attr("x", function (d, i) { return (i + 1) * 24 })
    .attr("y", 20) // 100 is where the first dot appears. 25 is the distance between dots
    .attr("width", 20)
    .attr("height", 20)
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
    .attr("width", 360)
    .attr("height", 30);

  var keys = [0, 3, 6, 9, 12, 14];

  Svg.selectAll("mydots")
    .data(keys)
    .enter()
    .append("rect")
    .attr("x", function (d, i) { return (i + 1) * 24 })
    .attr("y", 20)
    .attr("width", 20)
    .attr("height", 20)
    .style("fill", (d) => { return colors[d] });

  Svg.selectAll("mylabels")
    .data([0, 100])
    .enter()
    .append("text")
    .attr("x", function (d, i) { return i * 165 })
    .attr("y", 30)
    .text(function (d) { return d + '%' })
    .attr("text-anchor", "left")
    .style("alignment-baseline", "middle");
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