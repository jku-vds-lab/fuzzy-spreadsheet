import CellOperations from './celloperations';
import CellProperties from './cellproperties';
import SheetProperties from './sheetproperties';
import WhatIf from './operations/whatif';
import * as d3 from 'd3';

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
  document.getElementById("relationship").onclick = relationshipIcons;
  document.getElementById("inputRelationship").onclick = inputRelationship;
  document.getElementById("outputRelationship").onclick = outputRelationship;
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
      SheetProperties.cellOp.removeSpread(SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship, true);
      SheetProperties.cellOp.removeSpreadFromReferenceCell();
    }
    selectSomethingElse();
  } catch (error) {
    console.error(error);
  }
}

function showSpreadInTaskPane(cell: CellProperties) {

  try {

    // d3.selectAll("svg > *").remove();
    d3.select("svg").remove();

    // delete old data
    let data = cell.samples;
    data.forEach(function (d) {
      d.likelihood = +d.likelihood;
    });

    const margin = { top: 0, right: 0, bottom: 30, left: 0 };

    const width = 100 - margin.left - margin.right;
    const height = 125; // 125 - margin.top - margin.bottom;

    //Create the xScale
    const xScale = d3.scaleTime()
      .range([0, width]);

    //Create the yScale
    const yScale = d3.scaleLinear()
      .range([height, 0]);

    const svg = d3.select(".g-chart").append("svg")
      .attr("width", width + margin.left + margin.right)
      .attr("height", height)
      .append("g")
      .attr("transform", "translate(" + margin.left + "," + margin.top + ")");

    const div = d3.select(".g-chart").append("div")
      .attr("class", "tooltip")
      .style("opacity", 0);

    //Organizes the data
    d3.max(data, function (d) { return d.value; });

    //Defines the xScale max
    xScale.domain(d3.extent(data, function (d) { return d.value; }));

    //Defines the yScale max
    yScale.domain([0, 100]);

    svg.append("g")
      .attr("class", "x axis")
      .attr("transform", "translate(0," + height + ")")

    svg.selectAll("line.percent")
      .data(data)
      .enter()
      .append("line")
      .attr("class", "percentline")
      .attr("x1", (d) => { return xScale(d.value); })
      .attr("x2", (d) => { return xScale(d.value); })
      .attr("y1", 50)
      .attr("y2", 100)
      .style("stroke", "#cc0000")
      .style("stroke-width", 2)
      .style("opacity", (d) => { return d.likelihood })
      .on("mouseover", (d) => {

        try {
          var right = true;
          d3.select(this)
            .transition().duration(100)
            .attr("y1", 0)
            .style("stroke-width", 3)
            .style("opacity", 1);

          div.transition()
            .style("opacity", 1)
          div.html("<span class='bolded'>" + d.value + ": </span>" + d.likelihood * 100 + "%")

          let offset = right ? div.node().offsetWidth + 5 : -5;

          div
            .style("left", (d3.event.pageX - offset) + "px")
            .style("top", 425 + "px")
        } catch (error) {
          console.log(error);
        }

      })
      .on("mouseout", () => {
        d3.select(this)
          .transition().duration(100)
          .attr("y1", 50)
          .style("stroke-width", 2)
          .style("opacity", 0.4);

        div.transition()
          .style("opacity", 0)
      })
  } catch (error) {
    console.log(error);
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
  await whatif.drawChangedSpread(SheetProperties.referenceCell, SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);

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
    // SheetProperties.cellOp.removeAllImpacts();
    impact();
  }
  if (SheetProperties.isLikelihood) {
    // SheetProperties.cellOp.removeAllLikelihoods();
    likelihood();
  }
  if (SheetProperties.isSpread) {
    spread();
  }
  if (SheetProperties.isRelationship) {
    relationshipIcons();
  }
  selectSomethingElse();
}

function showInputRelationForOptions() {

  if (SheetProperties.isImpact) {
    SheetProperties.cellOp.showInputImpact(SheetProperties.degreeOfNeighbourhood);
  }
  if (SheetProperties.isLikelihood) {
    SheetProperties.cellOp.showInputLikelihood(SheetProperties.degreeOfNeighbourhood);
  }
  if (SheetProperties.isSpread) {
    SheetProperties.cellOp.showSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
  }
  if (SheetProperties.isRelationship) {
    SheetProperties.cellOp.showInputRelationship(SheetProperties.degreeOfNeighbourhood);
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

function hideOptions() {

  document.getElementById('referenceCell').hidden = true;
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
  document.getElementById('relationshipInfoDiv').hidden = true;
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
    var eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

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
        console.log("Address of current selection: " + event.address);

        if (SheetProperties.cells == null) {
          console.log('Returning because cells is undefined');
          return;
        }
        SheetProperties.cells.forEach((cell: CellProperties) => {
          if (cell.address.includes(event.address)) {

            console.log('Found a matching cell');
            if (cell.isSpread) {
              showSpreadInTaskPane(cell);
            }
          }
        })


      });
  }).catch((reason: any) => { console.log(reason) });
}

function selectSomethingElse() {
  Excel.run(function (context) {

    var sheet = context.workbook.worksheets.getActiveWorksheet();

    var range = sheet.getRange(SheetProperties.referenceCell.address);

    range.select();
    console.log('Select something else');

    return context.sync();
  })
}