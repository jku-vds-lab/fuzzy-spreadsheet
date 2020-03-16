import CellOperations from './celloperations';
import CellProperties from './cellproperties';
import SheetProperties from './sheetproperties';
import WhatIf from './operations/whatif';
import * as d3 from 'd3';
import * as jStat from 'jstat';
import { max, histogram } from 'd3';
import { range, dotMultiply, Matrix } from 'mathjs';
import { Bernoulli } from 'discrete-sampling';

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */
// discrete samples and continuous samples
Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("parseSheet").onclick = testjStatDistribution; //parseSheet;
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


function testjStatDistribution() {

  try {

    let normalSamples = new Array<number>();

    let values = range(0, 1, 0.01).toArray();

    values.forEach((el) => {
      normalSamples.push(jStat.normal.inv(el, 12, 1));
    })

    let sampleLength = normalSamples.length;

    var bern = Bernoulli(0.5);
    bern.draw();
    let bernoulliSamples = bern.sample(sampleLength);

    let finalLikelihood: any = dotMultiply(normalSamples, bernoulliSamples);
    console.log(finalLikelihood.length);
    drawHistogram(finalLikelihood);

  } catch (error) {
    console.log(error);
  }
}

function drawHistogram(data: number[]) {

  var margin = { top: 10, right: 30, bottom: 30, left: 40 },
    width = 460 - margin.left - margin.right,
    height = 400 - margin.top - margin.bottom;

  // append the svg object to the body of the page
  var svg = d3.select("body")
    .append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
    .append("g")
    .attr("transform",
      "translate(" + margin.left + "," + margin.top + ")");

  // get the data
  // X axis: scale and draw:

  let domain = d3.max(data, function (d) { return +d })
  var x = d3.scaleLinear()
    .domain([0, domain])     // can use this instead of 1000 to have the max of data: d3.max(data, function(d) { return +d.price })
    .range([0, width]);

  svg.append("g")
    .attr("transform", "translate(0," + height + ")")
    .call(d3.axisBottom(x));

  // set the parameters for the histogram
  var histogram = d3.histogram()
    .value(function (d) { return d })   // I need to give the vector of value
    .domain([0, domain])  // then the domain of the graphic
    .thresholds(x.ticks(100)); // then the numbers of bins

  // And apply this function to data to get the bins
  var bins = histogram(data);

  // Y axis: scale and draw:
  var y = d3.scaleLinear()
    .range([height, 0]);


  y.domain([0, 100]);   // d3.hist has to be called before the Y axis obviously
  // y.domain([0, d3.max(bins, function (d) { return d.length; })]);   // d3.hist has to be called before the Y axis obviously

  svg.append("g")
    .call(d3.axisLeft(y));

  // append the bar rectangles to the svg element
  svg.selectAll("rect")
    .data(bins)
    .enter()
    .append("rect")
    .attr("x", 1)
    .attr("transform", function (d) { return "translate(" + x(d.x0) + "," + y(d.length) + ")"; })
    .attr("width", function (d) { return x(d.x1) - x(d.x0) - 1; })
    .attr("height", function (d) { return height - y(d.length); })
    .style("fill", "#69b3a2")


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
    d3.select("svg").remove();
    var element = document.getElementById('tooltip')
    if (element) {
      element.remove();
    }

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
      .style("stroke", "#002499")
      .style("stroke-width", 3)
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
          div.html("<span class='bolded'>" + (d.value).toFixed(2) + ": </span>" + (d.likelihood * 100).toFixed(2) + "%")

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

function showSpreadAsColumnChartInTaskPane(cell: CellProperties) {
  d3.select("svg").remove();
  let data = cell.samples;
  let maxLikelihood = 0;
  data.forEach(function (d) {
    d.likelihood = (d.likelihood * 100) / 100;
    if (d.likelihood > maxLikelihood) {
      maxLikelihood = d.likelihood;
    }
    d.likelihood = +d.likelihood;
  });

  let margin = { top: 20, right: 20, bottom: 70, left: 40 };

  const width = 600 - margin.left - margin.right;
  const height = 300 - margin.top - margin.bottom;

  //Create the xScale
  const xScale = d3.scaleBand()
    .range([0, width]);

  //Create the yScale
  const yScale = d3.scaleLinear()
    .range([height, 0]);

  var xAxis = d3.axisBottom(xScale).scale(xScale);

  var yAxis = d3.axisLeft(yScale).ticks(10);

  const svg = d3.select("body").append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
    .append("g")
    .attr("transform", "translate(" + margin.left + "," + margin.top + ")");

  // const div = d3.select(".g-chart").append("div")
  //   .attr("class", "tooltip")
  //   .style("opacity", 0);

  //Organizes the data
  d3.max(data, function (d) { return d.value; });

  //Defines the xScale max
  xScale.domain(data.map((d) => d.value.toString()));

  //Defines the yScale max
  yScale.domain([0, maxLikelihood]);

  svg.append("g")
    .attr("class", "x axis")
    .attr("transform", "translate(0," + height + ")")
    .call(xAxis)
    .selectAll("text")
    .style("text-anchor", "end")
    .attr("dx", "-.8em")
    .attr("dy", "-.55em")
    .attr("transform", "rotate(-90)");

  svg.append("g")
    .attr("class", "y axis")
    .call(yAxis)
    .append("text")
    .attr("transform", "rotate(-90)")
    .attr("y", 5)
    .attr("dy", ".71em")
    .style("text-anchor", "end")
    .text("Frequency");


  svg.selectAll("bar")
    .data(data)
    .enter().append("rect")
    .attr("class", "bar")
    .attr("x", function (d) { return xScale(d.value.toString()); })
    .attr("width", xScale.bandwidth())
    .attr("y", function (d) {
      return yScale(d.likelihood
      );
    })
    .attr("height", function (d) { return height - yScale(d.likelihood); });
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
    await processWhatIf();
    document.getElementById('useNewValues').hidden = false;
    document.getElementById('dismissValues').hidden = false;
  } catch (error) {
    console.log(error);
  }
}

function performWhatIf() {
  Excel.run(function (context) {
    var worksheet = context.workbook.worksheets.getActiveWorksheet();
    console.log('Worksheet has changed');
    worksheet.onCalculated.add(processWhatIf);

    return context.sync()
      .then(function () {
        console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
      });
  }).catch((reason: any) => { console.log(reason) });
}


async function processWhatIf() {

  console.log('------------------Processing what-if');

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
  const whatif = new WhatIf();

  whatif.setNewCells(newCells, SheetProperties.cells, SheetProperties.referenceCell);

  console.log('Calculating the change');

  whatif.calculateChange();

  SheetProperties.cellOp.deleteUpdateshapes();

  whatif.showUpdateTextInCells(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);

  if (SheetProperties.isSpread) {
    console.log('Computing new spread');
    whatif.showNewSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
  }
}

async function useNewValues() {
  await parseSheet();
}

async function dismissValues() {

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let cellRanges = new Array<Excel.Range>();
    let cellValues = new Array<number>();

    SheetProperties.cells.forEach((cell: CellProperties) => {

      let range = sheet.getRange(cell.address);
      cellRanges.push(range.load('values'));
      cellValues.push(cell.value);
    })

    await context.sync();

    let i = 0;

    cellRanges.forEach((cellRange: Excel.Range) => {
      cellRange.values = [[cellValues[i]]];
      i++;
    })

  });
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
              // showSpreadAsColumnChartInTaskPane(cell);
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

function moveDivElement() {

  var div = document.getElementById('legend');
  console.log(div.style);
  div.style.position = 'relative';
  div.style.left = 5 + 'px';

}