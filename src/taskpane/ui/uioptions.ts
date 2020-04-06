/* global console, document */
import * as d3 from 'd3';
import Bins from '../operations/bins';
import CellProperties from '../cell/cellproperties';
export default class UIOptions {
  constructor() {
  }

  public hideOptions(isReferenceCellHidden: boolean = true) {
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

  public showReferenceCellOption() {
    document.getElementById('referenceCell').hidden = false;
  }

  public showVisualizationOption() {

    document.getElementById('relationshipDiv').hidden = false;
    document.getElementById('neighborhoodDiv').hidden = false;
    document.getElementById('impactDiv').hidden = false;
    this.drawImpactLegend(-200);
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

  public showAllOptions() {

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

  public isElementChecked(elementName: string) {
    let element = <HTMLInputElement>document.getElementById(elementName);

    if (element.checked) {
      return true;
    }
    return false;
  }
  public removeImpactInfoInTaskpane(id: string = 'impactPercentage') {
    document.getElementById(id).innerHTML = '';
  }

  public removeRelationshipInfoInTaskpane() {
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

  public removeHtmlSpreadInfoForOriginalChart() {
    try {
      d3.select("#" + 'originalChart').select('svg').remove();
      d3.select("#" + 'lines').select('svg').remove();
      d3.select("#" + 'spreadLegend').select('svg').remove();
      document.getElementById("mean").innerHTML = "";
    } catch (error) {
      console.log(error);
    }
  }

  public removeHtmlSpreadInfoForNewChart() {
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

  public addHtmlSpreadInfoForNewChart() {
    document.getElementById("newDistribution").hidden = false;
    document.getElementById("spaceHack").hidden = false;
  }

  public addImpactPercentage(cell: CellProperties, id: string = 'impactPercentage') {

    var impactText = document.getElementById(id);
    impactText.innerHTML = cell.impact + '%';
    impactText.style.position = 'relative';
    impactText.style.left = 5 + 'px';

  }

  public addLikelihoodPercentage(cell: CellProperties, id: string = 'likelihoodPercentage') {

    var likelihoodText = document.getElementById(id);
    likelihoodText.innerHTML = (cell.likelihood * 100).toFixed(2) + '%';
    likelihoodText.style.position = 'relative';
    likelihoodText.style.left = 5 + 'px';

  }

  public highlightInputRelationshipInfo(cell: CellProperties, n: number) {

    this.removeRelationshipInfoInTaskpane();

    if (!cell.isInputRelationship) {
      return;
    }

    if (cell.degreeToFocus == 1) {
      document.getElementById('diamond1').className = 'dotted';
      document.getElementById('number1').className = 'dotted';
    }

    if (n == 2) {
      if (cell.degreeToFocus > 1) {
        document.getElementById('diamond2').className = 'dotted';
        document.getElementById('number2').className = 'dotted';
      }
    }

    if (n == 3) {

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

  public highlightOutputRelationshipInfo(cell: CellProperties, n: number) {

    this.removeRelationshipInfoInTaskpane();

    if (!cell.isOutputRelationship) {
      return;
    }

    if (cell.degreeToFocus == 1) {
      document.getElementById('circle1').className = 'dotted';
      document.getElementById('number1').className = 'dotted';
    }

    if (n == 2) {
      if (cell.degreeToFocus > 1) {
        document.getElementById('circle2').className = 'dotted';
        document.getElementById('number2').className = 'dotted';
      }
    }

    if (n == 3) {
      if (cell.degreeToFocus == 2) {
        document.getElementById('circle2').className = 'dotted';
        document.getElementById('number2').className = 'dotted';
      } else {
        document.getElementById('circle3').className = 'dotted';
        document.getElementById('number3').className = 'dotted';
      }
    }
  }

  public showMeanAndStdDevValueInTaskpane(cell: CellProperties) {
    document.getElementById("mean").innerHTML = "Mean: " + cell.computedMean.toFixed(2) + " & Std Dev: " + cell.computedStdDev.toFixed(2);
  }

  public showNewMeanAndStdDevValueInTaskpane(cell: CellProperties) {
    document.getElementById("newMean").innerHTML = "Mean: " + cell.computedMean.toFixed(2) + " & Std Dev: " + cell.computedStdDev.toFixed(2);
  }

  public showSpreadInTaskPane(cell: CellProperties, divClass: string = '.g-chart', color: string = '#399bfc', isLegendOrange: boolean = false) {

    try {

      let data = cell.samples;

      var margin = { top: 10, right: 30, bottom: 20, left: 40 },
        width = 260 - margin.left - margin.right,
        height = 160 - margin.top - margin.bottom;

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
        .call(d3.axisBottom(x).tickValues(ticks))

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

      svg.append('text')
        .attr("transform", "rotate(-90)")
        .attr("y", 0 - margin.left)
        .attr("x", 0 - (height / 2))
        .attr("dy", "1em")
        .style("text-anchor", "middle")
        .style("font-size", "10px")
        .text('Probability in %');

      // svg.append('text')
      //   .attr("transform",
      //     "translate(" + width / 2 + " ," +
      //     (height + margin.bottom) + ")")
      //   .style("text-anchor", "middle")
      //   .style("font-size", "10px")
      //   .text('Values in Mio.(€)');



      if (isLegendOrange) {
        this.drawLinesBeneathChart(cell, isLegendOrange);
        this.drawLegend(isLegendOrange);
      } else {
        this.drawLinesBeneathChart(cell);
        this.drawLegend();

      }

    } catch (error) {
      console.log(error);
    }
  }

  public drawLinesBeneathChart(cell: CellProperties, isLegendOrange: boolean = false) {

    try {

      var colors;

      let div = '#lines';

      if (isLegendOrange) {
        div = '#newLines';
        colors = cell.binOrangeColors;
      } else {
        colors = cell.binBlueColors;
      }

      if (colors == undefined) {
        return;
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
    } catch (error) {
      console.log(error);
    }
  }

  public drawLegend(isLegendOrange: boolean = false) {

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


  public drawImpactLegend(impact: number = 0, color: string = 'green') {


    d3.select("#impactLegend").select('svg').remove();
    impact = Math.ceil(impact * 0.5);

    if (color == 'green') {
      impact = impact + 50;
    } else {
      impact = 50 - impact;
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

  public showWhatIfOptions() {
    document.getElementById('useNewValues').hidden = false;
    document.getElementById('dismissValues').hidden = false;
  }

  public hideWhatIfOptions() {
    document.getElementById('useNewValues').hidden = true;
    document.getElementById('dismissValues').hidden = true;
  }

}