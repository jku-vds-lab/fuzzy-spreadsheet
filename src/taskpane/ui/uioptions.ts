/* global console, document */
import * as d3 from 'd3';
import Bins from '../operations/bins';
import CellProperties from '../cell/cellproperties';
import { legendSize } from 'd3-svg-legend';
import { max } from 'd3';
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
    document.getElementById('startWhatIf').hidden = true;
    document.getElementById('useNewValues').hidden = true;
    document.getElementById('dismissValues').hidden = true;
  }

  public deSelectAllOoptions() {
    document.getElementById('referenceCell').hidden = true;
    document.getElementById('relationshipDiv').hidden = true;
    document.getElementById('neighborhoodDiv').hidden = true;
    document.getElementById('impactDiv').hidden = true;
    document.getElementById('likelihoodDiv').hidden = true;
    document.getElementById('spreadDiv').hidden = true;
    document.getElementById('startWhatIf').hidden = true;
    document.getElementById('useNewValues').hidden = true;
    document.getElementById('dismissValues').hidden = true;
    (<HTMLInputElement>document.getElementById('impact')).checked = false;
    (<HTMLInputElement>document.getElementById('likelihood')).checked = false;
    (<HTMLInputElement>document.getElementById('relationship')).checked = false;
    (<HTMLInputElement>document.getElementById('spread')).checked = false;
    (<HTMLInputElement>document.getElementById('inputRelationship')).checked = false;
    (<HTMLInputElement>document.getElementById('outputRelationship')).checked = false;
    (<HTMLInputElement>document.getElementById('zero')).checked = true;
    (<HTMLInputElement>document.getElementById('first')).checked = false;
    (<HTMLInputElement>document.getElementById('second')).checked = false;
    (<HTMLInputElement>document.getElementById('third')).checked = false;
    this.removeHtmlSpreadInfoForNewChart();
    this.removeHtmlSpreadInfoForOriginalChart();
    this.removeRelationshipInfoInTaskpane();
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
    this.drawLikelihoodLegend(-200);
    document.getElementById('spreadDiv').hidden = false;
    document.getElementById('startWhatIf').hidden = false;
    (<HTMLInputElement>document.getElementById("neighborhoodDiv")).disabled = true;
    (<HTMLInputElement>document.getElementById("impactDiv")).disabled = true;
    (<HTMLInputElement>document.getElementById("likelihoodDiv")).disabled = true;
    (<HTMLInputElement>document.getElementById("spreadDiv")).disabled = false;
    (<HTMLInputElement>document.getElementById("startWhatIf")).disabled = false;
  }

  public showAllOptions() {

    document.getElementById('relationshipDiv').hidden = false;
    document.getElementById('neighborhoodDiv').hidden = false;
    document.getElementById('impactDiv').hidden = false;
    document.getElementById('likelihoodDiv').hidden = false;
    document.getElementById('spreadDiv').hidden = false;
    document.getElementById('startWhatIf').hidden = false;
    (<HTMLInputElement>document.getElementById("neighborhoodDiv")).disabled = false;
    (<HTMLInputElement>document.getElementById("impactDiv")).disabled = false;
    (<HTMLInputElement>document.getElementById("likelihoodDiv")).disabled = false;
    (<HTMLInputElement>document.getElementById("spreadDiv")).disabled = false;
    (<HTMLInputElement>document.getElementById("startWhatIf")).disabled = false;
  }

  public isElementChecked(elementName: string) {
    let element = <HTMLInputElement>document.getElementById(elementName);

    if (element.checked) {
      return true;
    }
    return false;
  }
  public removeImpactInfoInTaskpane(id: string = 'impactPercentage') {
    // document.getElementById(id).innerHTML = '';
  }

  public removeRelationshipInfoInTaskpane() {

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
    } catch (error) {
      console.log(error);
    }
  }

  public removeHtmlSpreadInfoForNewChart() {
    try {
      d3.select("#" + 'whatIfChart').select('svg').remove();
      d3.select("#" + 'newLines').select('svg').remove();
      d3.select("#" + 'newSpreadLegend').select('svg').remove();
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

    // var impactText = document.getElementById(id);
    // // impactText.innerHTML = cell.impact + '%';
    // impactText.innerHTML = (Math.round(cell.impact * 100)/100).toFixed(2)  + '%';
    // impactText.style.position = 'relative';
    // impactText.style.left = 5 + 'px';

  }

  public addNewImpactPercentage(cell: CellProperties, id: string = 'newImpactPercentage') {

    // var newimpactText = document.getElementById(id);
    // newimpactText.innerHTML = (Math.round(cell.impact * 100)/100).toFixed(2)  + '%';
    // newimpactText.style.position = 'relative';
    // newimpactText.style.left = 5 + 'px';

  }

  public addLikelihoodPercentage(cell: CellProperties, id: string = 'likelihoodPercentage') {

    // var likelihoodText = document.getElementById(id);
    // likelihoodText.innerHTML = (cell.likelihood * 100).toFixed(2) + '%';
    // likelihoodText.style.position = 'relative';
    // likelihoodText.style.left = 5 + 'px';

  }

  public addNewLikelihoodPercentage(cell: CellProperties, id: string = 'newLikelihoodPercentage') {

    // var newLikelihoodText = document.getElementById(id);
    // newLikelihoodText.innerHTML = (cell.likelihood * 100).toFixed(2) + '%';
    // newLikelihoodText.style.position = 'relative';
    // newLikelihoodText.style.left = 5 + 'px';

  }

  public highlightInputRelationshipInfo(cell: CellProperties, n: number) {

    this.removeRelationshipInfoInTaskpane();

    if (!cell.isInputRelationship) {
      return;
    }

    if (cell.degreeOfRelationship == 1) {
      document.getElementById('diamond1').className = 'dotted';
    }


    if (cell.degreeOfRelationship == 2) {
      document.getElementById('diamond2').className = 'dotted';
    }


    if (cell.degreeOfRelationship == 3) {
      document.getElementById('diamond3').className = 'dotted';
    }
  }

  public highlightOutputRelationshipInfo(cell: CellProperties, n: number) {

    this.removeRelationshipInfoInTaskpane();

    if (!cell.isOutputRelationship) {
      return;
    }

    if (cell.degreeOfRelationship == 1) {
      document.getElementById('circle1').className = 'dotted';
    }


    if (cell.degreeOfRelationship == 2) {
      document.getElementById('circle2').className = 'dotted';
    }


    if (cell.degreeOfRelationship == 3) {
      document.getElementById('circle3').className = 'dotted';
    }
  }

  public showSpreadInTaskPane(cell: CellProperties, divClass: string = '.g-chart', color: string = '#399bfc', isLegendOrange: boolean = false) {

    try {


      let tooltipInfo = document.getElementById(divClass + "tooltip");

      if (tooltipInfo == null) {
        console.log('Null atm');
      } else {
        console.log('Deleting atm');
        document.getElementById(divClass + "tooltip").remove();
      }

      let data = cell.samples;

      let computedMean = cell.computedMean
      let computedStdDev = cell.computedStdDev;

      var margin = { top: 10, right: 30, bottom: 30, left: 40 },
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

      let toolTip = d3.select(divClass)
        .append("div")
        .attr("class", "tooltip")
        .attr("id", divClass + "tooltip");

      let mouseOver = function (d) {
        toolTip
          .style("opacity", 1)
        d3.select(this)
          .style("stroke", "black")
          .style("opacity", 1)
      }

      let mouseMove = function (d) {
        toolTip
          .html("P(" + d.x0 + " ≤ x < " + d.x1 + ") = " + d.length.toFixed(2) + "%")
        // .style("left", (d3.mouse(this)[0] + 70) + "px")
        // .style("top", (d3.mouse(this)[1]) + "px")
      }

      let mouseLeave = function (d) {
        toolTip
          .style("opacity", 0)
        d3.select(this)
          .style("stroke", "none")
          .style("opacity", 1)
      }


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
        .style("fill", color)
        .on('mouseover', mouseOver)
        .on("mousemove", mouseMove)
        .on("mouseleave", mouseLeave)




      svg.append("line")
        .data(bins)
        .attr("x1",
          function (d) {
            if (x(d.x0) == x(d.x1)) {
              return 1;
            }
            let x1 = ((-minDomain + computedMean) * x(d.x1) - x(d.x0)) / binWidth
            return x1;
          })
        .attr("y1", 130)
        .attr("x2", function (d) {
          if (x(d.x0) == x(d.x1)) {
            return 1;
          }
          let x1 = ((-minDomain + computedMean) * x(d.x1) - x(d.x0)) / binWidth
          return x1;
        })
        .attr("y2", 0)
        .style("stroke", "black")
        .style("stroke-dasharray", "4,4")

      svg.append("text")
        .data(bins)
        .attr("x",
          function (d) {
            if (x(d.x0) == x(d.x1)) {
              return 1;
            }
            let x1 = ((-minDomain + computedMean) * x(d.x1) - x(d.x0)) / binWidth
            return x1;
          })
        .attr("y", 0)
        .style("font-size", "10px")
        .style("fill", color)
        .text('M= ' + computedMean.toFixed(2));

      svg.append("line")
        .data(bins)
        .attr("x1",
          function (d) {
            if (x(d.x0) == x(d.x1)) {
              return 1;
            }
            let x1 = ((-minDomain + (computedMean - computedStdDev)) * x(d.x1) - x(d.x0)) / binWidth
            return x1;
          })
        .attr("y1", 20)
        .attr("x2", function (d) {
          if (x(d.x0) == x(d.x1)) {
            return 1;
          }
          let x1 = ((-minDomain + (computedMean - computedStdDev)) * x(d.x1) - x(d.x0)) / binWidth
          return x1;
        })
        .attr("y2", 10)
        .style("stroke", "grey")


      svg.append("text")
        .data(bins)
        .attr("x",
          function (d) {
            if (x(d.x0) == x(d.x1)) {
              return 1;
            }
            let x1 = ((-minDomain + (computedMean - computedStdDev)) * x(d.x1) - x(d.x0)) / binWidth
            return x1;
          })
        .attr("y", 10)
        .style("font-size", "10px")
        .style("fill", color)
        .text('-SD');

      svg.append("line")
        .data(bins)
        .attr("x1",
          function (d) {
            if (x(d.x0) == x(d.x1)) {
              return 1;
            }
            let x1 = ((-minDomain + (computedMean + computedStdDev)) * x(d.x1) - x(d.x0)) / binWidth
            return x1;
          })
        .attr("y1", 20)
        .attr("x2", function (d) {
          if (x(d.x0) == x(d.x1)) {
            return 1;
          }
          let x1 = ((-minDomain + (computedMean + computedStdDev)) * x(d.x1) - x(d.x0)) / binWidth
          return x1;
        })
        .attr("y2", 10)
        .style("stroke", "grey")


      svg.append("text")
        .data(bins)
        .attr("x",
          function (d) {
            if (x(d.x0) == x(d.x1)) {
              return 1;
            }
            let x1 = ((-minDomain + (computedMean + computedStdDev)) * x(d.x1) - x(d.x0)) / binWidth
            return x1;
          })
        .attr("y", 10)
        .style("font-size", "10px")
        .style("fill", color)
        .text('+SD');


      svg.append("line")
        .data(bins)
        .attr("x1",
          function (d) {
            if (x(d.x0) == x(d.x1)) {
              return 1;
            }
            let x1 = ((-minDomain + (computedMean + computedStdDev)) * x(d.x1) - x(d.x0)) / binWidth
            return x1;
          })
        .attr("y1", 15)
        .attr("x2", function (d) {
          if (x(d.x0) == x(d.x1)) {
            return 1;
          }
          let x1 = ((-minDomain + (computedMean - computedStdDev)) * x(d.x1) - x(d.x0)) / binWidth
          return x1;
        })
        .attr("y2", 15)
        .style("stroke", "grey");

      svg.append('text')
        .attr("transform", "rotate(-90)")
        .attr("y", 0 - margin.left)
        .attr("x", 0 - (height / 2))
        .attr("dy", "1em")
        .style("text-anchor", "middle")
        .style("font-size", "10px")
        .text('Probability in %');

      svg.append('text')
        .attr("transform",
          "translate(" + width + " ," +
          (height / 2) + ")")
        .style("text-anchor", "middle")
        .style("font-size", "10px")
        .style('fill', color)
        .text('SD=' + computedStdDev.toFixed(2));

      svg.append('text')
        .attr("transform",
          "translate(" + width / 2 + " ," +
          (height + margin.bottom) + ")")
        .style("text-anchor", "middle")
        .style("font-size", "10px")
        .text('Values in Mio.(€)');

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



  public drawImpactLegend(impact: number = -1, newImpact: number = -1, color: string = 'green') {


    d3.select("#impactLegend").select('svg').remove();
    let impactTemp = Math.ceil(impact * 0.5);
    let newImpactTemp = Math.ceil(newImpact * 0.5);

    if (color == 'green') {
      impactTemp = impactTemp + 50;
      newImpactTemp = newImpactTemp + 50;
    } else {
      impactTemp = 50 - impactTemp;
      newImpactTemp = 50 - newImpactTemp;
    }


    if (impact == -1) {
      impactTemp = -1;
    }

    if (newImpact == -1) {
      newImpactTemp = -1;
    }

    const minDomain = -5;
    const maxDomain = 40;
    const binWidth = 1;

    let binsObj = new Bins(minDomain, maxDomain, binWidth);
    var colors = binsObj.generateRedGreenColors();

    var Svg = d3.select('#impactLegend').append("svg")
      .attr("width", 200)
      .attr("height", 30);

    Svg.selectAll("mydots")
      .data(colors)
      .enter()
      .append("rect")
      .attr("x", function (d, i) { return (i) * 2 })
      .attr("y", function (d, i) {
        if (i == impactTemp || i == newImpactTemp) {
          return 15;
        }
        return 20;
      })
      .attr("width", function (d, i) {
        if (i == impactTemp || i == newImpactTemp) {
          return 2;
        }
        return 1;
      })
      .attr("height", function (d, i) {
        if (i == impactTemp || i == newImpactTemp) {
          return 15;
        }
        return 5;
      }
      )
      .style("fill", function (d, i) {
        if (i == impactTemp) {
          return "blue";
        } if (i == newImpactTemp) {
          return "orange";
        }
        return d;
      }
      );

    // add legend for impact
    Svg.selectAll("text")
      .data(colors)
      .enter()
      .append("text")
      .text(function (d, i) {
        if (i == impactTemp) {
          return impact + ' %';
        } if (i == newImpactTemp) {
          return newImpact + ' %';
        }
        return " ";
      })
      .style("fill", function (d, i) {
        if (i == impactTemp) {
          return "blue";
        } if (i == newImpactTemp) {
          return "orange";
        }
        return " ";
      })
      .style("font-size", function (d, i) {
        if (i == impactTemp || i == newImpactTemp) {
          return "10px";
        }
        return "14px";
      })
      .attr("x", function (d, i) { return (i) * 2 })
      .attr("y", function (d, i) {
        if (i == impactTemp) {
          return 10;
        }
        if (i == newImpactTemp) {
          return 10;
        }
        return 15;
      });
  }


  public drawLikelihoodLegend(likelihood: number = 0, newLikelihood: number = -1) {

    d3.select("#likelihoodLegend").select('svg').remove();

    let sizeArray = [0, 20, 40, 60, 80, 100];
    let sizeArrayText = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100];

    likelihood = likelihood * 100;

    var Svg = d3.select('#likelihoodLegend').append("svg")
      .attr("width", 200)
      .attr("height", 60);

    // add indicators for likelihood of occurrence (sqaures in grey)
    Svg.selectAll("mySquares")
      .data(sizeArray)
      .enter()
      .append("rect")
      .attr("x", function (d, i) { return (i) * (i + 1) * (sizeArray.length - 2.5) })
      .attr("y", function (d, i) {
        return Math.max.apply(null, sizeArray) / 3 - (i - 1) * sizeArray.length + 20;
      })
      .attr("width", function (d, i) {
        return d / 3;
      })
      .attr("height", function (d, i) {
        return d / 3;
      }
      )
      // .style("fill", (d) => { return d });
      .style("fill", function (d, i) {
        return "grey";
      }
      );

    Svg.selectAll("mySquaresIndicators")
      .data(sizeArray)
      .enter()
      .append("rect")
      .attr("x", function (d, i) { return (i) * (i + 1) * (sizeArray.length - 2.5) })
      .attr("y", function (d, i) {
        // return Math.max.apply(null, sizeArrayText) / 3 - d / 4;
        return Math.max.apply(null, sizeArray) / 3 - (i - 1) * sizeArray.length + 17;

      })
      .attr("width", function (d, i) {
        // if (d == likelihood || d == newLikelihood) {
        //   return 2;
        // }
        return d / 3;
      })
      .attr("height", function (d, i) {
        if (d == likelihood || d == newLikelihood) {
          return 2;
        }
        return d / 3;
      }
      )
      .style("fill", function (d, i) {
        if (d == likelihood) {
          return "blue";
        }
        return "rgba(0,0,0,0)";
      }
      );

    // add legend for impact
    Svg.selectAll("text")
      .data(sizeArrayText)
      .enter()
      .append("text")
      .text(function (d, i) {
        if (d == likelihood) {
          return likelihood + ' %';
        } if (d == newLikelihood) {
          return newLikelihood + ' %';
        }
        return " ";
      })
      .style("fill", function (d, i) {
        if (d == likelihood) {
          return "blue";
        } if (d == newLikelihood) {
          return "orange";
        }
        return " ";
      })
      .style("font-size", function (d, i) {
        if (d == likelihood || d == newLikelihood) {
          return "10px";
        }
        return "14px";
      })
      .attr("x", function (d, i) { return (d / 9) * (i) })
      .attr("y", function (d, i) {
        // return Math.max.apply(null,sizeArray)/3-(i-1)*sizeArray.length;
        // return d -5;
        return 100 / 3 - d / sizeArrayText.length - 2 * i + 18;
      });
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


// public static findPercentile(samples: number[], value: number = 11.1) {
//   return ((MainClass.cutoff(samples, value) / samples.length) * 100).toFixed(2);
// }

// public static cutoff(sortedValues, value, start = 0, end = sortedValues.length) {
//   if (sortedValues[end - 1] <= value) { return -1 }

//   while (start !== end - 1) {
//     const index = Math.floor((end + start) / 2)
//     if (sortedValues[index] <= value) {
//       start = index
//     } else {
//       end = index
//     }
//   }
//   return end
// }