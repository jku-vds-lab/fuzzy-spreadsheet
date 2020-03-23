/* global console, Excel */
import { ceil } from 'mathjs';
import * as jstat from 'jstat';
import CellProperties from '../cellproperties';
import SheetProperties from '../sheetproperties';
import * as outliers from 'outliers';

import { range, dotMultiply } from 'mathjs';
import { Bernoulli } from 'discrete-sampling';
import * as jStat from 'jstat';
import * as d3 from 'd3';
import CellOperations from '../celloperations';


// code cleaning required
// change to heatmap encoding
export default class Spread {

  private cells: CellProperties[];
  private oldCells: CellProperties[];
  private referenceCell: CellProperties;
  private oldReferenceCell: CellProperties;
  private color: string = null;

  constructor(cells: CellProperties[], oldCells: CellProperties[], referenceCell: CellProperties, color: string = null) {
    this.cells = cells;
    this.oldCells = oldCells;
    this.referenceCell = referenceCell;
    this.color = color;

    if (this.oldCells == null) {
      return;
    }

    this.oldCells.forEach((oldCell: CellProperties) => {
      if (oldCell.id == this.referenceCell.id) {
        this.oldReferenceCell = oldCell;
      }
    })
  }

  public showSpread(n: number, isInput: boolean, isOutput: boolean) {

    try {

      this.addVarianceInfo();

      this.showReferenceCellSpread();

      if (isInput) {
        this.showInputSpread(this.referenceCell.inputCells, n);
      }

      if (isOutput) {
        this.showOutputSpread(this.referenceCell.outputCells, n);
      }

    } catch (error) {
      console.log('Error in Show spread', error);
    }
  }

  public showReferenceCellSpread() {

    try {

      if (this.referenceCell.isSpread) {
        console.log('Returning because reference cell has a spread');
        return;
      }


      this.referenceCell.isSpread = true;
      this.addSamplesToCell(this.referenceCell, this.oldReferenceCell);
      this.showBarCodePlot(this.referenceCell, this.oldReferenceCell, 'ReferenceChart');

    } catch (error) {
      console.log('Error in Show Reference Cell Spread', error);
    }
  }

  public showInputSpread(cells: CellProperties[], i: number) {

    try {
      cells.forEach((cell: CellProperties) => {

        if (cell.isSpread) {
          console.log(cell.address + ' already has a spread');
          return;
        }

        cell.isSpread = true;

        let oldCell = null;

        if (this.oldCells != null) {
          oldCell = this.oldCells.find((oldCell: CellProperties) => oldCell.id == cell.id)
        }

        this.addSamplesToCell(cell, oldCell);

        if (cell.samples == null) {
          cell.isSpread = false;
          return;
        }

        this.showBarCodePlot(cell, oldCell, 'InputChart');

        if (i == 1) {
          return;
        }
        this.showInputSpread(cell.inputCells, i - 1);
      })

    } catch (error) {
      console.log('Error in Show Input Cell Spread', error);
    }
  }

  public showOutputSpread(cells: CellProperties[], i: number) {

    try {

      cells.forEach((cell: CellProperties) => {

        if (cell.isSpread) {
          console.log(cell.address + ' already has a spread');
          return;
        }

        cell.isSpread = true;

        let oldCell = null;

        if (this.oldCells != null) {
          oldCell = this.oldCells.find((oldCell: CellProperties) => oldCell.id == cell.id)
        }

        this.addSamplesToCell(cell, oldCell);

        if (cell.samples == null) {
          cell.isSpread = false;
          return;
        }

        this.showBarCodePlot(cell, oldCell, 'OutputChart');

        if (i == 1) {
          return;
        }
        this.showOutputSpread(cell.outputCells, i - 1);
      })

    } catch (error) {
      console.log('Error in Show Output Cell Spread', error);
    }
  }

  public showBarCodePlot(cell: CellProperties, oldCell: CellProperties, name: string) {
    try {

      if (cell.samples == null) {
        return;
      }

      if (oldCell == null) {
        this.drawBarCodePlot(cell, 'blue', name);
        return;
      }
      // remove the original bar code plot
      this.removeSpreadCellWise(oldCell);
      // add old bar code plot with half the length
      this.drawBarCodePlot(oldCell, 'blue', name, true)
      // add new bar code plot with half the length
      name = 'Update' + name;
      this.drawBarCodePlot(cell, 'orange', name, false, true);

    } catch (error) {
      console.log('Could not draw the bar code plot because of the following error', error);
    }
  }


  generateColor() {
    let i = 0;
    let darkColors = [];
    let r = 235;
    let g = 255;
    let b = 255;

    while (i < 20) {

      darkColors.push(d3.rgb(r, g, b).hex());
      r = 0;
      g = 255 - i * 10;
      b = 255 - i * 10;
      i++;
    }
    console.log('Dark Colors', darkColors);
    return darkColors;
  }

  public drawBarCodePlot(cell: CellProperties, color: string, name: string, isUpperHalf: boolean = false, isLowerHalf: boolean = false) {
    try {

      Excel.run((context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let totalHeight = cell.height;

        let startLineTop = cell.top;
        let startLineLeft = cell.left + 20;
        let endLineTop = cell.top + totalHeight;

        if (isUpperHalf) {
          endLineTop = cell.top + 0.5 * totalHeight;
        }

        if (isLowerHalf) {
          startLineTop = cell.top + 0.5 * totalHeight;
        }

        let sortedLinesWithColors = this.computeColorsAndBins(cell, color);

        sortedLinesWithColors.forEach((el) => {
          let rect = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
          rect.name = cell.address + name;
          rect.top = startLineTop;
          rect.left = startLineLeft + el.value;
          rect.width = 3;
          rect.height = totalHeight;
          rect.fill.setSolidColor(el.color);
          rect.lineFormat.color = el.color;
        })
        return context.sync();
      });
    } catch (error) {
      console.log('Could not draw the bar code plot because of the following error', error);
    }
  }

  private computeColorsAndBins(cell: CellProperties, color: string) {

    let sortedLinesWithColors = new Array<{ value: number, color: string }>();

    try {

      let blueColors = ['#DDE1E4', '#D5DADD', '#CAD1D5', '#BDC6CA', '#ACB8BD', '#97A6AC', '#7D9097', '#5C747D', '#33515D', '#002534']; //['#d8e1e7', '#98b0c2', '#4e7387', '#002e41', '#002534', '#001e2a', '#001822'] // light to dark
      let orangeColors = ['#ffe5cc', '#ffcc99', '#ffb266', '#ff9933', '#ff8000', '#cc6600', '#a35200'];

      // let colors = blueColors;

      let colors = this.generateColor();

      if (color == 'orange') {
        colors = orangeColors;
      }

      let data = cell.samples;

      if (cell.samples.length == 1) {
        const element = { value: cell.samples[0], color: colors[4] }
        sortedLinesWithColors.push(element);
        return sortedLinesWithColors;
      }

      var count = 10;

      let domain = d3.max(data, function (d) { return +d })

      domain = 35;

      var x = d3.scaleLinear().domain([0, domain]).nice(count);

      var histogram = d3.histogram().value(function (d) { return d }).domain([0, domain]).thresholds(x.ticks(count));
      var bins = histogram(data);

      let sortBinByValues = Object.assign([], bins);
      let sortBinByFreq = Object.assign([], bins);

      sortBinByValues.sort((n1, n2) => { return n1.x0 - n2.x0 });
      sortBinByFreq.sort((n1, n2) => { return n1.length - n2.length });

      sortBinByValues.forEach((valueBin) => {
        let binColor = '';
        let binValue = valueBin.x0;

        sortBinByFreq.forEach((freqBin, index: number) => {

          if (binValue == freqBin.x0) {

            if (freqBin.length == 0) {
              binColor = colors[0];
            } else {
              binColor = colors[index];
            }
            return;
          }
        })

        const element = { value: binValue, color: binColor };
        sortedLinesWithColors.push(element);
      })


    } catch (error) {
      console.log(error);
    }
    return sortedLinesWithColors;
  }

  public addSamplesToCell(cell: CellProperties, oldCell: CellProperties) {

    try {

      cell.samples = new Array<number>();

      const mean = cell.value;
      const variance = cell.variance;
      const likelihood = cell.likelihood;

      // temporary check: to be removed after adding mean & variance value to the formula cells
      if (oldCell != null && !cell.formula.includes('SUM') && !cell.formula.includes('-')) {
        const oldMean = oldCell.value;
        const oldVariance = oldCell.variance;
        const oldLikelihood = oldCell.likelihood;

        if (mean == oldMean) {
          if (variance == oldVariance) {
            if (likelihood == oldLikelihood) {
              cell.samples = null;
              return;
            }
          }
        }
      }

      if (variance == 0 && likelihood == 1) {
        cell.samples.push(mean);
      }
      else {

        if (cell.formula.includes('SUM')) {
          cell.samples = this.addSamplesToSumCell(cell);
          return;
        }

        if (cell.formula.includes('-')) {
          cell.samples = this.addSamplesToSumCell(cell, true);
          return;
        }

        cell.samples = this.addSamplesToAverageCell(cell);
      }
    } catch (error) {
      console.log(error);
    }
  }

  public addSamplesToAverageCell(cell: CellProperties) {

    try {

      const mean = cell.value;
      const variance = cell.variance;
      const likelihood = cell.likelihood;

      cell.samples = new Array<number>();

      let normalSamples = new Array<number>();
      const values = <number[]>range(0, 1, 0.01).toArray(); // for 100 samples

      values.forEach((val: number) => {
        normalSamples.push(jStat.normal.inv(val, mean, variance));
      })

      normalSamples = normalSamples.filter(outliers());

      const sampleLength = normalSamples.length;

      const bern = Bernoulli(likelihood);
      bern.draw();

      const bernoulliSamples = bern.sample(sampleLength);

      cell.samples = <number[]>dotMultiply(normalSamples, bernoulliSamples);

    } catch (error) {
      console.log('Error in Average Spread Computation', error);
    }

    return cell.samples;
  }

  public addSamplesToSumCell(cell: CellProperties, isDifference: boolean = false) {

    try {

      let inputCells = cell.inputCells;

      cell.inputCells.forEach((inCell: CellProperties) => {

        let oldCell = null;

        if (this.oldCells != null) {
          oldCell = this.oldCells.find((oldCell: CellProperties) => oldCell.id == cell.id)
        }
        this.addSamplesToCell(inCell, oldCell);
      })

      let index = 0;

      let resultantSample = new CellProperties();
      resultantSample.samples = new Array<number>();

      if (inputCells.length > 1) {
        resultantSample = this.addTwoSamples(inputCells[index], inputCells[index + 1], isDifference);
        index = index + 2;
      }

      while (index < inputCells.length) {

        resultantSample = this.addTwoSamples(resultantSample, inputCells[index], isDifference);
        index = index + 1;
      }

      resultantSample.samples.forEach((sample: number) => {
        cell.samples.push(sample);
      })
    } catch (error) {
      console.log('Error in Average Spread Computation', error);
    }
    return cell.samples;
  }

  private addTwoSamples(sample1: CellProperties, sample2: CellProperties, isDifference: boolean = false) {

    let resultantSample = new CellProperties();
    resultantSample.samples = new Array<number>();

    try {

      sample1.samples.forEach((sampleCell1: number, index: number) => {

        if (sample2.samples.length <= index) {
          return;
        }

        let value = sampleCell1 + sample2.samples[index];

        if (isDifference) {
          value = sampleCell1 - sample2.samples[index];
        }

        resultantSample.samples.push(value);

      })
    } catch (error) {
      console.log(error);
    }

    return resultantSample;
  }

  public addVarianceInfo() {

    try {
      for (let i = 0; i < this.cells.length; i++) {
        this.cells[i].variance = 0;

        if (this.cells[i].isUncertain) {

          this.cells[i].variance = this.cells[i + 1].value;
          this.cells[i].likelihood = this.cells[i + 2].value;
        }
      }
    } catch (error) {
      console.log(error);
    }
  }

  public removeSpread(isInput: boolean, isOutput: boolean, isRemoveAll: boolean) {

    this.cells.forEach((cell: CellProperties) => {
      cell.isSpread = false;
    })

    let name: string;

    if (!isInput) {
      name = 'InputChart';
      this.deleteBarCodePlot(name);
    }

    if (!isOutput) {
      name = 'OutputChart';
      this.deleteBarCodePlot(name);
    }

    if (isRemoveAll) {
      this.deleteBarCodePlot('InputChart');
      this.deleteBarCodePlot('OutputChart');
    }
  }

  public removeSpreadFromReferenceCell() {
    this.referenceCell.isSpread = false;
    this.deleteBarCodePlot('ReferenceChart');
  }

  public removeSpreadCellWise(cell: CellProperties) {
    this.deleteBarCodePlot(cell.address);
  }


  deleteBarCodePlot(name: string) {

    try {

      Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        var shapes = sheet.shapes;
        shapes.load("items/name");

        return context.sync().then(function () {
          shapes.items.forEach(function (shape) {
            if (shape.name.includes(name)) {
              shape.delete();
            }
          });
        }).catch((reason: any) => {
          console.log('Step 1:', reason, name)
        });
      });
    } catch (error) {
      console.log('Step 2:', error);
    }
  }

  async asyncDeleteBarCodePlot(name: string) {

    try {

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        var shapes = sheet.shapes;
        shapes.load("items/name");

        await context.sync();

        shapes.items.forEach(function (shape) {
          if (shape.name.includes(name)) {
            shape.delete();
          }
        })

        await context.sync();
      });
    } catch (error) {
      console.log('Async Delete Error:', error);
    }
  }
}