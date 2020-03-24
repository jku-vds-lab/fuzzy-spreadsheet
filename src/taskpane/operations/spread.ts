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
import { lineRadial } from 'd3';


// code cleaning required
// change to heatmap encoding
export default class Spread {

  private cells: CellProperties[];
  private oldCells: CellProperties[];
  private referenceCell: CellProperties;
  private oldReferenceCell: CellProperties;
  private colors: string[];
  private blueColors: string[];
  private orangeColors: string[];
  private nrOfColors: number;

  constructor(cells: CellProperties[], oldCells: CellProperties[], referenceCell: CellProperties) {
    this.cells = cells;
    this.oldCells = oldCells;
    this.referenceCell = referenceCell;
    this.nrOfColors = 16;
    this.blueColors = this.generateBlueColors(this.nrOfColors);
    this.orangeColors = this.generateOrangeColors(this.nrOfColors);
    this.colors = this.blueColors;

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
      this.makeFontColorWhite(n, isInput, isOutput);
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


  makeFontColorWhite(n: number, isInput: boolean, isOutput: boolean) {

    this.changeFontToWhite(this.referenceCell.address);

    if (isInput) {
      this.makeInputFontWhite(this.referenceCell.inputCells, n);
    }

    if (isOutput) {
      this.makeOutputFontWhite(this.referenceCell.outputCells, n);
    }
  }

  makeInputFontWhite(cells: CellProperties[], n: number) {

    try {
      cells.forEach((cell: CellProperties) => {

        this.changeFontToWhite(cell.address);
        if (n == 1) {
          return;
        }
        this.makeInputFontWhite(cell.inputCells, n - 1);
      })

    } catch (error) {
      console.log('Error', error);
    }
  }

  makeOutputFontWhite(cells: CellProperties[], n: number) {
    try {
      cells.forEach((cell: CellProperties) => {

        this.changeFontToWhite(cell.address);
        if (n == 1) {
          return;
        }
        this.makeOutputFontWhite(cell.outputCells, n - 1);
      })

    } catch (error) {
      console.log('Error', error);
    }
  }

  changeFontToWhite(address: string) {

    try {
      Excel.run(function (context) {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let range = sheet.getRange(address);
        range.format.font.color = 'white';
        return context.sync();
      });

    } catch (error) {
      console.log(error);
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

        let oldCell = null;

        if (this.oldCells != null) {
          oldCell = this.oldCells.find((oldCell: CellProperties) => oldCell.id == cell.id)
        }

        if (cell.isSpread) {
          console.log(cell.address + ' already has a spread');
          return;
        }

        cell.isSpread = true;

        if (cell.samples == null) {
          this.addSamplesToCell(cell, oldCell);
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

        if (cell.samples == null) {
          this.addSamplesToCell(cell, oldCell);
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

      if (oldCell == null) {
        this.drawBarCodePlot(cell, name);
        return;
      }

      if (cell.samples == oldCell.samples) {
        console.log('Returning because samples are the same for :' + cell.address);
        return;
      }

      // remove the original bar code plot
      this.removeSpreadCellWise(oldCell);
      // add old bar code plot with half the length
      this.drawBarCodePlot(oldCell, name, true);
      // add new bar code plot with half the length
      name = 'Update' + name;
      this.drawBarCodePlot(cell, name, false, true);
    } catch (error) {
      console.log('Could not draw the bar code plot because of the following error', error);
    }
  }


  generateBlueColors(n: number) {

    let i = 0;
    let blueColors = [];
    let r = 216;
    let g = 255;
    let b = 255;
    let factor = 255 / n;

    while (i < n) {

      blueColors.push(d3.rgb(r, g, b).hex());
      r = 0;
      g = 255 - i * factor;
      b = 255 - i * factor;
      i++;
    }
    return blueColors;
  }

  generateOrangeColors(n: number) {

    let i = 0;
    let orangeColors = [];
    let r = 255;
    let g = 248;
    let b = 235;
    let greenFactor = 165 / n;
    let redFactor = 235 / n;

    while (i < n) {

      orangeColors.push(d3.rgb(r, g, b).hex());
      r = 235 - i * redFactor;
      g = 165 - i * greenFactor;
      b = 0;
      i++;
    }
    return orangeColors;
  }

  public drawBarCodePlot(cell: CellProperties, name: string, isUpperHalf: boolean = false, isLowerHalf: boolean = false) {
    try {

      let isDrawLine = false;

      Excel.run((context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let height = cell.height;

        if (isUpperHalf || isLowerHalf) {
          height = height / 2;
        }

        let top = cell.top;
        let left = cell.left + 20;

        this.colors = this.blueColors;

        if (isLowerHalf) {
          top = cell.top + height;
          this.colors = this.orangeColors; // always use orange colors in the bottom half
          isDrawLine = true;
        }

        let sortedLinesWithColors = this.computeColorsAndBins(cell);

        if (isDrawLine) {
          let line = sheet.shapes.addLine(left, top, left + cell.width - 20, top);
          line.name = cell.address + name;
          line.lineFormat.color = 'white';
          line.lineFormat.weight = 2;
        }

        sortedLinesWithColors.forEach((el) => {
          let rect = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
          rect.name = cell.address + name;
          rect.top = top;
          rect.left = left + el.value;
          rect.width = 3;
          rect.height = height;
          rect.fill.setSolidColor(el.color);
          rect.lineFormat.color = el.color;
        })
        return context.sync();
      });
    } catch (error) {
      console.log('Could not draw the bar code plot because of the following error', error);
    }
  }

  private computeColorsAndBins(cell: CellProperties) {

    let sortedLinesWithColors = new Array<{ value: number, color: string }>();

    try {

      let data = cell.samples;

      if (cell.samples.length == 1) {
        const element = { value: cell.samples[0], color: this.colors[this.nrOfColors - 1] }
        sortedLinesWithColors.push(element);
        return sortedLinesWithColors;
      }

      var count = 10;

      let domain = 36;

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
              binColor = this.colors[0];
            } else {
              binColor = this.colors[index];
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
      const stdev = cell.stdev;
      const likelihood = cell.likelihood;

      let isCellUnary = false;


      if (cell.formula.includes('SUM')) {
        cell.samples = this.addSamplesToSumCell(cell);
      }

      if (cell.formula.includes('-')) {
        cell.samples = this.addSamplesToSumCell(cell, true);
      }

      if (cell.formula.includes('AVERAGE')) {
        cell.samples = this.addSamplesToAverageCell(cell);
        isCellUnary = true;
      }

      // fix values for certain cells
      if (cell.formula == "") {
        if (stdev == 0 && likelihood == 1) {
          cell.samples.push(mean);
        }
        isCellUnary = true;
      }

      cell.computedMean = jStat.mean(cell.samples);
      cell.computedStdDev = jStat.stdev(cell.samples);

      if (oldCell == null) {
        return;
      }

      if (isCellUnary) {

        const oldMean = oldCell.value;
        const oldStdev = oldCell.stdev;
        const oldLikelihood = oldCell.likelihood;

        if (mean == oldMean) {
          if (stdev == oldStdev) {
            if (likelihood == oldLikelihood) {
              cell.samples = oldCell.samples;
            }
          }
        }
      }
    } catch (error) {
      console.log(error);
    }
  }

  public addSamplesToAverageCell(cell: CellProperties) {

    try {

      const mean = cell.value;
      const stdev = cell.stdev;
      const likelihood = cell.likelihood;

      cell.samples = new Array<number>();

      let normalSamples = new Array<number>();
      const values = <number[]>range(0, 1, 0.01).toArray(); // for 100 samples

      values.forEach((val: number) => {
        normalSamples.push(jStat.normal.inv(val, mean, stdev));
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
      let oldCell = null;
      let count = 0;
      if (this.oldCells != null) {
        oldCell = this.oldCells.find((oldCell: CellProperties) => oldCell.id == cell.id)
      }

      cell.inputCells.forEach((inCell: CellProperties) => {

        let oldInCell = null;

        if (this.oldCells != null) {
          oldInCell = this.oldCells.find((oldCell: CellProperties) => oldCell.id == inCell.id)
        }

        this.addSamplesToCell(inCell, oldInCell);

        if (this.oldCells != null) {
          if (inCell.samples == oldInCell.samples) {
            count++;
          }
        }
      })

      if (oldCell != null) {

        if (count == cell.inputCells.length) {
          cell.samples = oldCell.samples;
          return cell.samples;
        }
      }

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
        this.cells[i].stdev = 0;

        if (this.cells[i].isUncertain) {

          this.cells[i].stdev = this.cells[i + 1].value;
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