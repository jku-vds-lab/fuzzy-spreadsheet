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


// the original file should not contain the variance and likelihood inforamtion at all, so adapt accordingly
export default class Spread {

  private cells: CellProperties[];
  private referenceCell: CellProperties;
  private color: string = null;

  constructor(cells: CellProperties[], referenceCell: CellProperties, color: string = null) {
    this.cells = cells;
    this.referenceCell = referenceCell;
    this.color = color;
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
        return;
      }
      this.referenceCell.isSpread = true;
      this.addSamplesToCell(this.referenceCell);
      this.drawBarCodePlot(this.referenceCell, 'ReferenceChart');

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
        this.addSamplesToCell(cell);
        this.drawBarCodePlot(cell, 'InputChart');

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
        this.addSamplesToCell(cell);
        this.drawBarCodePlot(cell, 'OutputChart');
        if (i == 1) {
          return;
        }
        this.showOutputSpread(cell.outputCells, i - 1);
      })

    } catch (error) {
      console.log('Error in Show Output Cell Spread', error);
    }
  }

  public drawBarCodePlot(cell: CellProperties, name: string) {
    try {
      Excel.run((context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let totalHeight = cell.height;

        let startLineTop = cell.top;
        let startLineLeft = cell.left + 20;
        let endLineTop = cell.top + totalHeight;

        let colors = ['#d8e1e7', '#98b0c2', '#4e7387', '#002e41', '#002534'] // light to dark

        var count = 5;
        let data = cell.samples;

        let domain = d3.max(data, function (d) { return +d })

        var x = d3.scaleLinear().domain([0, domain]).nice(count);

        var histogram = d3.histogram().value(function (d) { return d }).domain([0, domain]).thresholds(x.ticks(count));
        var bins = histogram(data);

        console.log('Bins for ' + cell.address);

        let sortedBins = bins.sort((n1, n2) => {
          if (n1.length > n2.length) {
            return 1;
          }

          if (n1.length < n2.length) {
            return -1;
          }
          return 0;
        })

        let colorIndex = 0;

        sortedBins.forEach((bin) => {

          const bin_zero = bin[0];

          if (bin_zero == undefined) {
            return;
          }

          let line = sheet.shapes.addLine(startLineLeft + bin_zero, startLineTop, startLineLeft + bin_zero, endLineTop);
          line.lineFormat.color = colors[colorIndex];
          line.name = name;
          line.lineFormat.weight = 3;
          line.lineFormat.transparency = 0.5;
          colorIndex++;
        })

        console.log('-------------------------------');

        return context.sync().then(() => {
          console.log('Finished drawing the bar code plot')
        }).
          catch((reason: any) => console.log('Failed to draw the bar code plot: ' + reason));
      });
    } catch (error) {
      console.log('Could not draw the bar code plot because of the following error', error);
    }

  }

  public addSamplesToCell(cell: CellProperties) {

    try {

      cell.samples = new Array<number>();

      const mean = cell.value;
      const variance = cell.variance;

      if (variance == 0) {
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

      cell.samples = new Array<number>();

      let normalSamples = new Array<number>();
      const values = <number[]>range(0, 1, 0.01).toArray(); // for 100 samples

      values.forEach((val: number) => {
        normalSamples.push(jStat.normal.inv(val, mean, variance));
      })

      normalSamples = normalSamples.filter(outliers());

      const sampleLength = normalSamples.length;
      console.log('SampleLength: ' + sampleLength);

      const bern = Bernoulli(cell.likelihood);
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

        if (inCell.samples == null) {
          this.addSamplesToCell(inCell);
        }
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

      console.log('Length: ' + resultantSample.samples.length);
      console.log(resultantSample.samples);
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
          return context.sync();
        }).catch((reason: any) => {
          console.log('Step 1:', reason, name)
        });
      });
    } catch (error) {
      console.log('Step 2:', error);
    }
  }
}