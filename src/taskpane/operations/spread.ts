/* global console, Excel */
import CellProperties from '../cell/cellproperties';
import * as outliers from 'outliers';
import { range, dotMultiply, norm, multiply } from 'mathjs';
import { Bernoulli } from 'discrete-sampling';
import * as jStat from 'jstat';
import Bins from './bins';
import { sum, contours } from 'd3';
import { BaseLegendLayout } from 'vega';

export default class Spread {

  private referenceCell: CellProperties;
  private colors: string[];
  private blueColors: string[];
  private orangeColors: string[];

  private minDomain =  -2 // for demo use case
  private maxDomain =  28 // for demo use case
  private binWidth = (this.maxDomain - this.minDomain) / 15;
  private binsObj: Bins;
  private inputCellsWithSpread: CellProperties[];
  private outputCellsWithSpread: CellProperties[];

  constructor(referenceCell: CellProperties) {

    this.referenceCell = referenceCell;
    this.binsObj = new Bins(this.minDomain, this.maxDomain, this.binWidth);
    this.blueColors = this.binsObj.generateGreenColors();
    this.orangeColors = this.binsObj.generatePinkColors();
    this.colors = this.blueColors;
    this.inputCellsWithSpread = new Array<CellProperties>();
    this.outputCellsWithSpread = new Array<CellProperties>();
  }

  public showSpread(n: number, isInput: boolean, isOutput: boolean, isDraw: boolean = true) {

    try {

      this.showReferenceCellSpread();

      if (isInput && n > 0) {
        this.inputCellsWithSpread = new Array<CellProperties>();
        this.showInputSpread(this.referenceCell.inputCells, n);
        if (isDraw) {
          this.drawBarCodePlot(this.inputCellsWithSpread, 'InputSpread');
        }
      }

      if (isOutput && n > 0) {
        this.outputCellsWithSpread = new Array<CellProperties>();
        this.showOutputSpread(this.referenceCell.outputCells, n);
        if (isDraw) {
          this.drawBarCodePlot(this.outputCellsWithSpread, 'OutputSpread');
        }
      }

    } catch (error) {
      console.log('Error in Show spread', error);
    }
  }

  public getInputCellsWithSpread() {
    return this.inputCellsWithSpread;
  }

  public getOutputCellsWithSpread() {
    return this.outputCellsWithSpread;
  }

  public showReferenceCellSpread() {

    try {

      if (this.referenceCell.isSpread) {
        return;
      }

      this.referenceCell.isSpread = true;
      this.addSamplesToCell(this.referenceCell);
      this.drawBarCodePlot([this.referenceCell], 'ReferenceSpread');

    } catch (error) {
      console.log('Error in Show Reference Cell Spread', error);
    }
  }

  public showInputSpread(cells: CellProperties[], i: number) {

    try {
      cells.forEach((cell: CellProperties) => {

        if (!cell.isSpread) {

          cell.isSpread = true;

          if (cell.samples == null) {
            this.addSamplesToCell(cell);
          }

          this.inputCellsWithSpread.push(cell);
        }

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


        if (!cell.isSpread) {

          cell.isSpread = true;

          if (cell.samples == null) {
            this.addSamplesToCell(cell);
          }

          this.outputCellsWithSpread.push(cell);
        }

        if (i == 1) {
          return;
        }
        this.showOutputSpread(cell.outputCells, i - 1);
      })

    } catch (error) {
      console.log('Error in Show Output Cell Spread', error);
    }
  }

  public drawBarCodePlot(cells: CellProperties[], name: string, color: string = 'blue', isUpperHalf: boolean = false, isLowerHalf: boolean = false) {
    try {

      Excel.run((context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();

        cells.forEach((cell: CellProperties) => {

          let height = cell.height - 2;
          let top = cell.top + 1;
          let left = cell.left + 48;
          if (isUpperHalf) {
            height = height / 2;
          }

          if (isLowerHalf) {
            top = top + height / 2;
            height = height / 2;
          }

          if (color == 'orange') {
            this.colors = this.orangeColors;
          } else {
            this.colors = this.blueColors;
          }

          let sortedLinesWithColors = this.computeColorsAndBins(cell);

          if (color == 'orange') {
            cell.binOrangeColors = new Array<string>();
            sortedLinesWithColors.forEach((el) => {
              cell.binOrangeColors.push(el.color);
            })
          } else {
            cell.binBlueColors = new Array<string>();
            sortedLinesWithColors.forEach((el) => {
              cell.binBlueColors.push(el.color);
            })
          }

          let i = 0;
          sortedLinesWithColors.forEach((el) => {
            let rect = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
            rect.name = cell.address + name;
            rect.top = top;
            // rect.left = left + el.value * 0.8 - 2;

            rect.width = 2;

            rect.left = left + rect.width * i;
            i++;
            rect.height = height;
            rect.fill.setSolidColor(el.color);
            // rect.fill.transparency = 0.5;
            rect.lineFormat.visible = false;
          })
        })
        let range = sheet.getRange(this.referenceCell.address);
        range.select();
        return context.sync();
      });
    } catch (error) {
      console.log('Could not draw the bar code plot because of the following error', error);
    }
  }

  private computeColorsAndBins(cell: CellProperties) {

    let sortedLinesWithColors = new Array<{ value: number, color: string, freq: number }>();

    try {

      let data = cell.samples;

      let bins = this.binsObj.createBins(data);

      bins.forEach((bin) => {
        let binValue = bin.x0;
        let binFreq = bin.length;
        let binColorIndex = Math.ceil((binFreq / data.length) * bins.length);
        let binColor = this.colors[binColorIndex];

        const element = { value: binValue, color: binColor, freq: binFreq };
        sortedLinesWithColors.push(element);
      })
    } catch (error) {
      console.log(error);
    }
    return sortedLinesWithColors;
  }

  private addSamplesToCell(cell: CellProperties) {

    try {

      cell.samples = new Array<number>();

      const mean = cell.value;
      const stdev = cell.stdev;
      const likelihood = cell.likelihood;

      if (cell.formula.includes('SUM')) {
        cell.samples = this.addSamplesToSumCell(cell);
      }

      if (cell.formula.includes('-')) {
        cell.samples = this.addSamplesToSumCell(cell, true);
      }

      if (cell.formula.includes('*')) {
        cell.samples = this.addSamplesToSumCell(cell, false, true);
      }

      if (cell.formula.includes('AVERAGE')) {
        cell.samples = this.addSamplesToAverageCell(cell);
      }

      // fix values for certain cells
      if (cell.formula == "") {
        if (stdev == 0 && likelihood == 1) {
          let i = 0;
          while (i < 95) {
            cell.samples.push(mean);
            i++;
          }
        }
      }

      cell.computedMean = jStat.mean(cell.samples);
      cell.computedStdDev = jStat.stdev(cell.samples);

    } catch (error) {
      console.log(error);
    }
  }

  private addSamplesToAverageCell(cell: CellProperties) {

    try {

      const mean = cell.value;
      const stdev = cell.stdev;
      const likelihood = cell.likelihood;


      cell.samples = new Array<number>();

      if (stdev === 0 && likelihood === 1) {

        console.log('For cell: ' +cell.address + ' we are just computing no distribution');

        let i = 0;
        while (i < 95) {
          cell.samples.push(mean);
          i++;
        }
      } else if(stdev === 0) {
        console.log('For cell: ' +cell.address + ' we are computing bernoulli samples');
        cell.samples = this.computeBernoulliSamples(mean, likelihood);
      } else if (likelihood == 1) {
        console.log('For cell: ' +cell.address + ' we are computing normal samples');
        cell.samples = this.computeNormalSamples(mean, stdev).normalSamples;
      } else {
        console.log('For cell: ' +cell.address + ' we are computing all samples');
        const normal = this.computeNormalSamples(mean, stdev);
        const normalSamples = normal.normalSamples;
        const sampleLength = normal.sampleLength;
        const bernoulliSamples = this.computeBernoulliSamples(likelihood, sampleLength);
        cell.samples = <number[]>dotMultiply(normalSamples, bernoulliSamples);
      }

    } catch (error) {
      console.log('Error in Average Spread Computation', error);
    }

    return cell.samples;
  }

  private computeNormalSamples(mean: number, stdev: number) {
    let normalSamples = new Array<number>();
    const values = <number[]>range(0, 1, 0.01).toArray(); // for 100 samples

    values.forEach((val: number) => {
      normalSamples.push(jStat.normal.inv(val, mean, stdev));
    })

    normalSamples = normalSamples.filter(outliers());
    const sampleLength = normalSamples.length;
    return { normalSamples: normalSamples, sampleLength: sampleLength };
  }

  private computeBernoulliSamples(mean: number = 1, likelihood: number = 1, sampleLength: number = 95) {
    const bern = Bernoulli(likelihood);
    bern.draw();

    const bernoulliSamples = <number[]>multiply(bern.sample(sampleLength), mean);
    return bernoulliSamples;
  }

  private addSamplesToSumCell(cell: CellProperties, isDifference: boolean = false, isMulti: boolean = false) {

    try {

      let inputCells = cell.inputCells;

      cell.inputCells.forEach((inCell: CellProperties) => {
        this.addSamplesToCell(inCell);
      })


      let index = 0;

      let resultantSample = new CellProperties();
      resultantSample.samples = new Array<number>();

      if (inputCells.length > 1) {
        resultantSample = this.addTwoSamples(inputCells[index], inputCells[index + 1], isDifference, isMulti);
        index = index + 2;
      }

      while (index < inputCells.length) {
        resultantSample = this.addTwoSamples(resultantSample, inputCells[index], isDifference, isMulti);
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

  private addTwoSamples(sample1: CellProperties, sample2: CellProperties, isDifference: boolean = false, isMulti: boolean = false) {

    let resultantSample = new CellProperties();
    resultantSample.samples = new Array<number>();

    try {
      sample1.samples.forEach((sampleCell1: number, index: number) => {

        if (sample2.samples.length <= index) {
          return;
        }

        let value = 0;

        if (isDifference) {
          value = sampleCell1 - sample2.samples[index];
        } else if (isMulti) {
          value = sampleCell1 * sample2.samples[index];
        } else {
          value = sampleCell1 + sample2.samples[index];
        }

        resultantSample.samples.push(value);

      })
    } catch (error) {
      console.log(error);
    }

    return resultantSample;
  }
}