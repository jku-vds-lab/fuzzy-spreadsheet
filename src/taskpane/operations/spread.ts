/* global console, Excel */
import { ceil } from 'mathjs';
import * as jstat from 'jstat';
import CellProperties from '../cellproperties';
import SheetProperties from '../sheetproperties';

export default class Spread {
  private chartType: string;
  private cells: CellProperties[];
  private referenceCell: CellProperties;
  private sheetName: string;
  private rangeAddresses: Excel.Range[];

  constructor(cells: CellProperties[], referenceCell: CellProperties, sheetName: string = 'CheatSheet') {
    this.cells = cells;
    this.referenceCell = referenceCell;
    this.sheetName = sheetName;
    this.rangeAddresses = new Array<Excel.Range>();
  }

  public async showSpread(n: number) {

    await this.addSpreadInfoToCells();

    this.drawLineChart(this.referenceCell);

    this.showInputSpread(this.referenceCell.inputCells, n);
    this.showOutputSpread(this.referenceCell.outputCells, n);
  }

  public async removeSpread() {

    this.cells.forEach((cell: CellProperties) => {
      cell.isSpread = false;
    })

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      var charts = sheet.charts;
      charts.load("items/$none");
      return context.sync().then(function () {
        charts.items.forEach(function (chart) {
          chart.delete();
        });
        return context.sync();
      });
    });
  }

  public async addSpreadInfoToCells() {

    try {

      this.addVarianceInfo();

      let values = new Array<Array<number>>();

      this.cells.forEach((cell: CellProperties) => {
        if (isNaN(cell.value)) {
          return;
        }
        this.addSamplesToCell(cell);

        let sampleValues = new Array<number>();
        let sampleLikelihood = new Array<number>();

        cell.mySamples.forEach((sample: { value: number, likelihood: number }) => {
          sampleValues.push(sample.value);
          sampleLikelihood.push(sample.likelihood);
        })

        values.push(sampleValues);
        values.push(sampleLikelihood);
      })

      await this.addValuesToSheet(values);

      let index = 0;

      this.cells.forEach((cell: CellProperties) => {
        if (isNaN(cell.value)) {
          return;
        }

        if (this.rangeAddresses[index] == null) {
          console.log('Returning for cell: ' + cell.address + 'because range is null');
        }

        cell.spreadRange = this.rangeAddresses[index].address;
        index++;
      })

    } catch (error) {
      console.log(error);
    }
  }

  public async createNewSheet(isDeleteSheet: boolean = false) {

    try {
      let isCreateNewSheet = true;

      await Excel.run(async (context) => {

        let cheatsheet = context.workbook.worksheets.getItemOrNullObject(this.sheetName);
        await context.sync();


        if (!cheatsheet.isNullObject) {

          isCreateNewSheet = false;

          if (isDeleteSheet) {
            cheatsheet.delete();
            isCreateNewSheet = true;
          }
        }

        if (isCreateNewSheet) {

          cheatsheet = context.workbook.worksheets.add(this.sheetName);
        }

        await context.sync();
      });
    } catch (error) {
      console.log(error);
    }
  }

  public addSamplesToCell(cell: CellProperties) {

    cell.mySamples = new Array<{ value: number, likelihood: number }>();


    const mean = cell.value;
    const variance = cell.variance;

    if (variance == 0) {
      cell.mySamples.push({ value: mean, likelihood: 1 });
    }
    else {

      if (cell.formula.includes('SUM')) {
        let resultSamples = this.addSamplesToSumCell(cell);

        cell.mySamples = resultSamples;

        return;
      }

      if (cell.formula.includes('-')) {
        let resultSamples = this.addSamplesToSumCell(cell, true);

        cell.mySamples = resultSamples;

        return;
      }

      cell.mySamples.push({ value: 0, likelihood: (1 - cell.likelihood) });

      let numberOfSamples = 0;
      let i = mean - variance;

      while (i <= mean + variance) {
        numberOfSamples++;
        i++;
      }

      for (let i = mean - variance; i <= mean + variance; i++) {

        cell.mySamples.push({ value: i, likelihood: (cell.likelihood / numberOfSamples) });
        cell.isLineChart = true;
      }
    }
  }

  public addSamplesToSumCell(cell: CellProperties, isDifference: boolean = false) {

    let inputCells = cell.inputCells;
    let index = 0;

    let resultantSample = new CellProperties();
    resultantSample.mySamples = new Array<{ value: number, likelihood: number }>();

    if (inputCells.length > 1) {
      resultantSample = this.addTwoSamples(inputCells[index], inputCells[index + 1], isDifference);
      index = index + 2;
    }

    while (index < inputCells.length) {

      resultantSample = this.addTwoSamples(resultantSample, inputCells[index], isDifference);
      index = index + 1;
    }

    resultantSample.mySamples.forEach((sample: { value: number, likelihood: number }) => {
      cell.mySamples.push(sample);
    })

    return cell.mySamples;
  }

  private addTwoSamples(sample1: CellProperties, sample2: CellProperties, isDifference: boolean = false) {

    let resultantSample = new CellProperties();
    resultantSample.mySamples = new Array<{ value: number, likelihood: number }>();

    sample1.mySamples.forEach((sampleCell1: { value: number, likelihood: number }) => {

      sample2.mySamples.forEach((sampleCell2: { value: number, likelihood: number }) => {

        let value = sampleCell1.value + sampleCell2.value;

        if (isDifference) {
          value = sampleCell1.value - sampleCell2.value;
        }
        const likelihood = sampleCell1.likelihood * sampleCell2.likelihood;

        let allowInsert = true;

        resultantSample.mySamples.forEach((result: { value: number, likelihood: number }) => {
          if (result.value == value) {
            result.likelihood += likelihood;
            allowInsert = false;
            return;
          }
        })

        if (allowInsert) {
          resultantSample.mySamples.push({ value: value, likelihood: likelihood });
        }

      })
    })

    return resultantSample;
  }

  public async addValuesToSheet(values: any[][]) {


    await Excel.run(async (context) => {
      const cheatsheet = context.workbook.worksheets.getItem(this.sheetName);


      values.forEach((value: number[], index: number) => {
        let range = cheatsheet.getRangeByIndexes(index, 0, 1, value.length);
        range.values = [value];

      })

      for (let index = 0; index < values.length; index = index + 2) {

        let range = cheatsheet.getRangeByIndexes(index, 0, 2, values[index].length);
        this.rangeAddresses.push(range.load('address'));
      }

      await context.sync();
    })

    return this.rangeAddresses;
  }

  private showInputSpread(cells: CellProperties[], i: number) {

    cells.forEach((cell: CellProperties) => {

      if (cell.isSpread) {
        console.log(cell.address + ' already has a spread');
        return;
      }

      cell.isSpread = true;
      this.drawLineChart(cell);

      if (i == 1) {
        return;
      }
      this.showInputSpread(cell.inputCells, i - 1);
    })
  }

  private showOutputSpread(cells: CellProperties[], i: number) {

    cells.forEach((cell: CellProperties) => {

      if (cell.isSpread) {
        return;
      }

      cell.isSpread = true;

      this.drawLineChart(cell);
      if (i == 1) {
        return;
      }
      this.showOutputSpread(cell.outputCells, i - 1);
    })
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

  public drawLineChart(cell: CellProperties, color: string = null, lineWeight: number = 2, chartName: string = 'Chart') {

    if (cell.spreadRange == null) {
      console.log('Returning because spreadrange is null');
      return;
    }

    try {
      Excel.run((context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const cheatSheet = context.workbook.worksheets.getItem(this.sheetName);
        const dataRange = cheatSheet.getRange(cell.spreadRange);

        let chart: Excel.Chart;

        chart = sheet.charts.add(Excel.ChartType.columnClustered, dataRange, Excel.ChartSeriesBy.rows);

        chart.setPosition(cell.address, cell.address);
        // only if chatt type is line, if it is column, use the fill
        if (color != null) {
          chart.series.getItemAt(0).format.line.color = color;
        }

        chart.name = chartName;
        chart.series.getItemAt(0).format.line.weight = lineWeight;
        chart.left = cell.left + 0.2 * cell.width;
        chart.title.visible = false;
        chart.legend.visible = false;
        chart.axes.valueAxis.minimum = 0;
        // chart.axes.valueAxis.maximum = 0.21;
        chart.dataLabels.showValue = false;
        chart.axes.valueAxis.visible = false;
        chart.axes.categoryAxis.visible = false;
        chart.axes.valueAxis.majorGridlines.visible = false;
        chart.plotArea.top = 0;
        chart.plotArea.left = 0;
        chart.plotArea.width = cell.width - 0.4 * cell.width;
        chart.plotArea.height = 100;
        chart.format.fill.clear();
        chart.format.border.clear();
        chart.series.getItemAt(0).delete(); // labels are not right
        return context.sync().then(() => {
          console.log('Finished drawing the chart')
        }).
          catch((reason: any) => console.log('Failed to draw a chart: ' + reason));
      });
    } catch (error) {
      console.log('Could not draw chart because of the following error', error);
    }
  }
}