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

    this.addVarianceInfo();

    let values = new Array<Array<number>>();

    this.cells.forEach((cell: CellProperties) => {
      if (isNaN(cell.value)) {
        return;
      }
      this.addSamplesToCell(cell);
      values.push(cell.samples);
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
  }

  public async createNewSheet(isDeleteSheet: boolean = false) {

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
  }

  public addSamplesToCell(cell: CellProperties) {

    let overallMin = -15;
    let overallMax = 40;

    cell.samples = new Array<number>();
    const mean = cell.value;
    const variance = cell.variance;

    if (variance == 0) {

      for (let i = overallMin; i <= overallMax; i++) {
        if (i == ceil(cell.value)) {
          cell.samples.push(1);
        } else {
          cell.samples.push(0);
        }
      }
    } else {
      const sampleSize = (variance * 2) / 50;

      for (let i = overallMin; i <= overallMax; i = i + sampleSize) {
        cell.samples.push(jstat.normal.pdf(i, mean, variance));
        cell.isLineChart = true;
      }
    }
  }

  public async addValuesToSheet(values: any[][]) {

    await Excel.run(async (context) => {
      const cheatsheet = context.workbook.worksheets.getItem(this.sheetName);
      let rowIndex = 0;

      values.forEach((value: number[]) => {
        let range = cheatsheet.getRangeByIndexes(rowIndex, 0, 1, value.length);
        range.values = [value];
        rowIndex++;

        this.rangeAddresses.push(range.load('address'));
      })

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

        if (cell.isLineChart) {
          chart = sheet.charts.add(Excel.ChartType.line, dataRange, Excel.ChartSeriesBy.rows);
        } else {
          chart = sheet.charts.add(Excel.ChartType.columnClustered, dataRange, Excel.ChartSeriesBy.rows);
        }

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
        return context.sync().then(() => { console.log('Finished drawing the chart') }).
          catch((reason: any) => console.log('Failed to draw a chart: ' + reason));
      });
    } catch (error) {
      console.log('Could not draw chart because of the following error', error);
    }
  }
}