/* global console, Excel */
import { ceil } from 'mathjs';
import * as jstat from 'jstat';
import CellProperties from '../cellproperties';
import SheetProperties from '../sheetproperties';
export default class Spread {
  private chartType: string;
  private cells: CellProperties[];
  private referenceCell: CellProperties;

  constructor(cells: CellProperties[], referenceCell: CellProperties) {
    this.cells = cells;
    this.referenceCell = referenceCell;
  }

  public showSpread(n: number) {

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

  public async createCheatSheet() {

    this.addVarianceInfo();
    await Excel.run(async (context) => {

      let cheatsheet = context.workbook.worksheets.getItemOrNullObject("CheatSheet");
      await context.sync();

      if (!cheatsheet.isNullObject) {
        cheatsheet.delete();
      }

      SheetProperties.isCheatSheetExist = true;

      cheatsheet = context.workbook.worksheets.add("CheatSheet");
      let rowIndex = -1;
      // let min = mean - variance * 2;
      // let max = mean + variance * 2;

      for (let c = 0; c < this.cells.length; c++) {

        if (isNaN(this.cells[c].value)) {
          continue;
        }

        this.cells[c].samples = new Array<number>();

        let overallMin = -15;
        let overallMax = 40;
        let mean = this.cells[c].value;

        let variance = this.cells[c].variance

        if (variance > 0) {
          rowIndex++;
          let sampleSize = (variance * 2) / 50;

          for (let i = overallMin; i <= overallMax; i = i + sampleSize) {
            this.cells[c].samples.push(jstat.normal.pdf(i, mean, variance));
            this.cells[c].isLineChart = true;
          }
        }
        else {
          rowIndex++;
          for (let i = overallMin; i <= overallMax; i++) {
            if (i == ceil(this.cells[c].value)) {
              this.cells[c].samples.push(1);
            } else {
              this.cells[c].samples.push(0);
            }
          }
        }

        let range = cheatsheet.getRangeByIndexes(rowIndex, 0, 1, this.cells[c].samples.length);
        range.values = [this.cells[c].samples];
        range.load('address');
        await context.sync();
        this.cells[c].spreadRange = range.address;
      }

      await context.sync();
    });
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

  private addVarianceInfo() {

    for (let i = 0; i < this.cells.length; i++) {
      this.cells[i].variance = 0;
      if (this.cells[i].isUncertain) {
        this.cells[i].variance = this.cells[i + 1].value;
      }
    }
  }

  private drawLineChart(cell: CellProperties) {

    if (cell.spreadRange == null) {
      return;
    }

    try {
      Excel.run((context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const cheatSheet = context.workbook.worksheets.getItem("CheatSheet");
        const dataRange = cheatSheet.getRange(cell.spreadRange);

        let chart: Excel.Chart;

        if (cell.isLineChart) {
          chart = sheet.charts.add(Excel.ChartType.line, dataRange, Excel.ChartSeriesBy.rows);
        } else {
          chart = sheet.charts.add(Excel.ChartType.columnClustered, dataRange, Excel.ChartSeriesBy.rows);
        }

        chart.setPosition(cell.address, cell.address);
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
          catch(() => console.log('Failed to draw a chart'));
      });
    } catch (error) {
      console.log('Could not draw chart because of the following error', error);
    }
  }
}