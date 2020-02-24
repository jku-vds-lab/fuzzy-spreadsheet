// /* global console, Excel */
// import * as jstat from 'jstat';
// export default class Spread {
//   chartType: string;

//   private addVarianceInfo() {

//     // use the info of uncertain cells
//     for (let i = 0; i < this.cells.length; i++) {
//       for (let r = 5; r < 22; r++) {
//         let id = "R" + r + "C8";
//         if (this.cells[i].id == id) {
//           this.cells[i].variance = this.cells[i + 2].value;
//           console.log('Variance:' + this.cells[i].variance);
//         }
//       }
//     }
//   }

//   async createNormalDistributions() {

//     this.addVarianceInfo();
//     await Excel.run(async (context) => {

//       let cheatsheet = context.workbook.worksheets.getItemOrNullObject("CheatSheet");
//       await context.sync();

//       if (!cheatsheet.isNullObject) {
//         cheatsheet.delete();
//       }

//       cheatsheet = context.workbook.worksheets.add("CheatSheet");
//       let rowIndex = -1;
//       // let min = mean - variance * 2;
//       // let max = mean + variance * 2;

//       for (let c = 0; c < this.cells.length; c++) {

//         this.cells[c].samples = new Array<number>();


//         let overallMin = -10;
//         let overallMax = 40;
//         let mean = this.cells[c].value;


//         let variance = this.cells[c].variance

//         if (variance > 0) {
//           rowIndex++;
//           let sampleSize = (variance * 2) / 50;

//           for (let i = overallMin; i <= overallMax; i = i + sampleSize) {
//             this.cells[c].samples.push(jstat.normal.pdf(i, mean, variance));
//             this.cells[c].isLineChart = true;
//           }
//         }
//         else {
//           rowIndex++;
//           if (this.cells[c].degreeToFocus >= 0) {
//             for (let i = overallMin; i <= overallMax; i++) {
//               if (i == ceil(this.cells[c].value)) {
//                 this.cells[c].samples.push(1);
//               } else {
//                 this.cells[c].samples.push(0);
//               }
//             }
//           }
//         }

//         if (this.cells[c].samples.length == 0) {
//           continue;
//         }

//         let range = cheatsheet.getRangeByIndexes(rowIndex, 0, 1, this.cells[c].samples.length);
//         range.values = [this.cells[c].samples];
//         range.load('address');
//         await context.sync();
//         this.cells[c].spreadRange = range.address;
//       }

//       await context.sync();
//     });
//   }

//   async addSpread(focusCell: CellProperties) {

//     await this.createNormalDistributions();

//     this.drawLineChart(focusCell);
//     // this.drawCompleteLineChart(focusCell);

//     focusCell.inputCells.forEach((cell: CellProperties) => {
//       this.drawLineChart(cell);

//     })

//     focusCell.outputCells.forEach((cell: CellProperties) => {
//       this.drawLineChart(cell);
//     })
//   }

//   private drawLineChart(cell: CellProperties) {

//     if (cell.spreadRange == null) {
//       return;
//     }
//     try {
//       Excel.run((context) => {

//         const sheet = context.workbook.worksheets.getActiveWorksheet();
//         const cheatSheet = context.workbook.worksheets.getItem("CheatSheet");
//         const dataRange = cheatSheet.getRange(cell.spreadRange);

//         let chart: Excel.Chart;

//         if (cell.isLineChart) {
//           console.log("Line chart");
//           chart = sheet.charts.add(Excel.ChartType.line, dataRange, Excel.ChartSeriesBy.rows);
//         } else {
//           console.log("Column chart");
//           chart = sheet.charts.add(Excel.ChartType.columnClustered, dataRange, Excel.ChartSeriesBy.rows);
//         }

//         chart.setPosition(cell.address, cell.address);
//         chart.left = cell.left + 0.2 * cell.width;
//         chart.title.visible = false;
//         chart.legend.visible = false;
//         chart.axes.valueAxis.minimum = 0;
//         // chart.axes.valueAxis.maximum = 0.21;
//         chart.dataLabels.showValue = false;
//         chart.axes.valueAxis.visible = false;
//         chart.axes.categoryAxis.visible = false;
//         chart.axes.valueAxis.majorGridlines.visible = false;
//         chart.plotArea.top = 0;
//         chart.plotArea.left = 0;
//         chart.plotArea.width = cell.width - 0.4 * cell.width;
//         chart.plotArea.height = 100;
//         chart.format.fill.clear();
//         chart.format.border.clear();
//         return context.sync();
//       });
//     } catch (error) {
//       console.log('Could not draw chart because of the following error', error);
//     }
//   }

// }