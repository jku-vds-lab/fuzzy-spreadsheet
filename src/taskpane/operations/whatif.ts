import CellProperties from "../cellproperties";
import Spread from "./spread";

/* global console, Excel */
export default class WhatIf {
  public value: number = 0;
  public variance: number = 0;
  public likelihood: number = 0;
  public spreadRange: string = null;
  private newCells: CellProperties[];

  setNewCells(newCells) {
    this.newCells = newCells;
  }

  async calculateUpdatedNumber(referenceCell: CellProperties) {

    try {
      this.newCells.forEach(async (newCell: CellProperties, index: number) => {

        if (referenceCell.id == newCell.id) {
          referenceCell.whatIf = new WhatIf();
          referenceCell.whatIf.value = newCell.value - referenceCell.value;
          referenceCell.whatIf.variance = this.newCells[index + 1].value - referenceCell.variance;
          return;
        }
      })

    } catch (error) {
      console.log('Error: ', error);
    }
  }

  // check the variance
  async drawChangedSpread(referenceCell: CellProperties, oldVariance: number) {

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      const charts = sheet.charts;
      charts.load("items/name");
      await context.sync().then(function () {
        charts.items.forEach(function (c) {
          if (c.name == 'UpdateChart')
            c.delete();
          console.log('Deleted old chart');
        });
      });

    });

    const spread = new Spread(this.newCells, referenceCell, 'MyCheatSheet');
    await spread.createNewSheet(true);
    spread.addVarianceInfo();

    this.newCells.forEach(async (newCell: CellProperties) => {
      if (newCell.id == referenceCell.id) {
        console.log('Reference cell variance: ' + newCell.variance + ' and oldVariance: ' + oldVariance);
        spread.addSamplesToCell(newCell);
        const rangeAddress = await spread.addValuesToSheet([newCell.samples]);
        // eslint-disable-next-line require-atomic-updates
        newCell.spreadRange = rangeAddress[0].address; // should be in whatif?
        spread.drawLineChart(newCell, 'orange', 1, 'UpdateChart');
        return;
      }
    })
  }
}