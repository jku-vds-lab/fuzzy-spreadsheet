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

    this.newCells.forEach(async (newCell: CellProperties) => {

      if (referenceCell.id == newCell.id) {
        referenceCell.whatIf = new WhatIf();
        referenceCell.whatIf.value = newCell.value - referenceCell.value;
        return;
      }
    })
  }

  // check the variance
  async drawChangedSpread(referenceCell: CellProperties) {
    // const spread = new Spread(this.newCells, referenceCell, 'MyCheatSheet');
    // await spread.createCheatSheet();
    // spread.showSpread(1);
  }
}