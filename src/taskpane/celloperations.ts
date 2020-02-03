/* global console, Excel */
import CellProperties from './cellproperties';
import ShapeProperties from './shapeproperties';

export default class CellOperations {
  chartType: string;
  cells: CellProperties[];
  shapes: ShapeProperties[];

  CellOperations() {
    this.cells = new Array<CellProperties>();
  }

  getShapes() {
    return this.shapes;
  }

  setCells(cells: CellProperties[]) {
    this.cells = cells;
  }


  addSpreadInfo() {
    Excel.run(async (context) => {
      const cheatsheet = context.workbook.worksheets.add("CheatSheet");
      let data: number[][] = new Array<Array<number>>();
      let means = [32, 13, 7, 12, 26.6, 0.6, 1, 9, 9, 7]; // make it dynamic
      let stdDev = [6.38, 2.5, 2.9, 1.8, 4.8, 0.2, 0.4, 2.7, 2.2, 1.34]; // make it dynamic
      for (let i = 0; i < 47; i++) {
        let row = new Array<number>();
        for (let j = 0; j < 10; j++) {
          var normalVal = context.workbook.functions.norm_Dist(i + 1, means[j], stdDev[j], false);
          normalVal.load("value");
          await context.sync();
          row.push(normalVal.value);
        }
        data.push(row);
      }
      var range = cheatsheet.getRange("A1:J47");
      range.values = data;
      await context.sync();
    });
  }

  private addImpactInfo(focusCell: CellProperties) {

    let height = 5;
    let width = 5;

    this.shapes = new Array<ShapeProperties>();

    focusCell.inputCells.forEach((inCell: CellProperties) => {

      let prop = this.calculateInTransparency(inCell.value, focusCell.value, focusCell.inputCells);

      this.shapes.push(
        new ShapeProperties().setShapeProperties(inCell, Excel.GeometricShapeType.rectangle, prop.color, prop.transparency, height, width)
      );
    })

    focusCell.outputCells.forEach((outCell: CellProperties) => {

      let prop = this.calculateOutTransparency(outCell.value, focusCell.value, outCell.inputCells);

      this.shapes.push(
        new ShapeProperties().setShapeProperties(outCell, Excel.GeometricShapeType.rectangle, prop.color, prop.transparency, height, width)
      );
    })
  }

  async addImpact(focusCell: CellProperties) {
    this.addImpactInfo(focusCell);

    try {
      // Ensure cells and shapes are the same length
      await Excel.run(async (context) => {

        const sheet = context.workbook.worksheets.getItem("Probability");
        for (let i = 0; i < this.shapes.length; i++) {
          var impact = sheet.shapes.addGeometricShape("Rectangle"); // shapes[i].shapeType
          impact.name = "Impact" + i;
          console.log("Impact object: " + impact.name);
          impact.height = this.shapes[i].height;
          impact.width = this.shapes[i].width;
          impact.left = this.shapes[i].cell.left + 2;
          impact.top = this.shapes[i].cell.top + this.shapes[i].cell.height / 4;
          impact.rotation = 0;
          impact.fill.transparency = this.shapes[i].transparency;
          impact.lineFormat.weight = 0;
          impact.lineFormat.color = this.shapes[i].color;
          impact.fill.setSolidColor(this.shapes[i].color);
        }
        // createImpactLegend().then(function () { });
        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  }

  private calculateInTransparency(cellValue: number, focusCellValue: number, cells: CellProperties[]) {

    let color = "green";
    let transparency = 0;

    if (focusCellValue > 0 && cellValue < 0) {
      color = "red";
    }

    if (focusCellValue < 0 && cellValue < 0 && cellValue < focusCellValue) { // because of the negative sign, the smaller the number the higher it is
      color = "red";
    }

    // Make both values positive
    if (cellValue < 0) {
      cellValue = -cellValue;
    }

    if (focusCellValue < 0) {
      focusCellValue = -focusCellValue;
    }

    if (cellValue < focusCellValue) {

      let value = cellValue / focusCellValue;

      transparency = 1 - value;
    }
    else {

      let maxValue = cellValue;
      // go through the input cells of the focus cell
      cells.forEach((c: CellProperties) => {
        let val = c.value;

        if (val < 0) {
          val = -val;
        }
        if (val > maxValue) {
          maxValue = val;
        }
      })

      transparency = 1 - (cellValue / maxValue);
    }

    return { color: color, transparency: transparency };
  }

  // Fix color values for negative values
  private calculateOutTransparency(cellValue: number, focusCellValue: number, cells: CellProperties[]) {

    let color = "green";
    let transparency = 0;

    if (focusCellValue > 0 && cellValue < 0) {
      color = "red";
    }

    if (focusCellValue < 0 && cellValue < 0 && cellValue < focusCellValue) { // because of the negative sign, the smaller the number the higher it is
      color = "red";
    }

    if (focusCellValue < 0 && cellValue > 0) { // because of the negative sign, the smaller the number the higher it is
      color = "red";
    }

    // Make both values positive
    if (cellValue < 0) {
      cellValue = -cellValue;
    }

    if (focusCellValue < 0) {
      focusCellValue = -focusCellValue;
    }

    if (cellValue > focusCellValue) {

      let value = focusCellValue / cellValue;

      transparency = 1 - value;

    }
    else {
      let maxValue = cellValue;
      // go through the input cells of the output cell
      cells.forEach((c: CellProperties) => {
        let val = c.value;
        if (val < 0) {
          val = -val;
        }
        if (val > maxValue) {
          maxValue = val;
        }
      })

      transparency = 1 - (focusCellValue / maxValue);
    }

    return { color: color, transparency: transparency };
  }

  addLikelihoodInfo() {
    for (let i = 0; i < this.cells.length; i++) {
      for (let r = 5; r < 18; r++) {
        let id = "R" + r + "C8";
        if (this.cells[i].id == id) {
          this.cells[i].likelihood = this.cells[i + 1].value;
        }
      }
    }

    if (this.shapes.length > 0) {
      for (let i = 0; i < this.shapes.length; i++) {
        this.shapes[i].height = this.shapes[i].cell.likelihood / 10;
        this.shapes[i].width = this.shapes[i].cell.likelihood / 10;
      }
    }
  }

  // // Not possible without impact yet
  async addLikelihood() {
    this.addLikelihoodInfo();

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");

      for (let i = 0; i < this.shapes.length; i++) {
        var shape = sheet.shapes.getItem("Impact" + i);
        shape.load(["height", "width"]);
        await context.sync();

        console.log("Rectangle Found");
        shape.height = this.shapes[i].height;
        shape.width = this.shapes[i].width;

      }
      // createLikelihoodLegend().then(function () { });
      await context.sync();
    });
  }
}