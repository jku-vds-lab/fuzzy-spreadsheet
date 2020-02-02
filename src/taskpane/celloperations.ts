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

  async addImpactInfo(cells: CellProperties[], focusCellAddress: string) {

    let x = new CellProperties();
    x.getNeighbouringCells(cells, focusCellAddress);
    console.log(cells);
    // use these cells for getting impact infor and other info

    // console.log(this.cells);
    // this.cells = cells;
    // let color = "green";
    // let transparency = 0;
    // let height = 5;
    // let width = 5;
    // let secondDegreeDivisor = -1;
    // this.shapes = new Array<ShapeProperties>();
    // // Finding the firstDegreeDivisor
    // this.cells.forEach((cell: CellProperties) => {
    //   let val = cell.value;
    //   if (val < 0) {
    //     val = -cell.value;
    //   }
    //   if (cell.degreeToFocus == 1 && val > firstDegreeDivisor) {
    //     firstDegreeDivisor = val;
    //   }
    // });
    // // Finding the secondDegreeDivisor & assigning shape properties
    // this.cells.forEach((cell: CellProperties) => {
    //   let val = cell.value;
    //   if (val < 0) {
    //     val = -cell.value;
    //   }
    //   if (cell.value < 0) {
    //     color = "red";
    //   }
    //   if (cell.degreeToFocus == 1) {
    //     secondDegreeDivisor = val;
    //     transparency = 1 - val / firstDegreeDivisor;
    //   } else if (cell.degreeToFocus == 2) {
    //     transparency = 1 - val / secondDegreeDivisor;
    //   }
    //   this.shapes.push(
    //     new ShapeProperties().getShapeProperties(Excel.GeometricShapeType.rectangle, color, transparency, height, width)
    //   );
    // });
  }

  // async function addImpact() {
  //   try {
  //     // Ensure cells and shapes are the same length
  //     await Excel.run(async (context) => {
  //       let dim = new CellOperations();
  //       let cellAddresses = ["I6", "I7", "I8", "I9", "I11", "I12", "I13", "I14",
  //         "I15", "I16"];
  //       await dim.scanRange(cellAddresses, "I18");
  //       let cells = dim.getCells();
  //       dim.addImpactInfo(cells);
  //       let shapes = dim.getShapes();
  //       const sheet = context.workbook.worksheets.getItem("Probability");
  //       for (let i = 0; i < cells.length; i++) {
  //         var impact = sheet.shapes.addGeometricShape("Rectangle"); // shapes[i].shapeType
  //         impact.name = "Impact" + i;
  //         impact.height = shapes[i].height;
  //         impact.width = shapes[i].width;
  //         impact.left = cells[i].left + 2;
  //         impact.top = cells[i].top + cells[i].height / 4;
  //         impact.rotation = 0;
  //         impact.fill.transparency = shapes[i].transparency;
  //         impact.lineFormat.weight = 0;
  //         impact.lineFormat.color = shapes[i].color;
  //         impact.fill.setSolidColor(shapes[i].color);
  //       }
  //       // createImpactLegend().then(function () { });
  //       await context.sync();
  //     });
  //   } catch (error) {
  //     console.error(error);
  //   }
  // }

  async addLikelihoodInfo(cells: CellProperties[], shapes: ShapeProperties[], likelihoodAddresses: string[]) {
    this.cells = cells;
    this.shapes = shapes;
    let likelihoodCell: number;
    if (this.shapes.length > 0) {
      if (this.shapes.length != likelihoodAddresses.length) {
        return;
      }
      for (let i = 0; i < this.shapes.length; i++) {
        likelihoodCell = 50; //await new CellProperties().getCellValue(likelihoodAddresses[i]);
        this.shapes[i].height = likelihoodCell / 10;
        this.shapes[i].width = likelihoodCell / 10;
      }
    }
  }

}