/* global console, Excel */
import { std } from 'mathjs';
import CellProperties from './cellproperties';
import CustomShape from './customshape';

export default class CellOperations {
  chartType: string;
  cells: CellProperties[];
  customShapes: CustomShape[];

  CellOperations() {
    this.cells = new Array<CellProperties>();
  }

  setCells(cells: CellProperties[]) {
    this.cells = cells;
  }


  // Creating cheat sheet
  async addCheatSheet() {
    await Excel.run(async (context) => {
      const cheatsheet = context.workbook.worksheets.add("CheatSheet");
      let data: number[][] = new Array<Array<number>>();
      let means = [32, 13, 7, 12, 26.6, 0.6, 1, 9, 9, 7, 5.4]; // make it dynamic
      let stdDev = [6.38, 2.5, 2.9, 1.8, 4.8, 0.2, 0.4, 2.7, 2.2, 1.34, 5.84]; // make it dynamic
      for (let i = 0; i < 47; i++) {
        let row = new Array<number>();
        for (let j = 0; j < 11; j++) {
          var normalVal = context.workbook.functions.norm_Dist(i + 1, means[j], stdDev[j], false);
          normalVal.load("value");
          await context.sync();
          row.push(normalVal.value);
        }
        data.push(row);
      }
      var range = cheatsheet.getRange("A1:K47");
      range.values = data;
      await context.sync();
    });
  }

  private addImpactInfo(focusCell: CellProperties) {

    this.customShapes = new Array<CustomShape>();

    if (focusCell.formula.includes("GEOMEAN")) {
      console.log("Compute normalized euclidean distance");
      focusCell.inputCells.forEach((inCell: CellProperties) => {

        let colorProperties = this.inputColorPropertiesMedian(inCell.value, focusCell.value, focusCell.inputCells);

        let customShape: CustomShape = { cell: inCell, shape: null, color: colorProperties.color, transparency: colorProperties.transparency }
        this.customShapes.push(customShape);
      })
    }

    if (focusCell.formula.includes("SUM") || focusCell.formula.includes('-')) {
      focusCell.inputCells.forEach((inCell: CellProperties) => {

        let colorProperties = this.inputColorProperties(inCell.value, focusCell.value, focusCell.inputCells);

        console.log("SUM: " + colorProperties.transparency);

        let customShape: CustomShape = { cell: inCell, shape: null, color: colorProperties.color, transparency: colorProperties.transparency }
        this.customShapes.push(customShape);
      })
    }

    focusCell.outputCells.forEach((outCell: CellProperties) => {

      let colorProperties = this.outputColorProperties(outCell.value, focusCell.value, outCell.inputCells);

      console.log("OUT: " + colorProperties.transparency);

      let customShape: CustomShape = { cell: outCell, shape: null, color: colorProperties.color, transparency: colorProperties.transparency }
      this.customShapes.push(customShape);
    })
  }

  async addImpact(focusCell: CellProperties) {
    this.addImpactInfo(focusCell);

    try {
      Excel.run(async (context) => {

        const sheet = context.workbook.worksheets.getItem("Probability");
        let i = 0;

        this.customShapes.forEach((customShape: CustomShape) => {
          customShape.shape = sheet.shapes.addGeometricShape("Rectangle");
          customShape.shape.name = "Impact" + i;
          customShape.shape.left = customShape.cell.left + 2;
          customShape.shape.top = customShape.cell.top + customShape.cell.height / 4;
          customShape.shape.height = 5;
          customShape.shape.width = 5;
          customShape.shape.geometricShapeType = Excel.GeometricShapeType.rectangle;
          customShape.shape.fill.setSolidColor(customShape.color);
          customShape.shape.fill.transparency = customShape.transparency;
          customShape.shape.lineFormat.weight = 0;
          customShape.shape.lineFormat.color = customShape.color;
          i++;
        })

        // createImpactLegend().then(function () { });
        await context.sync();
      });

    } catch (error) {
      console.error(error);
    }
  }

  private computeColor(cellValue: number, focusCellValue: number, cells: CellProperties[]) {

    let color = "green";


    if (focusCellValue > 0 && cellValue < 0) {
      color = "red";
    }

    if (focusCellValue < 0 && cellValue > 0) { // because of the negative sign, the smaller the number the higher it is
      color = "red";
    }

    if (focusCellValue < 0 && cellValue < 0) { // because of the negative sign, the smaller the number the higher it is
      let isAnyCellPositive = false;

      cells.forEach((cell: CellProperties) => {
        if (cell.value > 0) {
          isAnyCellPositive = true;
        }
      })

      if (isAnyCellPositive) { // if even one cell is positive, then it means that only that cell is contributing positively and rest all are contributing negatively
        color = "red";
      }
    }
    return color;
  }

  private inputColorProperties(cellValue: number, focusCellValue: number, cells: CellProperties[]) {

    let transparency = 0;
    const color = this.computeColor(cellValue, focusCellValue, cells);

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

  private inputColorPropertiesMedian(cellValue: number, focusCellValue: number, cells: CellProperties[]) {

    let transparency = 0;
    let values: number[] = new Array<number>();


    cells.forEach((cell: CellProperties) => {
      values.push(cell.value);
    });

    let stdDev = std(values);

    console.log(" Stddev: " + stdDev);

    transparency = (focusCellValue - cellValue) / (2 * stdDev);

    if (transparency < 0) {
      transparency = -transparency;
    }

    if (transparency > 1) {
      transparency = 1;
    }

    console.log(" Transparency: " + transparency);
    return { color: "green", transparency: transparency }
  }

  // Fix color values for negative values
  private outputColorProperties(cellValue: number, focusCellValue: number, cells: CellProperties[]) {

    let transparency = 0;
    const color = this.computeColor(cellValue, focusCellValue, cells);

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

  private addLikelihoodInfo() {

    for (let i = 0; i < this.cells.length; i++) {
      for (let r = 5; r < 18; r++) {
        let id = "R" + r + "C8";
        if (this.cells[i].id == id) {
          this.cells[i].likelihood = this.cells[i + 1].value;
        }
      }
    }
  }

  // // Not possible without impact yet
  async addLikelihood() {

    this.addLikelihoodInfo();

    await Excel.run(async (context) => {

      for (let i = 0; i < this.customShapes.length; i++) {

        const sheet = context.workbook.worksheets.getItem("Probability");

        let shape = sheet.shapes.getItem("Impact" + i);
        shape.load(["height", "width", "top"]);

        await context.sync();

        let likelihood = this.customShapes[i].cell.likelihood / 10;

        if (likelihood == 10) {
          likelihood = this.cells[15].height;
          shape.top = shape.top - 4;
        }
        shape.height = likelihood;
        shape.width = likelihood;
      }
      // createLikelihoodLegend().then(function () { });
      await context.sync();
    });
  }

  private addSpreadInfo() {


    // make it dynamic
    let ranges: string[] = [
      "A1:A47",
      "B1:B47",
      "C1:C47",
      "D1:D47",
      "E1:E47",
      "F1:F47",
      "G1:G47",
      "H1:H47",
      "I1:I47",
      "J1:J47",
      "K1:K47"
    ];

    let rangeIndex = 0;

    for (let i = 0; i < this.cells.length; i++) {
      for (let r = 5; r < 18; r++) {
        let id = "R" + r + "C8";
        if (this.cells[i].id == id) {

          this.cells[i].spreadRange = ranges[rangeIndex];
          rangeIndex++;
        }
      }
    }
  }

  addSpread(focusCell: CellProperties) {

    this.addSpreadInfo();
    this.drawLineChart(focusCell);

    focusCell.inputCells.forEach((cell: CellProperties) => {
      this.drawLineChart(cell);
    })

    focusCell.outputCells.forEach((cell: CellProperties) => {
      this.drawLineChart(cell);
    })
  }

  private drawLineChart(cell: CellProperties) {

    console.log(cell, cell.spreadRange);
    if (cell.spreadRange == "") {
      return;
    }

    Excel.run((context) => {

      const sheet = context.workbook.worksheets.getItem("Probability");
      const cheatSheet = context.workbook.worksheets.getItem("CheatSheet");
      const dataRange = cheatSheet.getRange(cell.spreadRange);
      let chart = sheet.charts.add("Line", dataRange, Excel.ChartSeriesBy.auto);

      chart.setPosition(cell.address, cell.address);
      chart.left = cell.left + 0.2 * cell.width;
      chart.title.visible = false;
      chart.legend.visible = false;
      chart.axes.valueAxis.minimum = 0;
      chart.axes.valueAxis.maximum = 0.21;
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
      return context.sync();
    });
  }

  addInArrows(focusCell: CellProperties, cells: CellProperties[]) {

    let distance: number = 0; // distance should contain info : top, left, up , down, as well as height

    Excel.run(async (context) => {

      for (let i = 0; i < cells.length; i++) {

        let type: Excel.GeometricShapeType;
        var shapes = context.workbook.worksheets.getItem("Probability").shapes;

        if (focusCell.top == cells[i].top) {
          // negative distance is not handled at the moment
          distance = (focusCell.left - cells[i].left);
          type = Excel.GeometricShapeType.curvedDownArrow;

          let arrow = shapes.addGeometricShape(type);
          arrow.width = distance + focusCell.width + (i + 1);
          arrow.left = cells[i].left;
          arrow.top = cells[i].top + 10;
          arrow.height = 10 * (cells.length - i); // 10 is to be replaced by something dynamic, depending on the samples
          arrow.incrementTop(-10 * (cells.length - i));
          arrow.fill.setSolidColor("orange");
          // arrow.fill.transparency = 0.9;
          arrow.lineFormat.visible = false;
          arrow.name = "arrow";
          // arrow.rotation = rotation;

        }

        if (focusCell.left == cells[i].left) {
          distance = (focusCell.top - cells[i].top);
          type = Excel.GeometricShapeType.curvedLeftArrow;
          let arrow = shapes.addGeometricShape(type);

          if (distance < 0) {
            distance = -distance;
          }

          arrow.width = 10;
          arrow.left = cells[i].left;
          arrow.top = cells[i].top;
          arrow.height = distance;
          arrow.incrementTop(-10 * (i + 1));
          arrow.fill.setSolidColor("orange");
          // arrow.fill.transparency = 0.7;
          arrow.lineFormat.visible = false;
          arrow.name = "arrow";
          arrow.rotation = 180;
        }

        await context.sync();
      }
    })
  }
  addOutArrows(focusCell: CellProperties, cells: CellProperties[]) {

    let distance: number = 0; // distance should contain info : top, left, up , down, as well as height

    Excel.run(async (context) => {

      for (let i = 0; i < cells.length; i++) {

        let type: Excel.GeometricShapeType;
        var shapes = context.workbook.worksheets.getItem("Probability").shapes;

        if (focusCell.top == cells[i].top) {
          // negative distance is not handled at the moment
          distance = (focusCell.left - cells[i].left);
          type = Excel.GeometricShapeType.curvedDownArrow;

          if (distance < 0) {
            console.log("Top: ", distance);
          }

          let arrow = shapes.addGeometricShape(type);
          arrow.width = distance + focusCell.width + (i + 1);
          arrow.left = focusCell.left;
          arrow.top = focusCell.top + 10;
          arrow.height = 10 * (cells.length - i); // 10 is to be replaced by something dynamic, depending on the samples
          arrow.incrementTop(-10 * (cells.length - i));
          arrow.fill.setSolidColor("blue");
          arrow.fill.transparency = 0.5;
          arrow.lineFormat.visible = false;
          arrow.name = "arrow";
          // arrow.rotation = rotation;

        }

        if (focusCell.left == cells[i].left) {
          distance = (focusCell.top - cells[i].top);
          type = Excel.GeometricShapeType.curvedRightArrow;
          let arrow = shapes.addGeometricShape(type);
          let incrementLeft = 0;

          if (distance > 0) {
            console.log("Incrementing: " + focusCell.width);
            incrementLeft = focusCell.width;
          }

          if (distance < 0) {
            console.log("Left: ", distance);
            distance = -distance;
            let rotation = 0;
          }

          arrow.width = 10;
          arrow.left = focusCell.left;
          arrow.incrementLeft(incrementLeft);
          arrow.top = focusCell.top;
          arrow.height = distance;
          arrow.incrementTop(10 * (i + 1));
          arrow.fill.setSolidColor("blue");
          arrow.fill.transparency = 0.5;
          arrow.lineFormat.visible = false;
          arrow.name = "arrow";
          // arrow.rotation = rotation;

        }

        await context.sync();
      }
    })
  }
}