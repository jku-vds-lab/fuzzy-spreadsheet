/* global console, Excel */
import { std, ceil } from 'mathjs';
import * as jstat from 'jstat';
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

    if (focusCell.formula.includes("SUM")) {
      focusCell.inputCells.forEach((inCell: CellProperties) => {

        let colorProperties = this.inputColorProperties(inCell.value, focusCell.value, focusCell.inputCells);

        console.log("SUM: " + colorProperties.transparency);

        let customShape: CustomShape = { cell: inCell, shape: null, color: colorProperties.color, transparency: colorProperties.transparency }
        this.customShapes.push(customShape);
      })
    }

    if (focusCell.formula.includes('-')) {

      let formula: string = focusCell.formula;
      let idx = formula.indexOf('-');
      let subtrahend = formula.substring(idx + 1, formula.length);

      console.log(subtrahend);

      focusCell.inputCells.forEach((inCell: CellProperties) => {
        let isSubtrahend = false;

        if (inCell.address.includes(subtrahend)) {
          isSubtrahend = true;
          console.log(inCell.address);
        }

        let colorProperties = this.inputColorProperties(inCell.value, focusCell.value, focusCell.inputCells, isSubtrahend);

        console.log("SUM: " + colorProperties.transparency);

        let customShape: CustomShape = { cell: inCell, shape: null, color: colorProperties.color, transparency: colorProperties.transparency }
        this.customShapes.push(customShape);

      });
    }

    focusCell.outputCells.forEach((outCell: CellProperties) => {

      let colorProperties = this.outputColorProperties(outCell.value, focusCell.value, outCell.inputCells);

      console.log("OUT: " + colorProperties.transparency);

      let customShape: CustomShape = { cell: outCell, shape: null, color: colorProperties.color, transparency: colorProperties.transparency }
      this.customShapes.push(customShape);
    })
  }

  async addImpact(focusCell: CellProperties) {

    if (this.customShapes != null) {
      this.deleteCustomShapes();
    }
    this.addImpactInfo(focusCell);
    this.drawCustomShapes();
  }

  private drawCustomShapes() {
    try {
      Excel.run((context) => {

        const sheet = context.workbook.worksheets.getItem("Probability");
        let i = 0;

        this.customShapes.forEach((customShape: CustomShape) => {
          customShape.shape = sheet.shapes.addGeometricShape("Rectangle");
          customShape.shape.name = "Impact" + i;
          console.log(customShape.shape.name);
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
        return context.sync();
      });

    } catch (error) {
      console.error(error);
    }
  }

  private deleteCustomShapes() {
    try {
      Excel.run(async (context) => {

        const sheet = context.workbook.worksheets.getItem("Probability");
        let i = 0;

        this.customShapes.forEach((customShape: CustomShape) => {
          customShape.shape = sheet.shapes.getItem("Impact" + i)
          customShape.shape.delete();
          i++;

        })
        // createImpactLegend().then(function () { });
        await context.sync();
      });

    } catch (error) {
      console.error(error);
    }
  }

  private computeColor(cellValue: number, focusCellValue: number, cells: CellProperties[], isSubtrahend: boolean = false) {

    let color = "green";


    if (isSubtrahend) {
      if (cellValue > 0) {
        color = "red";
      }
      return color;
    }

    if (focusCellValue > 0 && cellValue < 0) {
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

  private inputColorProperties(cellValue: number, focusCellValue: number, cells: CellProperties[], isSubtrahend: boolean = false) {

    let transparency = 0;
    const color = this.computeColor(cellValue, focusCellValue, cells, isSubtrahend);

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

    // use the info of uncertain cells
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
  async addLikelihood(focusCell: CellProperties) {

    this.addLikelihoodInfo();
    let k = 0;
    try {

      if (this.customShapes == null) {

        this.customShapes = new Array<CustomShape>();

        focusCell.inputCells.forEach((inCell: CellProperties) => {
          let customShape: CustomShape = { cell: inCell, shape: null, color: 'gray', transparency: 0 }
          this.customShapes.push(customShape);
        });

        focusCell.outputCells.forEach((outCell: CellProperties) => {
          let customShape: CustomShape = { cell: outCell, shape: null, color: 'gray', transparency: 0 }
          this.customShapes.push(customShape);
        });

        this.drawCustomShapes();
      }

      await Excel.run(async (context) => {

        for (let i = 0; i < this.customShapes.length; i++) {

          const sheet = context.workbook.worksheets.getItem("Probability");
          let key = "Impact" + i;
          let shape = sheet.shapes.getItem(key);
          k++;
          shape.load(["height", "width", "top"]);

          await context.sync();

          let likelihood = this.customShapes[i].cell.likelihood / 10;

          if (likelihood == 10) {
            likelihood = this.cells[i].height;
            shape.top = shape.top - 4;
          }
          shape.height = likelihood;
          shape.width = likelihood;
        }

        // createLikelihoodLegend().then(function () { });
        await context.sync();
      });
    } catch (error) {
      console.log("Didn't work for " + k);
      console.error(error);
    }

  }

  private addVarianceInfo() {

    // use the info of uncertain cells
    for (let i = 0; i < this.cells.length; i++) {
      for (let r = 5; r < 18; r++) {
        let id = "R" + r + "C8";
        if (this.cells[i].id == id) {
          this.cells[i].variance = this.cells[i + 2].value;
        }
      }
    }
  }

  async createNormalDistributions() {

    this.addVarianceInfo();
    await Excel.run(async (context) => {

      let cheatsheet = context.workbook.worksheets.getItemOrNullObject("CheatSheet");
      await context.sync();

      if (!cheatsheet.isNullObject) {
        cheatsheet.delete();
      }

      cheatsheet = context.workbook.worksheets.add("CheatSheet");
      let rowIndex = -1;
      // let min = mean - variance * 2;
      // let max = mean + variance * 2;

      for (let c = 0; c < this.cells.length; c++) {

        this.cells[c].samples = new Array<number>();


        let overallMin = -10;
        let overallMax = 40;
        let mean = this.cells[c].value;


        let variance = this.cells[c].variance

        if (variance > 0) {
          rowIndex++;
          let sampleSize = (variance * 2) / 50;

          for (let i = overallMin; i <= overallMax; i = i + sampleSize) {
            this.cells[c].samples.push(jstat.normal.pdf(i, mean, variance));
          }
        }
        else {
          rowIndex++;
          if (this.cells[c].degreeToFocus >= 0) {
            for (let i = overallMin; i <= overallMax; i++) {
              if (i == ceil(this.cells[c].value)) {
                this.cells[c].samples.push(1);
              } else {
                this.cells[c].samples.push(0);
              }
            }
          }
        }

        if (this.cells[c].samples.length == 0) {
          continue;
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

  async addSpread(focusCell: CellProperties) {

    await this.createNormalDistributions();

    this.drawLineChart(focusCell);
    // this.drawCompleteLineChart(focusCell);

    focusCell.inputCells.forEach((cell: CellProperties) => {
      this.drawLineChart(cell);

    })

    focusCell.outputCells.forEach((cell: CellProperties) => {
      this.drawLineChart(cell);
    })
  }

  private drawLineChart(cell: CellProperties) {

    console.log(cell, cell.spreadRange);
    if (cell.spreadRange == null) {
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
      return context.sync();
    });
  }

  private showPopUpWindow(cell: CellProperties) {

    if (cell.spreadRange == null) {
      return;
    }

    Excel.run((context) => {

      const sheet = context.workbook.worksheets.getItem("Probability");
      const cheatSheet = context.workbook.worksheets.getItem("CheatSheet");

      let MARGIN = 2 * cell.width;
      let TEXTMARGIN = 50;
      let TOPMARGIN = 3 * cell.height;

      let impact = sheet.shapes.addGeometricShape("Rectangle");
      impact.left = cell.left + MARGIN;
      impact.top = cell.top + cell.height;
      impact.height = 5;
      impact.width = 5;
      impact.geometricShapeType = Excel.GeometricShapeType.rectangle;
      impact.fill.setSolidColor('green');
      impact.lineFormat.weight = 0;
      impact.lineFormat.color = 'green';

      let textbox = sheet.shapes.addTextBox('50 % Positive Impact');
      textbox.left = cell.left + MARGIN + TEXTMARGIN;
      textbox.top = cell.top;

      let likelihood = sheet.shapes.addGeometricShape("Rectangle");
      likelihood.left = cell.left + MARGIN;
      likelihood.top = cell.top + TOPMARGIN;
      likelihood.height = 20;
      likelihood.width = 20;
      likelihood.geometricShapeType = Excel.GeometricShapeType.rectangle;
      likelihood.fill.setSolidColor('gray');
      likelihood.lineFormat.weight = 0;
      likelihood.lineFormat.color = 'gray';

      let textbox2 = sheet.shapes.addTextBox('100 % Likelihood');
      textbox2.left = cell.left + MARGIN + TEXTMARGIN
      textbox2.top = cell.top + TOPMARGIN;

      const dataRange = cheatSheet.getRange(cell.spreadRange);
      let chart = sheet.charts.add("Line", dataRange, Excel.ChartSeriesBy.auto);
      chart.setPosition(cell.address);
      chart.left = cell.left + MARGIN;
      chart.top = cell.top + 2 * TOPMARGIN;
      chart.title.visible = false;
      chart.format.fill.clear();
      chart.format.border.clear();

      let textbox3 = sheet.shapes.addTextBox('Mean and Variance');
      textbox3.left = cell.left + MARGIN;
      textbox3.top = cell.top + 2 * TOPMARGIN + 250;

      let shape1 = sheet.shapes.addGeometricShape("Rectangle");
      // shape.name = "Impact" + i;
      shape1.left = cell.left + cell.width;
      shape1.top = cell.top - cell.height;
      shape1.height = 400;
      shape1.width = 500;
      shape1.geometricShapeType = Excel.GeometricShapeType.rectangle;
      shape1.fill.setSolidColor('ADD8E6');
      shape1.lineFormat.weight = 0;
      shape1.lineFormat.color = 'ADD8E6';
      shape1.setZOrder(Excel.ShapeZOrder.sendToBack);
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