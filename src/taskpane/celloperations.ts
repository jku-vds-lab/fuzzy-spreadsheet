/* global console, Excel */
import CellProperties from './cellproperties';
import Impact from './operations/impact';
import Likelihood from './operations/likelihood';
import Spread from './operations/spread';
import Relationship from './operations/relationship';
import SheetProperties from './sheetproperties';
import DiscreteSpread from './operations/spread';

export default class CellOperations {

  private cells: CellProperties[];
  private referenceCell: CellProperties;
  private degreeOfNeighbourhood: number = 1;
  private impact: Impact;
  private likelihood: Likelihood;
  private spread: Spread;
  private relationship: Relationship;

  constructor(cells: CellProperties[], referenceCell: CellProperties, n: number) {
    this.cells = cells;
    this.referenceCell = referenceCell;
    this.degreeOfNeighbourhood = n;
    this.impact = new Impact(this.referenceCell, this.cells);
    this.likelihood = new Likelihood(this.cells, this.referenceCell);
    this.spread = new Spread(this.cells, this.referenceCell);
    this.relationship = new Relationship(this.cells, this.referenceCell);
  }

  getCells() {
    return this.cells;
  }

  getDegreeOfNeighbourhood() {
    return this.degreeOfNeighbourhood;
  }

  showImpact(n: number) {
    this.impact.showImpact(n);
  }

  async removeImpact(n: number) {
    await this.impact.removeImpact(n);
  }

  showLikelihood(n: number) {
    this.likelihood.showLikelihood(n);
  }

  async removeLikelihood(n: number) {
    await this.likelihood.removeLikelihood(n);
  }

  showSpread(n: number) {
    this.spread.showSpread(n);
  }

  async removeSpread() {
    await this.spread.removeSpread();
  }

  showInputRelationship(n: number) {
    this.relationship.showInputRelationship(n);
  }

  removeInputRelationship() {
    this.relationship.removeInputRelationship();
  }

  showOutputRelationship(n: number) {
    this.relationship.showOutputRelationship(n);
  }

  removeOutputRelationship() {
    this.relationship.removeOutputRelationship();
  }

  deleteUpdateshapes() {

    Excel.run(function (context) {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      const oldTextbox = sheet.shapes;
      oldTextbox.load("items/name");
      return context.sync().then(function () {
        oldTextbox.items.forEach(function (c) {
          if (c.name.includes('Update'))
            c.delete();
          console.log('Deleted shape: ' + c.name);
        });
      });

    })
  }



  async showPopUpWindow(address: string) {
    // this.removePopUps();
    // console.log(address);

    // if (SheetProperties.isImpact) {

    //   this.cells.forEach(async (cell: CellProperties) => {
    //     if (cell.address.includes(address)) {
    //       await this.showImpactInPopUp(cell);
    //       return;
    //     }
    //   })
    // }
    // this.cells.forEach((cell: CellProperties) => {
    //   if (cell.address.includes(address)) {
    //     if (cell.spreadRange == null) {
    //       return;
    //     }

    //     Excel.run((context) => {

    //       const sheet = context.workbook.worksheets.getActiveWorksheet();
    //       const cheatSheet = context.workbook.worksheets.getItem("CheatSheet");

    //       let MARGIN = 120
    //       let TEXTMARGIN = 20;
    //       let TOPMARGIN = 15;

    //       console.log("MArgin: " + MARGIN + " TOP MARGIN: " + TOPMARGIN);
    //       let shape1 = sheet.shapes.addGeometricShape("Rectangle");
    //       shape1.name = "Pop7";
    //       shape1.left = cell.left + cell.width;
    //       shape1.top = cell.top - cell.height;
    //       shape1.height = 250;
    //       shape1.width = 300;
    //       shape1.geometricShapeType = Excel.GeometricShapeType.rectangle;
    //       shape1.fill.setSolidColor('ADD8E6');
    //       shape1.lineFormat.weight = 0;
    //       shape1.lineFormat.color = 'ADD8E6';

    //       if (cell.isImpact) {

    //         this.customShapes.forEach((cShape: CustomShape) => {


    //       if (cell.isLikelihood) {

    //         let likelihood = sheet.shapes.addGeometricShape("Rectangle");
    //         likelihood.name = "Pop3";
    //         likelihood.left = cell.left + MARGIN;
    //         likelihood.top = cell.top + 2 * TOPMARGIN;
    //         likelihood.height = cell.likelihood / 10;
    //         likelihood.width = cell.likelihood / 10;
    //         likelihood.geometricShapeType = Excel.GeometricShapeType.rectangle;
    //         likelihood.fill.setSolidColor('gray');
    //         likelihood.lineFormat.weight = 0;
    //         likelihood.lineFormat.color = 'gray';
    //         likelihood.setZOrder(Excel.ShapeZOrder.bringForward);

    //         let text = cell.likelihood + '% Likelihood';

    //         let textbox2 = sheet.shapes.addTextBox(text);
    //         textbox2.name = "Pop4";
    //         textbox2.left = cell.left + MARGIN + TEXTMARGIN
    //         textbox2.top = cell.top + 2 * TOPMARGIN;
    //         textbox2.height = 20;
    //         textbox2.width = 150;
    //         textbox2.setZOrder(Excel.ShapeZOrder.bringForward);
    //       }

    //       const dataRange = cheatSheet.getRange(cell.spreadRange);
    //       let chart = sheet.charts.add(Excel.ChartType.columnClustered, dataRange, Excel.ChartSeriesBy.auto);
    //       chart.name = "Pop5";
    //       chart.setPosition(cell.address);
    //       chart.left = cell.left + MARGIN;
    //       chart.top = cell.top + 3 * TOPMARGIN;
    //       chart.height = 180;
    //       chart.width = 210;
    //       // chart.axes.valueAxis.minimum = -10;
    //       // chart.axes.valueAxis.maximum = 40;
    //       // chart.axes.valueAxis.tickLabelSpacing = 1;
    //       chart.axes.categoryAxis.visible = false;
    //       chart.axes.valueAxis.majorGridlines.visible = false;
    //       chart.title.visible = false;
    //       chart.format.fill.clear();
    //       chart.format.border.clear();

    //       // let textbox3 = sheet.shapes.addTextBox('Mean and Variance');
    //       // textbox3.name = "Pop6";
    //       // textbox3.left = cell.left + MARGIN;
    //       // textbox3.top = cell.top + 180;
    //       // textbox3.setZOrder(Excel.ShapeZOrder.bringForward);
    //       return context.sync();
    //     });
    //   }
    // })
  }
  private async showImpactInPopUp(cell: CellProperties) {

    let MARGIN = 120
    let TEXTMARGIN = 20;
    let TOPMARGIN = 15;

    if (cell.impact == 0) {
      return;
    }

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let impact = sheet.shapes.addGeometricShape("Rectangle");
      impact.name = "Pop1";
      impact.left = cell.left + MARGIN;
      impact.top = cell.top + TOPMARGIN;
      impact.height = 5;
      impact.width = 5;
      impact.geometricShapeType = Excel.GeometricShapeType.rectangle;
      impact.fill.setSolidColor(cell.rectColor);
      impact.fill.transparency = cell.rectTransparency;
      impact.lineFormat.weight = 0;
      impact.lineFormat.color = cell.rectColor;
      impact.setZOrder(Excel.ShapeZOrder.bringForward);

      let text = cell.impact + '%';

      if (cell.rectColor == 'green') {
        text += 'Positive Impact';
      } else {
        text += 'Negative Impact';
      }

      let textbox = sheet.shapes.addTextBox(text);
      textbox.name = "Pop2";
      textbox.left = cell.left + MARGIN + TEXTMARGIN;
      textbox.top = cell.top;
      textbox.height = 20;
      textbox.width = 150;

      textbox.setZOrder(Excel.ShapeZOrder.bringForward);

      await context.sync();
    });

  }
  async removePopUps() {
    // remove();
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      var shapes = sheet.shapes;
      shapes.load("items/name");

      return context.sync().then(function () {
        shapes.items.forEach(function (shape) {
          if (shape.name.includes('Pop')) {
            shape.delete();
          }
        });
        return context.sync();
      });
    });
  }
}