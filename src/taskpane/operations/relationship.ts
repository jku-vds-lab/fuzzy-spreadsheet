import CellProperties from "../cellproperties";

/* global console, Excel */
export default class Relationship {

  private referenceCell: CellProperties;
  private degreeOfNeighbourhood: number;

  constructor(referenceCell: CellProperties) {
    this.referenceCell = referenceCell;
  }

  showInputRelationship() {
    this.addInputRelation(this.referenceCell.inputCells);
  }

  showOutputRelationship() {
    this.addOutputRelation(this.referenceCell.outputCells);
  }

  removeInputRelationship() {
    this.deleteTriangles('Input');
  }

  removeOutputRelationship() {
    this.deleteTriangles('Output');
  }

  private async deleteTriangles(type: string) {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      var shapes = sheet.shapes;
      shapes.load("items/name");

      return context.sync().then(function () {
        shapes.items.forEach(function (shape) {
          if (shape.name.includes(type)) {
            shape.delete();
          }
        });
        return context.sync();
      });
    });
  }

  private addInputRelation(cells: CellProperties[]) {

    Excel.run(async (context) => {

      for (let i = 0; i < cells.length; i++) {

        let type: Excel.GeometricShapeType;
        var shapes = context.workbook.worksheets.getActiveWorksheet().shapes;

        type = Excel.GeometricShapeType.triangle;
        let triangle = shapes.addGeometricShape(type);
        triangle.name = "Input" + i;
        triangle.rotation = 90;
        triangle.left = cells[i].left;
        triangle.top = cells[i].top + cells[i].height / 4;
        triangle.height = 3;
        triangle.width = 6;
        triangle.lineFormat.weight = 0;
        triangle.lineFormat.color = 'black';
        triangle.fill.setSolidColor('black');
      }

      await context.sync();
    })
  }

  private addOutputRelation(cells: CellProperties[]) {

    Excel.run(async (context) => {

      for (let i = 0; i < cells.length; i++) {
        let type: Excel.GeometricShapeType;
        var shapes = context.workbook.worksheets.getActiveWorksheet().shapes;

        type = Excel.GeometricShapeType.triangle;
        let triangle = shapes.addGeometricShape(type);
        triangle.name = "Output" + i;
        triangle.rotation = 270;
        triangle.left = cells[i].left;
        triangle.top = cells[i].top + cells[i].height / 4;
        triangle.height = 3;
        triangle.width = 6;
        triangle.lineFormat.weight = 0;
        triangle.lineFormat.color = 'black';
        triangle.fill.setSolidColor('black');
      }
      await context.sync();
    })
  }
}