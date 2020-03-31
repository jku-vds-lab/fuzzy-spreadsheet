import CellProperties from "../cellproperties";
import { timeTuesday, stackOrderInsideOut } from "d3";
import CellOperations from "../celloperations";

/* global console, Excel */
export default class Relationship {

  private referenceCell: CellProperties;
  private cells: CellProperties[];
  private degreeOfNeighbourhood: number;
  private diamonds: Promise<void>[];

  constructor(cells: CellProperties[], referenceCell: CellProperties) {
    this.cells = cells;
    this.referenceCell = referenceCell;
    this.diamonds = new Array<Promise<void>>();
  }

  showInputRelationship(n: number) {
    let colors = new Array<string>('black', 'grey', 'lightgrey');
    this.addInputRelation(this.referenceCell, n, 0, colors);
  }

  showOutputRelationship(n: number) {
    let colors = new Array<string>('black', 'grey', 'lightgrey');
    this.addOutputRelation(this.referenceCell, n, 0, colors);
  }

  removeInputRelationship() {
    this.cells.forEach((cell: CellProperties) => {
      if (cell.isInputRelationship) {
        cell.isInputRelationship = false;
      }
    })
    this.deleteRelationshipIcons('InputRelationship');
  }

  removeOutputRelationship() {
    this.cells.forEach((cell: CellProperties) => {
      if (cell.isOutputRelationship) {
        cell.isOutputRelationship = false;
      }
    })
    this.deleteRelationshipIcons('OutputRelationship');
  }

  private async deleteRelationshipIcons(name: string) {

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        var shapes = sheet.shapes;
        shapes.load("items/name");

        return context.sync().then(function () {
          shapes.items.forEach(function (shape) {
            if (shape.name.includes(name)) {
              shape.delete();
            }
          });
          return context.sync();
        });
      });
    } catch (error) {
      console.log(error);
    }
  }

  private drawInputRelation(cell: CellProperties, color: string) {
    try {

      Excel.run(function (context) {

        let type: Excel.GeometricShapeType;
        var shapes = context.workbook.worksheets.getActiveWorksheet().shapes;

        type = Excel.GeometricShapeType.diamond;
        let diamond = shapes.addGeometricShape(type);
        diamond.name = cell.address + "InputRelationship";
        diamond.left = cell.left;
        diamond.top = cell.top + cell.height / 4;
        diamond.height = 6;
        diamond.width = 6;
        diamond.lineFormat.weight = 0;
        diamond.lineFormat.color = color;
        diamond.fill.setSolidColor(color);

        return context.sync(); //.then(() => { console.log('Success in drawing relationship') }).catch((reason: any) => console.log('Could not draw relationship: ' + reason));
      })
    } catch (error) {
      console.log('Input Relationship Error: ', error);
    }
  }

  private addInputRelation(cell: CellProperties, n: number, colorIndex: number, colors: string[]) {

    cell.inputCells.forEach((inCell: CellProperties) => {

      if (!inCell.isInputRelationship) {
        // eslint-disable-next-line no-undef
        setTimeout(() => this.drawInputRelation(inCell, colors[colorIndex]), 100);
        inCell.isInputRelationship = true;
      }

      if (n == 1) {
        return;
      }

      let newColorIndex = colorIndex + 1;

      this.addInputRelation(inCell, n - 1, newColorIndex, colors);
    })
  }

  private addOutputRelation(cell: CellProperties, n: number, colorIndex: number, colors: string[]) {

    cell.outputCells.forEach((outCell: CellProperties) => {

      if (!outCell.isOutputRelationship) {
        // eslint-disable-next-line no-undef
        setTimeout(() => this.drawOutputRelation(outCell, colors[colorIndex]), 100);
        outCell.isOutputRelationship = true;
      }

      if (n == 1) {
        return;
      }

      let newColorIndex = colorIndex + 1;

      this.addOutputRelation(outCell, n - 1, newColorIndex, colors);
    })
  }

  private drawOutputRelation(cell: CellProperties, color: string) {

    try {
      Excel.run(async (context) => {
        let type: Excel.GeometricShapeType;
        var shapes = context.workbook.worksheets.getActiveWorksheet().shapes;

        type = Excel.GeometricShapeType.ellipse;
        let circle = shapes.addGeometricShape(type);
        circle.name = cell.address + "OutputRelationship"
        circle.left = cell.left;
        circle.top = cell.top + cell.height / 4;
        circle.height = 6;
        circle.width = 6;
        circle.lineFormat.weight = 0;
        circle.lineFormat.color = color;
        circle.fill.setSolidColor(color);
        await context.sync();
      })

    } catch (error) {
      console.log('Output relationship error: ', error);
    }
  }
}