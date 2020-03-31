import CellProperties from "../cellproperties";

/* global console, Excel */
export default class Relationship {

  private referenceCell: CellProperties;
  private inputCellsWithRelationship: { cell: CellProperties, color: string }[];
  private outputCellsWithRelationship: { cell: CellProperties, color: string }[];

  constructor(referenceCell: CellProperties) {
    this.referenceCell = referenceCell;
  }

  showInputRelationship(n: number) {

    let colors = new Array<string>('black', 'grey', 'lightgrey');
    this.inputCellsWithRelationship = new Array<{ cell: CellProperties, color: string }>();
    this.addInputRelation(this.referenceCell, n, 0, colors);
    this.drawInputRelation(this.inputCellsWithRelationship, 'InputRelationship');
  }

  showOutputRelationship(n: number) {

    let colors = new Array<string>('black', 'grey', 'lightgrey');
    this.outputCellsWithRelationship = new Array<{ cell: CellProperties, color: string }>();
    this.addOutputRelation(this.referenceCell, n, 0, colors);
    this.drawOutputRelation(this.outputCellsWithRelationship, 'OutputRelationship');
  }

  private addInputRelation(cell: CellProperties, n: number, colorIndex: number, colors: string[]) {

    cell.inputCells.forEach((inCell: CellProperties) => {

      if (!inCell.isInputRelationship) {

        this.inputCellsWithRelationship.push({ cell: inCell, color: colors[colorIndex] });
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

        this.outputCellsWithRelationship.push({ cell: outCell, color: colors[colorIndex] });
        outCell.isOutputRelationship = true;
      }

      if (n == 1) {
        return;
      }

      let newColorIndex = colorIndex + 1;

      this.addOutputRelation(outCell, n - 1, newColorIndex, colors);
    })
  }

  private drawInputRelation(cellsWithColors: { cell: CellProperties, color: string }[], name: string) {
    try {

      Excel.run(function (context) {

        var shapes = context.workbook.worksheets.getActiveWorksheet().shapes;

        cellsWithColors.forEach((element: { cell: CellProperties, color: string }) => {
          let type = Excel.GeometricShapeType.diamond;
          let diamond = shapes.addGeometricShape(type);
          diamond.name = element.cell.address + name;
          diamond.left = element.cell.left;
          diamond.top = element.cell.top + element.cell.height / 4;
          diamond.height = 6;
          diamond.width = 6;
          diamond.lineFormat.weight = 0;
          diamond.lineFormat.color = element.color;
          diamond.fill.setSolidColor(element.color);
        })

        return context.sync();
      })
    } catch (error) {
      console.log('Input Relationship Error: ', error);
    }
  }

  private drawOutputRelation(cellsWithColors: { cell: CellProperties, color: string }[], name: string) {

    try {
      Excel.run(async (context) => {

        var shapes = context.workbook.worksheets.getActiveWorksheet().shapes;

        cellsWithColors.forEach((element: { cell: CellProperties, color: string }) => {
          let type = Excel.GeometricShapeType.ellipse;
          let circle = shapes.addGeometricShape(type);
          circle.name = element.cell.address + name
          circle.left = element.cell.left;
          circle.top = element.cell.top + element.cell.height / 4;
          circle.height = 6;
          circle.width = 6;
          circle.lineFormat.weight = 0;
          circle.lineFormat.color = element.color;
          circle.fill.setSolidColor(element.color);
        });

        return context.sync();
      })
    } catch (error) {
      console.log('Output relationship error: ', error);
    }
  }
}