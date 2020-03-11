/* global console, Excel */

import CellProperties from "../cellproperties";
import SheetProperties from "../sheetproperties";

export default class CommonOperations {

  drawRectangle(cell: CellProperties, type: string) {

    Excel.run((context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let i = 0;
      let MARGIN = 5;
      let height = 5;
      let width = 5;

      cell.rect = sheet.shapes.addGeometricShape("Rectangle");
      cell.rect.name = "Shape" + type + i;
      cell.rect.left = cell.left + MARGIN;
      cell.rect.top = cell.top + cell.height / 4;

      if (SheetProperties.isLikelihood) {
        height = cell.likelihood * 10;
        width = cell.likelihood * 10;
      }

      cell.rect.height = height;
      cell.rect.width = width;

      cell.rect.geometricShapeType = Excel.GeometricShapeType.rectangle;
      cell.rect.fill.setSolidColor(cell.rectColor);
      cell.rect.fill.transparency = cell.rectTransparency;
      cell.rect.lineFormat.weight = 0;
      cell.rect.lineFormat.color = cell.rectColor;
      i++;
      return context.sync();
    });
  }

  async deleteRectangles(cells: CellProperties[], type: string) {

    try {

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        var shapes = sheet.shapes;
        shapes.load("items/name");

        return context.sync().then(function () {
          shapes.items.forEach(function (shape) {
            if (shape.name.includes('Shape' + type)) {
              shape.delete();
            }
          });
        }).catch((reason: any) => console.log('Could not delete the shape: ' + reason));
      });

    } catch (error) {
      console.log(error);
    }
  }
}