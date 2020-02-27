/* global console, Excel */

import CellProperties from "../cellproperties";
import SheetProperties from "../sheetproperties";

export default class CommonOperations {

  drawRectangle(cell: CellProperties) {

    Excel.run((context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let i = 0;
      let MARGIN = 5;
      let height = 5;
      let width = 5;

      cell.rect = sheet.shapes.addGeometricShape("Rectangle");
      cell.rect.name = "Shape" + i;
      cell.rect.left = cell.left + MARGIN;
      cell.rect.top = cell.top + cell.height / 4;

      if (SheetProperties.isLikelihood) {
        height = cell.likelihood;
        width = cell.likelihood;
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

  async deleteRectangles(cells: CellProperties[]) {

    // remove both impact and likelihood here?

    cells.forEach((cell: CellProperties) => {
      cell.isImpact = false;
      cell.isLikelihood = false;
    })

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      var shapes = sheet.shapes;
      shapes.load("items/name");

      return context.sync().then(function () {
        shapes.items.forEach(function (shape) {
          if (shape.name.includes('Shape')) {
            shape.delete();
          }
        });
        return context.sync();
      });
    });
  }
}