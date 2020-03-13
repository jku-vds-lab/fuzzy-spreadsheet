/* global console, Excel */
import * as OfficeHelpers from '@microsoft/office-js-helpers';
import CellProperties from "../cellproperties";
import SheetProperties from "../sheetproperties";

export default class CommonOperations {

  drawRectangle(cell: CellProperties, type: string) {

    Excel.run((context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let i = 0;
      let MARGIN = 10;
      let height = 5;
      let width = 5;

      cell.rect = sheet.shapes.addGeometricShape("Rectangle");
      cell.rect.name = "Shape" + type;
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

  deleteRectangles(cells: CellProperties[], type: string) {

    try {

      Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        var shapes = sheet.shapes;
        shapes.load("items/name");

        return context.sync().then(function () {
          shapes.items.forEach(function (shape) {
            if (shape.name.includes('Shape' + type)) {
              console.log('Name: ' + shape.name);
              shape.delete();
            }
          });
          return context.sync();
        }).catch((reason: any) => {
          console.log('Step 1:', reason, type)
        });
      });
    } catch (error) {
      console.log('Step 2:', error);
      OfficeHelpers.Utilities.log(error);
    }
  }
}