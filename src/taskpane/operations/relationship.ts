// /* global console, Excel */
// export default class Relationship {
//   addInArrows(focusCell: CellProperties, cells: CellProperties[]) {

//     Excel.run(async (context) => {

//       for (let i = 0; i < cells.length; i++) {

//         let type: Excel.GeometricShapeType;
//         var shapes = context.workbook.worksheets.getActiveWorksheet().shapes;

//         type = Excel.GeometricShapeType.triangle;
//         let triangle = shapes.addGeometricShape(type);
//         triangle.rotation = 90;
//         triangle.left = cells[i].left;
//         triangle.top = cells[i].top + cells[i].height / 4;
//         triangle.height = 3;
//         triangle.width = 6;
//         triangle.lineFormat.weight = 0;
//         triangle.lineFormat.color = 'black';
//         triangle.fill.setSolidColor('black');
//       }

//       await context.sync();
//     })
//   }

//   addOutArrows(focusCell: CellProperties, cells: CellProperties[]) {

//     Excel.run(async (context) => {

//       for (let i = 0; i < cells.length; i++) {
//         let type: Excel.GeometricShapeType;
//         var shapes = context.workbook.worksheets.getActiveWorksheet().shapes;

//         type = Excel.GeometricShapeType.triangle;
//         let triangle = shapes.addGeometricShape(type);
//         triangle.rotation = 270;
//         triangle.left = cells[i].left;
//         triangle.top = cells[i].top + cells[i].height / 4;
//         triangle.height = 3;
//         triangle.width = 6;
//         triangle.lineFormat.weight = 0;
//         triangle.lineFormat.color = 'black';
//         triangle.fill.setSolidColor('black');
//       }
//       await context.sync();
//     })
//   }
// }