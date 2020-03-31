/* global console, Excel */
import * as OfficeHelpers from '@microsoft/office-js-helpers';
import CellProperties from "../cellproperties";
import SheetProperties from "../sheetproperties";

export default class CommonOperations {
  private referenceCell: CellProperties;

  constructor(referenceCell: CellProperties) {
    this.referenceCell = referenceCell;
  }

  drawRectangle(cells: CellProperties[], name: string) {

    try {
      Excel.run((context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let MARGIN = 10;
        let height = 5;
        let width = 5;

        cells.forEach((cell: CellProperties) => {

          cell.rect = sheet.shapes.addGeometricShape("Rectangle");
          cell.rect.name = cell.address + name;
          cell.rect.left = cell.left + MARGIN;
          cell.rect.top = cell.top + cell.height / 4;

          // if impact does not exist, only then
          if (cell.isLikelihood) {
            height = cell.likelihood * 10;
            width = cell.likelihood * 10;
            cell.rectColor = 'gray';
            cell.rectTransparency = 0;
          }

          cell.rect.height = height;
          cell.rect.width = width;

          cell.rect.geometricShapeType = Excel.GeometricShapeType.rectangle;
          cell.rect.fill.setSolidColor(cell.rectColor);
          cell.rect.fill.transparency = cell.rectTransparency;
          cell.rect.lineFormat.weight = 0;
          cell.rect.lineFormat.color = cell.rectColor;
        })

        return context.sync();
      });
    } catch (error) {
      console.log('---' + name + ' : ', error);
    }
  }

  // To remove shapes from reference cell
  removeShapesReferenceCellWise() {
    this.deleteShapes('Reference');
  }

  // To remove a particular option: such as spread
  removeShapesOptionWise(optionName: string) {
    this.deleteShapes(optionName);
  }

  // To remove a particular influence: such as influence by (or input cells)
  removeShapesInfluenceWise(influenceType: string) {
    this.deleteShapes(influenceType);
  }

  // To remove updated shapes
  removeShapesUpdatedWise() {
    this.deleteShapes('Update');
  }

  // To remove shapes based of degree of neighbourhood
  removeShapesNeighbourWise(n: number) {
    if (n == 1) {
      this.removeSecondDegreeInputNeighbours();
      this.removeThirdDegreeInputNeighbours();

      this.removeSecondDegreeOutputNeighbours();
      this.removeThirdDegreeOutputNeighbours();
    }

    if (n == 2) {
      this.removeThirdDegreeInputNeighbours();

      this.removeThirdDegreeOutputNeighbours();
    }
  }

  private setPropertiesToFalse(cell: CellProperties) {
    cell.isSpread = false;
    cell.isInputRelationship = false;
    cell.isOutputRelationship = false;
    cell.isLikelihood = false;
    cell.isImpact = false;
  }

  private removeSecondDegreeInputNeighbours() {

    let names = new Array<string>();
    this.referenceCell.inputCells.forEach((inCell: CellProperties) => {
      inCell.inputCells.forEach((inincell: CellProperties) => {
        this.setPropertiesToFalse(inincell);
        names.push(inincell.address);
      })
    })
    this.deleteShapesInCells(names);
  }

  private removeThirdDegreeInputNeighbours() {
    let names = new Array<string>();
    this.referenceCell.inputCells.forEach((inCell: CellProperties) => {
      inCell.inputCells.forEach((inincell: CellProperties) => {
        inincell.inputCells.forEach((ininincell: CellProperties) => {
          this.setPropertiesToFalse(ininincell);
          names.push(ininincell.address);
        })
      })
    })
    this.deleteShapesInCells(names);
  }

  private removeSecondDegreeOutputNeighbours() {
    let names = new Array<string>();
    this.referenceCell.outputCells.forEach((outCell: CellProperties) => {
      outCell.outputCells.forEach((outoutcell: CellProperties) => {
        this.setPropertiesToFalse(outoutcell);
        names.push(outoutcell.address);
      })
    })
    this.deleteShapesInCells(names);
  }

  private removeThirdDegreeOutputNeighbours() {
    let names = new Array<string>();
    this.referenceCell.outputCells.forEach((outCell: CellProperties) => {
      outCell.outputCells.forEach((outoutcell: CellProperties) => {
        outoutcell.outputCells.forEach((outoutoutCell: CellProperties) => {
          this.setPropertiesToFalse(outoutoutCell);
          names.push(outoutoutCell.address);
        })
      })
    })
    this.deleteShapesInCells(names);
  }

  private deleteShapesInCells(names: string[]) {

    try {

      Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        var shapes = sheet.shapes;
        shapes.load("items/name");

        return context.sync().then(function () {
          names.forEach((name: string) => {
            shapes.items.forEach(function (shape) {
              if (shape.name.includes(name)) {
                shape.delete();
              }
            });
          })
        }).catch((reason: any) => {
          console.log('Step 1:', reason)
        });
      });
    } catch (error) {
      console.log('Step 2:', error);
    }
  }

  private deleteShapes(name: string) {
    try {

      Excel.run((context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        var shapes = sheet.shapes;
        shapes.load("items/name");

        return context.sync().then(function () {
          shapes.items.forEach(function (shape) {
            if (shape.name.includes(name)) {
              shape.delete();
            }
          });
        }).catch((reason: any) => {
          console.log('Step 1:', reason, name)
        });
      });
    } catch (error) {
      console.log('Step 2:', error);
    }
  }
}