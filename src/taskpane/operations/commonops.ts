/* global console, Excel */
import * as OfficeHelpers from '@microsoft/office-js-helpers';
import CellProperties from "../cellproperties";
import SheetProperties from "../sheetproperties";

export default class CommonOperations {
  private referenceCell: CellProperties;
  private cells: CellProperties[];
  private isImpact: boolean;
  private isLikelihood: boolean;
  private isRelationshipIcons: boolean;
  private isSpread: boolean;
  private isInputRelationship: boolean;
  private isOutputRelationship: boolean;

  constructor(referenceCell: CellProperties) {
    this.referenceCell = referenceCell;
    this.isImpact = false;
    this.isLikelihood = false;
    this.isRelationshipIcons = false;
    this.isSpread = false;
    this.isInputRelationship = false;
    this.isOutputRelationship = false;
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

          let color = 'gray';
          let transparency = 0;

          if (cell.isImpact) {
            color = cell.rectColor;
            transparency = cell.rectTransparency;
          }

          if (cell.isLikelihood) {
            height = cell.likelihood * 10;
            width = cell.likelihood * 10;
          }

          cell.rect.height = height;
          cell.rect.width = width;

          cell.rect.geometricShapeType = Excel.GeometricShapeType.rectangle;
          cell.rect.fill.setSolidColor(color);
          cell.rect.fill.transparency = transparency;
          cell.rect.lineFormat.weight = 0;
          cell.rect.lineFormat.color = color;
        })

        return context.sync();
      });
    } catch (error) {
      console.log('---' + name + ' : ', error);
    }
  }

  setCells(cells: CellProperties[]) {
    this.cells = cells;
  }

  setOptions(isImpact: boolean, isLikelihood: boolean, isRelationshipIcons: boolean, isSpread: boolean, isInputRelationship: boolean, isOutputRelationship: boolean) {
    this.isImpact = isImpact;
    this.isLikelihood = isLikelihood;
    this.isRelationshipIcons = isRelationshipIcons;
    this.isSpread = isSpread;
    this.isInputRelationship = isInputRelationship;
    this.isOutputRelationship = isOutputRelationship;
  }

  // To remove shapes from reference cell
  removeShapesReferenceCellWise() {
    this.referenceCell.isSpread = false;
    this.deleteShapes('Reference');
  }

  // To remove a particular option: such as spread
  removeShapesOptionWise(optionName: string) {

    this.cells.forEach((cell: CellProperties) => {

      if (!this.isImpact) {
        cell.isImpact = false;
      }

      if (!this.isLikelihood) {
        cell.isLikelihood = false;
      }

      if (!this.isRelationshipIcons) {
        cell.isInputRelationship = false;
        cell.isOutputRelationship = false;
      }

      if (!(this.isInputRelationship || this.isRelationshipIcons)) {
        cell.isInputRelationship = false;
      }

      if (!(this.isOutputRelationship || this.isRelationshipIcons)) {
        cell.isOutputRelationship = false;
      }
    })

    this.deleteShapes(optionName);
  }

  // To remove a particular influence: such as influence by (or input cells)
  removeShapesInfluenceWise(influenceType: string) {

    if (influenceType.includes('Input')) {
      this.setInputCellsToFalse(this.referenceCell.inputCells, 3);
    }

    if (influenceType.includes('Output')) {
      this.setOutputCellsToFalse(this.referenceCell.outputCells, 3);
    }

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

  private removeSecondDegreeInputNeighbours() {

    let names = new Array<string>();
    this.referenceCell.inputCells.forEach((inCell: CellProperties) => {
      inCell.inputCells.forEach((inincell: CellProperties) => {
        this.setInputPropertiesToFalse(inincell);
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
          this.setInputPropertiesToFalse(ininincell);
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
        this.setOutputPropertiesToFalse(outoutcell);
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
          this.setOutputPropertiesToFalse(outoutoutCell);
          names.push(outoutoutCell.address);
        })
      })
    })
    this.deleteShapesInCells(names);
  }

  private setInputCellsToFalse(cells: CellProperties[], n: number) {

    cells.forEach((cell: CellProperties) => {
      this.setInputPropertiesToFalse(cell);

      if (n == 1) {
        return;
      }
      this.setInputCellsToFalse(cell.inputCells, n - 1);
    })
  }

  private setOutputCellsToFalse(cells: CellProperties[], n: number) {

    cells.forEach((cell: CellProperties) => {
      this.setOutputPropertiesToFalse(cell);

      if (n == 1) {
        return;
      }
      this.setOutputCellsToFalse(cell.outputCells, n - 1);
    })
  }


  private setPropertiesToFalse(cell: CellProperties) {

    cell.isImpact = false;
    cell.isLikelihood = false;
    cell.isSpread = false;
  }

  private setInputPropertiesToFalse(cell: CellProperties) {
    this.setPropertiesToFalse(cell);
    cell.isInputRelationship = false;
  }


  private setOutputPropertiesToFalse(cell: CellProperties) {
    this.setPropertiesToFalse(cell);
    cell.isOutputRelationship = false;
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