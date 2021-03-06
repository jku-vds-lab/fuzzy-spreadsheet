/* global setTimeout, console, Excel */
import CellProperties from "../cell/cellproperties";
import Impact from "./impact";
import Likelihood from "./likelihood";

export default class CommonOperations {
  private referenceCell: CellProperties;
  private cells: CellProperties[];
  private isImpact: boolean;
  private isLikelihood: boolean;
  private isRelationshipIcons: boolean;
  private isSpread: boolean;
  private isInputRelationship: boolean;
  private isOutputRelationship: boolean;
  private isDelete: boolean;

  constructor(referenceCell: CellProperties, isDelete: boolean = true) {
    this.referenceCell = referenceCell;
    this.isImpact = false;
    this.isLikelihood = false;
    this.isRelationshipIcons = false;
    this.isSpread = false;
    this.isInputRelationship = false;
    this.isOutputRelationship = false;
    this.isDelete = isDelete;
  }

  drawRectangle(cells: CellProperties[], name: string) {

    try {
      Excel.run((context) => {

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        let MARGIN = 15;
        let height = 5;
        let width = 5;
        let top = 0;
        let left = 0;

        cells.forEach((cell: CellProperties) => {

          cell.rect = sheet.shapes.addGeometricShape("Rectangle");
          cell.rect.name = cell.address + name;

          let color = '#d9d9d9';
          let borderColor = 'gray';
          let transparency = 0;

          if (cell.isImpact) {
            color = cell.rectColor;

            if (cell.isImpactPositive) {
              borderColor = '#2166ac';
            } else {
              borderColor = '#b2182b';
            }
          }

          if (cell.isLikelihood) {
            height = cell.likelihood * 10;
            width = cell.likelihood * 10;
          }

          top = cell.top + (cell.height - height) / 2; // cell.height is 15
          left = cell.left + MARGIN + (10 - width) / 2;

          cell.rect.height = height;
          cell.rect.width = width;
          cell.rect.top = top;
          cell.rect.left = left;

          cell.rect.geometricShapeType = Excel.GeometricShapeType.rectangle;
          cell.rect.fill.setSolidColor(color);
          cell.rect.fill.transparency = transparency;
          cell.rect.lineFormat.weight = 1;
          cell.rect.lineFormat.color = borderColor;
        })

        let range = sheet.getRange(this.referenceCell.address);
        range.select();
        return context.sync();
      });
    } catch (error) {
      console.log('---' + name + ' : ', error);
    }
  }

  setCells(cells: CellProperties[]) {
    this.cells = cells;
  }

  setOptions(isImpact: boolean, isLikelihood: boolean, isSpread: boolean, isInputRelationship: boolean, isOutputRelationship: boolean) {
    this.isImpact = isImpact;
    this.isLikelihood = isLikelihood;
    this.isSpread = isSpread;
    this.isInputRelationship = isInputRelationship;
    this.isOutputRelationship = isOutputRelationship;
  }

  // To remove shapes from reference cell
  removeShapesReferenceCellWise() {
    this.referenceCell.isSpread = false;
    if (this.isDelete) {
      this.deleteShapes('Reference');
    }
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

      if (!this.isSpread) {
        cell.isSpread = false;
      }

      if (!this.isInputRelationship) {
        cell.isInputRelationship = false;
      }

      if (!this.isOutputRelationship) {
        cell.isOutputRelationship = false;
      }
    })

    if (this.isDelete) {
      this.deleteShapes(optionName);
    }
  }

  // To remove a particular influence: such as influence by (or input cells)
  removeShapesInfluenceWise(influenceType: string) {

    if (influenceType.includes('Input')) {
      this.setInputCellsToFalse(this.referenceCell.inputCells, 3);
    }

    if (influenceType.includes('Output')) {
      this.setOutputCellsToFalse(this.referenceCell.outputCells, 3);
    }

    if (this.isDelete) {
      this.deleteShapes(influenceType);
    }
  }

  // To remove updated shapes
  removeShapesUpdatedWise() {
    this.deleteShapes('Update');
  }

  // To remove all the shapes when a reference cell is changed
  removeAllShapes() {
    this.cells.forEach((cell: CellProperties) => {
      this.setInputPropertiesToFalse(cell);
      this.setOutputPropertiesToFalse(cell);
    });
    if (this.isDelete) {
      this.deleteShapes('Task');
    }
  }

  removeSpreadCellWise(cells: CellProperties[], name: string) {
    let names = new Array<string>();

    cells.forEach((cell: CellProperties) => {
      cell.isSpread = false;
      names.push(cell.address + name);
    });

    this.deleteShapesInCells(names);
  }

  // To remove shapes based of degree of neighbourhood
  removeShapesNeighbourWise(n: number) {


    if (n == 0) {
      this.removeShapesInfluenceWise('Input');
      this.removeShapesInfluenceWise('Output');
    }

    if (n == 1) {
      this.removeThirdDegreeInputNeighbours();
      setTimeout(() => this.removeSecondDegreeInputNeighbours(), 1000);

      this.removeThirdDegreeOutputNeighbours();
      setTimeout(() => this.removeSecondDegreeOutputNeighbours(), 1000);
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
    if (this.isDelete) {
      this.deleteShapesInCells(names);
    }
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
    if (this.isDelete) {
      this.deleteShapesInCells(names);
    }
  }

  private removeSecondDegreeOutputNeighbours() {
    let names = new Array<string>();
    this.referenceCell.outputCells.forEach((outCell: CellProperties) => {
      outCell.outputCells.forEach((outoutcell: CellProperties) => {
        this.setOutputPropertiesToFalse(outoutcell);
        names.push(outoutcell.address);
      })
    })
    if (this.isDelete) {
      this.deleteShapesInCells(names);
    }
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
    if (this.isDelete) {
      this.deleteShapesInCells(names);
    }
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