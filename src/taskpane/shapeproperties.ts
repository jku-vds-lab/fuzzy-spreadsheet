import CellProperties from "./cellproperties";

export default class ShapeProperties {
  cell: CellProperties;
  shapeType: string;
  color: string;
  transparency: number;
  height: number;
  width: number;
  setShapeProperties(cell: CellProperties, shapeType: string, color: string, transparency: number, height: number, width: number) {
    this.cell = cell;
    this.shapeType = shapeType;
    this.color = color;
    this.transparency = transparency;
    this.height = height;
    this.width = width;
    return this;
  }
}