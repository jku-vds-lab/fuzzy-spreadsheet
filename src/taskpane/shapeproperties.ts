export default class ShapeProperties {
  shapeType: string;
  color: string;
  transparency: number;
  height: number;
  width: number;
  getShapeProperties(shapeType: string, color: string, transparency: number, height: number, width: number) {
    this.shapeType = shapeType;
    this.color = color;
    this.transparency = transparency;
    this.height = height;
    this.width = width;
    return this;
  }
}