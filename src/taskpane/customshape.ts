import CellProperties from "./cellproperties";

/* global Excel */

export default interface CustomShape {
  shape: Excel.Shape;
  cell: CellProperties;
  color: string;
  transparency: number;
}