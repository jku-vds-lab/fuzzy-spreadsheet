import CellOperations from "./celloperations";
import CellProperties from "./cellproperties";

export default class SheetProperties {

  static isInputRelationship: boolean = false;
  static isOutputRelationship: boolean = false;
  static isRelationship: boolean = false;
  static isImpact: boolean = false;
  static isLikelihood: boolean = false;
  static isSpread: boolean = false;
  static isReferenceCell: boolean = false;
  static degreeOfNeighbourhood: number = 1;
  static isCheatSheetExist: boolean = false;
  static cellOp: CellOperations;
  static cellProp = new CellProperties();
  static cells: CellProperties[];
  static newCells: CellProperties[] = null;
  static referenceCell: CellProperties = null;
  static isSheetParsed = false;
  static newValues: any[][];
  static newFormulas: any[][];
  static originalTopBorder: Excel.RangeBorder;
  static originalBottomBorder: Excel.RangeBorder;
  static originalLeftBorder: Excel.RangeBorder;
  static originalRightBorder: Excel.RangeBorder;
}