/* global Excel */

export default class CellProperties {
  public cell: string;
  public value: number;
  public top: number;
  public left: number;
  public height: number;
  public width: number;
  public isFocus: boolean;
  public degreeToFocus: number;
  public formula: any;
  public inputCells: Excel.Range[];
  public outputCells: string[];
  public inCells: CellProperties[];
  public outCells: CellProperties[];
  CellProperties() {
    this.cell = "";
    this.value = 0;
    this.top = 0;
    this.left = 0;
    this.height = 0;
    this.width = 0;
    this.isFocus = false;
    this.degreeToFocus = 0;
    this.formula = "";
    this.outputCells = new Array<string>();
  }
  async getCellProperties(cellAddress: string, focusCell: string, degreeToFocus: number) {
    this.cell = cellAddress;
    this.isFocus = false;
    if (cellAddress == focusCell) {
      this.isFocus = true;
      this.degreeToFocus = 0;
    }
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      const cell = sheet.getRange(cellAddress);
      cell.load(["values", "top", "left", "height", "width", "formulas"]);
      await context.sync();
      this.value = cell.values[0][0]; // gets the current cell value
      this.top = cell.top;
      this.left = cell.left;
      this.height = cell.height;
      this.width = cell.width;
      this.degreeToFocus = degreeToFocus;
      this.formula = cell.formulas[0][0]; // gets the formula of the current cell
      await context.sync();
    });
    return this;
  }
  async getCellValue(cellAddress: string) {
    let value: number = 0;
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Probability");
      const cell = sheet.getRange(cellAddress);
      cell.load("values");
      return context.sync().then(function () {
        value = cell.values[0][0]; // gets the current cell value
      });
    });
    return value;
  }
}