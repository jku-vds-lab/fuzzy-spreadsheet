/* global console, Excel */
import CellProperties from './cellproperties';
import Impact from './operations/impact';
import Likelihood from './operations/likelihood';
import Spread from './operations/spread';
import Relationship from './operations/relationship';
import SheetProperties from './sheetproperties';
import DiscreteSpread from './operations/spread';

export default class CellOperations {

  private cells: CellProperties[];
  private referenceCell: CellProperties;
  private degreeOfNeighbourhood: number = 1;
  private impact: Impact;
  private likelihood: Likelihood;
  private spread: Spread;
  private relationship: Relationship;

  constructor(cells: CellProperties[], referenceCell: CellProperties, n: number) {
    this.cells = cells;
    this.referenceCell = referenceCell;
    this.degreeOfNeighbourhood = n;
    this.impact = new Impact(this.referenceCell, this.cells);
    this.likelihood = new Likelihood(this.cells, this.referenceCell);
    this.spread = new Spread(this.cells, null, this.referenceCell);
    this.relationship = new Relationship(this.cells, this.referenceCell);
  }

  getCells() {
    return this.cells;
  }

  getDegreeOfNeighbourhood() {
    return this.degreeOfNeighbourhood;
  }

  showInputImpact(n: number) {
    this.impact.showInputImpact(n);
  }


  showOutputImpact(n: number) {
    this.impact.showOutputImpact(n);
  }

  removeInputImpact(n: number) {
    this.impact.removeInputImpact(n);
  }

  removeOutputImpact(n: number) {
    this.impact.removeOutputImpact(n);
  }

  removeAllImpacts() {
    this.impact.removeAllImpacts();
  }

  addLikelihoodInfo() {
    this.likelihood.addLikelihoodInfo();
  }

  showInputLikelihood(n: number) {
    this.likelihood.showInputLikelihood(n);
  }

  showOutputLikelihood(n: number) {
    this.likelihood.showOutputLikelihood(n);
  }

  removeInputLikelihood(n: number) {
    this.likelihood.removeInputLikelihood(n);
  }

  removeOutputLikelihood(n: number) {
    this.likelihood.removeOutputLikelihood(n);
  }

  removeAllLikelihoods() {
    this.likelihood.removeAllLikelihoods();
  }

  showSpread(n: number, isInput: boolean, isOutput: boolean) {
    this.spread.showSpread(n, isInput, isOutput);
  }

  removeSpread(isInput: boolean, isOutput: boolean, isRemoveAll: boolean) {
    this.spread.removeSpread(isInput, isOutput, isRemoveAll);
  }

  removeSpreadFromReferenceCell() {
    this.spread.removeSpreadFromReferenceCell();
  }

  showInputRelationship(n: number) {
    this.relationship.showInputRelationship(n);
  }

  removeInputRelationship() {
    this.relationship.removeInputRelationship();
  }

  showOutputRelationship(n: number) {
    this.relationship.showOutputRelationship(n);
  }

  removeOutputRelationship() {
    this.relationship.removeOutputRelationship();
  }

  deleteUpdateshapes() {

    Excel.run(function (context) {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      const oldTextbox = sheet.shapes;
      oldTextbox.load("items/name");
      return context.sync().then(function () {
        oldTextbox.items.forEach(function (c) {
          if (c.name.includes('Update'))
            c.delete();
          console.log('Deleted shape: ' + c.name);
        });
      });

    })
  }
}