/* global console, Excel */
import CellProperties from './cellproperties';
import Impact from '../operations/impact';
import Likelihood from '../operations/likelihood';
import Spread from '../operations/spread';
import Relationship from '../operations/relationship';
import CommonOperations from '../operations/commonops';

export default class CellOperations {

  private cells: CellProperties[];
  private referenceCell: CellProperties;
  private degreeOfNeighbourhood: number = 1;
  private impact: Impact;
  private likelihood: Likelihood;
  private spread: Spread;
  private relationship: Relationship;
  private commonOps: CommonOperations;

  constructor(cells: CellProperties[], referenceCell: CellProperties, n: number) {
    this.cells = cells;
    this.referenceCell = referenceCell;
    this.degreeOfNeighbourhood = n;
    this.impact = new Impact(this.referenceCell);
    this.likelihood = new Likelihood(this.referenceCell);
    this.spread = new Spread(this.referenceCell);
    this.relationship = new Relationship(this.referenceCell);
    this.commonOps = new CommonOperations(this.referenceCell);
  }

  getCells() {
    return this.cells;
  }

  getDegreeOfNeighbourhood() {
    return this.degreeOfNeighbourhood;
  }


  setCells(cells: CellProperties[]) {
    this.cells = cells;
    this.commonOps.setCells(this.cells);
  }
  setOptions(isImpact: boolean, isLikelihood: boolean, isRelationshipIcons: boolean, isSpread: boolean, isInputRelationship: boolean, isOutputRelationship: boolean) {
    this.commonOps.setOptions(isImpact, isLikelihood, isRelationshipIcons, isSpread, isInputRelationship, isOutputRelationship);
  }

  removeShapesReferenceCellWise() {
    this.commonOps.removeShapesReferenceCellWise();
  }

  removeShapesOptionWise(optionName: string) {
    this.commonOps.removeShapesOptionWise(optionName);
  }

  removeShapesInfluenceWise(influenceType: string) {
    this.commonOps.removeShapesInfluenceWise(influenceType);
  }

  removeShapesUpdatedWise() {
    this.commonOps.removeShapesUpdatedWise();
  }

  removeShapesNeighbourWise(n: number) {
    this.commonOps.removeShapesNeighbourWise(n);
  }

  removeAllShapes() {
    this.commonOps.removeAllShapes();
  }

  showInputImpact(n: number, isDraw: boolean) {
    this.impact.showInputImpact(n, isDraw);
  }
  showOutputImpact(n: number, isDraw: boolean) {
    this.impact.showOutputImpact(n, isDraw);
  }

  showInputLikelihood(n: number, isDraw: boolean) {
    this.likelihood.showInputLikelihood(n, isDraw);
  }

  showOutputLikelihood(n: number, isDraw: boolean) {
    this.likelihood.showOutputLikelihood(n, isDraw);
  }

  showSpread(n: number, isInput: boolean, isOutput: boolean) {
    this.spread.showSpread(n, isInput, isOutput);
  }

  showInputRelationship(n: number) {
    this.relationship.showInputRelationship(n);
  }

  showOutputRelationship(n: number) {
    this.relationship.showOutputRelationship(n);
  }
}