import SheetProp from './sheet/sheetproperties';
import WhatIfProps from './sheet/whatifproperties';

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global setTimeout, document, Office */


Office.initialize = () => {

  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("parseSheet").onclick = MainClass.parseSheet;
  document.getElementById("referenceCell").onclick = MainClass.markAsReferenceCell;
  document.getElementById("inputRelationship").onclick = MainClass.inputRelationship;
  document.getElementById("outputRelationship").onclick = MainClass.outputRelationship;
  document.getElementById("first").onchange = MainClass.first;
  document.getElementById("second").onchange = MainClass.second;
  document.getElementById("third").onchange = MainClass.third;
  document.getElementById("impact").onclick = MainClass.impact;
  document.getElementById("likelihood").onclick = MainClass.likelihood;
  document.getElementById("spread").onclick = MainClass.spread;
  document.getElementById("relationship").onclick = MainClass.relationshipIcons;
  document.getElementById("startWhatIf").onclick = MainClass.whatIf;
  document.getElementById("useNewValues").onclick = MainClass.keepNewValues;
  document.getElementById("dismissValues").onclick = MainClass.dismissNewValues;
}

class MainClass {

  public static sheetProp: SheetProp = new SheetProp();
  public static whatIfProp: WhatIfProps;
  public static isWhatIfStarted: boolean = false;

  public static parseSheet() {
    MainClass.sheetProp.parseSheet();
  }

  public static markAsReferenceCell() {
    MainClass.sheetProp.markAsReferenceCell();
    setTimeout(() => MainClass.whatIfProp = new WhatIfProps(MainClass.sheetProp.getCells(), MainClass.sheetProp.getReferenceCell()), 1000);
  }

  public static inputRelationship() {
    MainClass.sheetProp.inputRelationship();

    if (MainClass.isWhatIfStarted) {
      MainClass.whatIfProp.inputRelationship();
    }
  }

  public static outputRelationship() {
    MainClass.sheetProp.outputRelationship();

    if (MainClass.isWhatIfStarted) {
      MainClass.whatIfProp.outputRelationship();
    }
  }

  public static first() {
    MainClass.sheetProp.setDegreeOfNeighbourhood(1);

    if (MainClass.isWhatIfStarted) {
      setTimeout(() => MainClass.whatIfProp.setDegreeOfNeighbourhood(1), 1000);
    }
  }

  public static second() {
    MainClass.sheetProp.setDegreeOfNeighbourhood(2);

    if (MainClass.isWhatIfStarted) {
      setTimeout(() => MainClass.whatIfProp.setDegreeOfNeighbourhood(2), 1000);
    }
  }

  public static third() {
    MainClass.sheetProp.setDegreeOfNeighbourhood(3);

    if (MainClass.isWhatIfStarted) {
      setTimeout(() => MainClass.whatIfProp.setDegreeOfNeighbourhood(3), 1000);
    }
  }

  public static impact() {
    MainClass.sheetProp.impact();

    if (MainClass.isWhatIfStarted) {
      MainClass.whatIfProp.impact();
    }
  }

  public static likelihood() {
    MainClass.sheetProp.likelihood();

    if (MainClass.isWhatIfStarted) {
      MainClass.whatIfProp.likelihood();
    }
  }

  public static relationshipIcons() {
    MainClass.sheetProp.relationshipIcons();

    if (MainClass.isWhatIfStarted) {
      MainClass.whatIfProp.relationshipIcons();
    }
  }

  public static spread() {
    MainClass.sheetProp.spread();

    if (MainClass.isWhatIfStarted) {
      setTimeout(() => MainClass.whatIfProp.spread(), 1000);
    }
  }


  public static whatIf() {
    MainClass.isWhatIfStarted = true;
    MainClass.whatIfProp.showUIOptionsForWhatIf();
    setTimeout(() => {
      MainClass.whatIfProp = new WhatIfProps(MainClass.sheetProp.getCells(), MainClass.sheetProp.getReferenceCell());
      MainClass.whatIfProp.startWhatIf();
    }, 1000);
  }

  public static dismissNewValues() {
    MainClass.isWhatIfStarted = false;
    MainClass.whatIfProp.dismissNewValues();
    setTimeout(() => MainClass.sheetProp.keepOldValues(), 1000);
  }

  public static keepNewValues() {
    MainClass.isWhatIfStarted = false;
    MainClass.whatIfProp.keepNewValues();
    setTimeout(() => MainClass.sheetProp.processNewValues(), 1000);
  }
}