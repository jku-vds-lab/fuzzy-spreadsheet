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
      MainClass.whatIfProp.setDegreeOfNeighbourhood(1);
    }
  }

  public static second() {
    MainClass.sheetProp.setDegreeOfNeighbourhood(2);

    if (MainClass.isWhatIfStarted) {
      MainClass.whatIfProp.setDegreeOfNeighbourhood(2);
    }
  }

  public static third() {
    MainClass.sheetProp.setDegreeOfNeighbourhood(3);

    if (MainClass.isWhatIfStarted) {
      MainClass.whatIfProp.setDegreeOfNeighbourhood(3);
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
      MainClass.whatIfProp.spread();
    }
  }


  public static whatIf() {
    MainClass.isWhatIfStarted = true;
    MainClass.whatIfProp = new WhatIfProps(MainClass.sheetProp.getCells(), MainClass.sheetProp.getReferenceCell());
    MainClass.whatIfProp.startWhatIf();
  }

  public static dismissNewValues() {
    MainClass.isWhatIfStarted = false;
    MainClass.whatIfProp.dismissNewValues();
  }

  public static keepNewValues() {
    MainClass.isWhatIfStarted = false;
    MainClass.whatIfProp.keepNewValues();
    setTimeout(() => MainClass.sheetProp.processNewValues(), 1000);
  }
}


// async function startWhatIf() {
//   try {
//     (<HTMLInputElement>document.getElementById("startWhatIf")).disabled = true;
//     document.getElementById('useNewValues').hidden = true;
//     document.getElementById('dismissValues').hidden = true;
//     performWhatIf();
//     document.getElementById('useNewValues').hidden = false;
//     document.getElementById('dismissValues').hidden = false;
//     (<HTMLInputElement>document.getElementById("startWhatIf")).disabled = false;
//   } catch (error) {
//     console.log(error);
//   }
// }
// var eventResult;

// function performWhatIf() {
//   Excel.run(function (context) {
//     var worksheet = context.workbook.worksheets.getActiveWorksheet();
//     console.log('Worksheet has changed');
//     eventResult = worksheet.onChanged.add(processWhatIf); // onCalculated

//     return context.sync()
//       .then(function () {
//         console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
//       });
//   }).catch((reason: any) => { console.log(reason) });
// }


// async function processWhatIf() {

//   if (SheetProperties.referenceCell == null) {
//     console.log('Returning because reference cell is null');
//     return;
//   }

//   await Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getActiveWorksheet();
//     const range = sheet.getUsedRange(true);
//     range.load(['formulas', 'values']);
//     await context.sync();
//     SheetProperties.newValues = range.values;
//     SheetProperties.newFormulas = range.formulas;
//   });


//   // let x = await SheetProperties.cellProp.getCellsFormulasValues();
//   // // eslint-disable-next-line require-atomic-updates
//   // SheetProperties.newValues = x.values;
//   // // eslint-disable-next-line require-atomic-updates
//   // SheetProperties.newFormulas = x.formulas;
//   // console.log('Original' + SheetProperties.cells[0].value);
//   // console.log('New' + SheetProperties.newValues[0]);

//   // eslint-disable-next-line require-atomic-updates
//   SheetProperties.newCells = SheetProperties.cellProp.updateNewValues(SheetProperties.newValues, SheetProperties.newFormulas);

//   const whatif = new WhatIf(SheetProperties.newCells, SheetProperties.cells, SheetProperties.referenceCell);

//   whatif.calculateChange();

//   SheetProperties.cellOp.deleteUpdateshapes();

//   whatif.showUpdateTextInCells(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);

//   // if (SheetProperties.isSpread) {
//   //   whatif.showNewSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
//   // }
// }


// // To be fixed!!
// async function useNewValues() {
//   try {
//     document.getElementById('useNewValues').hidden = true;
//     document.getElementById('dismissValues').hidden = true;
//     removeHandler();
//     removeHtmlSpreadInfoForOriginalChart();
//     removeHtmlSpreadInfoForNewChart();
//     removeAllShapes();
//     SheetProperties.newCells = null;
//     SheetProperties.cellProp = new CellProperties();
//     // eslint-disable-next-line require-atomic-updates
//     SheetProperties.cells = await SheetProperties.cellProp.getCells();
//     SheetProperties.cellProp.getRelationshipOfCells();
//     // eslint-disable-next-line require-atomic-updates
//     SheetProperties.referenceCell = SheetProperties.cellProp.getReferenceAndNeighbouringCells(SheetProperties.referenceCell.address);
//     SheetProperties.cellProp.checkUncertainty(SheetProperties.cells);
//     // eslint-disable-next-line require-atomic-updates
//     SheetProperties.cellOp = new CellOperations(SheetProperties.cells, SheetProperties.referenceCell, 1);
//     // eslint-disable-next-line require-atomic-updates
//     SheetProperties.isReferenceCell = true;
//     displayOptions();
//   } catch (error) {
//     console.log(error);
//   }
// }


// async function dismissValues() {

//   try {
//     document.getElementById('useNewValues').hidden = true;
//     document.getElementById('dismissValues').hidden = true;
//     console.log('Remove Event Handler');

//     removeHandler();

//     if (SheetProperties.isSpread) {
//       const whatif = new WhatIf(SheetProperties.newCells, SheetProperties.cells, SheetProperties.referenceCell);
//       whatif.deleteNewSpread(SheetProperties.degreeOfNeighbourhood, SheetProperties.isInputRelationship, SheetProperties.isOutputRelationship);
//       removeHtmlSpreadInfoForNewChart();
//     }

//     SheetProperties.cellOp.deleteUpdateshapes();

//     SheetProperties.newCells = null;

//     await Excel.run(async (context) => {
//       const sheet = context.workbook.worksheets.getActiveWorksheet();
//       let cellRanges = new Array<Excel.Range>();
//       let cellValues = new Array<number>();
//       let cellFormulas = new Array<any>();

//       SheetProperties.cells.forEach((cell: CellProperties) => {

//         let range = sheet.getRange(cell.address);
//         cellRanges.push(range.load(['values', 'formulas']));
//         cellValues.push(cell.value);

//         let formula = cell.formula;
//         if (formula == "") {
//           formula = cell.value.toString();
//         }
//         cellFormulas.push(formula);
//       })

//       await context.sync();

//       let i = 0;

//       cellRanges.forEach((cellRange: Excel.Range) => {
//         cellRange.values = [[cellValues[i]]];
//         cellRange.formulas = [[cellFormulas[i]]];
//         i++;
//       })
//     });

//   } catch (error) {
//     console.log('Error: ', error);
//   }
// }

// function removeHandler() {
// return Excel.run(eventResult.context, function (context) {
//   eventResult.remove();

//   return context.sync()
//     .then(function () {
//       eventResult = null;
//       console.log("Event handler successfully removed.");
//     });
// }).catch((reason: any) => { console.log(reason) });
// }
