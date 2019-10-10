function PrintFORM() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('O2:X58').activate();
};

function EditFORM() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('X8').activate();
};

function HautFORM() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('T8').activate();
};

function LancementAutocrat() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('Y:Y').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('Y2'));
  spreadsheet.getActiveSheet().showColumns(25, 21);
  spreadsheet.getRange('AQ:AT').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('Z:AT').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AT1'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('Y:Y').activate();
  AutoCrat3.onStart();
};

function ReinitPrepaAG() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('S:S').activate();
  spreadsheet.getActiveSheet().showColumns(17, 4);
  spreadsheet.getRange('Q7:R').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('S:AH').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('S:AH').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('S7'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('S:S').activate();
  spreadsheet.getRange('Q7').activate();
};

function ReinitDiluant() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('Q7:R27').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('S:AG').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('S:AG').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('S7'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('S:S').activate();
  spreadsheet.getRange('Q7').activate();
};