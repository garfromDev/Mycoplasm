/**
Is called when the spreadsheet opens
Create the navigation menu and the sidebar
NOTE : the sidebar content is in sheet "sideBar.html"
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🛰 Stocks mycoplasme')
      .addItem('Etat des lieux', 'goToEtatDesLieux')
      .addItem('Stock prépa antigène', 'goToStockAg')
      .addItem('Commande diluant', 'goToCommandeDiluant')
      .addItem('Stock diluants', 'goToStockDiluant')
      .addItem('Simulation production AG', 'goToSimulationProductionAG')
      .addItem('Stock antigène coloré', 'goToStockAntigeneColores')
      .addSeparator()
      .addItem('Afficher la barre de navigation', 'showSidebar')
      .addToUi();
  showSidebar();

}


function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sideBar')
      .setTitle('Navigation')
      .setWidth(70); //il suffit de chanegr ce nombre pour adapter la largeur
  SpreadsheetApp.getUi().showSidebar(html);
}


function goToEtatDesLieux(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CST().ETAT_DES_LIEUX_NAME).activate();
}


function goToStockAg(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CST().STOCK_AG_NAME).activate();
}

function goToCommandeDiluant(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CST().COMMANDE_DILUANT_NAME).activate();
}

function goToStockDiluant(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CST().STOCK_DILUANT_NAME).activate();
}

function goToSimulationProductionAG(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CST().SIMULATION_PRODUCTION_AG_NAME).activate();
}

function goToStockAntigeneColores(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CST().STOCK_ANTIGENE_COLORE_NAME).activate();
}

