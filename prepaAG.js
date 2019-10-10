// appelé par le bouton "Archiver" sur la page stock prépa AG
function archiveAGLot(){
 // 1 récupére le n° de lot correspondant à la ligne sélectionnée
  var sheetToArchiveName = getValueForRange(CST().STOCK_AG_LOT_NO_TO_ARCHIVE);
  var lineToArchive = getValueForRange(CST().STOCK_AG_LINE_NO_OF_LOT_TO_ARCHIVE);
  
  archiveFromTable(CST().STOCK_AG_FORMULA_FROM_ZONE , lineToArchive, sheetToArchiveName);  
}



function createAGLot(){ // called by button <Créer le nouveau lot>
  stockAG = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CST().STOCK_AG_NAME);
  // 1 Vérifier que le n° de lot est valide

  if(!value("validAGLotNo")){
     alert("No de lot invalide");
     return;
   }
  // 2 Vérifier que le type n'est pas vide
  if(value("inputAGType").length <= 0){
     alert("Sélectionner un type d'antigène");
     return;
   }
  // 3 vérification que la ligne de référence n'est pas vide
  if(stockAG.getRange(CST().STOCK_AG_FORMULA_TO_ZONE).getFormulas()[0].length < 2){
     alert("Impossible de créer un lot, le tableau est corrompu");
     return;
  }
  // 4 création d'une feuille pour le lot
  toast("creation du lot " +myRange("inputLotNo").getValue() );
  var newLot = createAGLotWithTypeAndLotNo(value("inputAGType"), value("inputLotNo"));
  if(newLot == null ){
     alert("Ce lot existe déjà ou une erreur s'est produite");
     return;
   }
   // 5 Insérer une ligne dans le listing (toujours en 1ere position)
  insertLineInTable(CST().STOCK_AG_FORMULA_FROM_ZONE, CST().STOCK_AG_FORMULA_TO_ZONE);
  
   // 7 Ecrire les valeurs de type et le no de lot
  stockAG.getRange(CST().STOCK_AG_LOT_NO).setValue(value("inputLotNo"));
  stockAG.getRange(CST().STOCK_AG_TYPE).setValue(value("inputAGType"));
  
  // 8 Créer le lien hypertexte vers la feuille de lot AG
  var link = getLinkToSheet(newLot)
  addHyperlinkToCell(stockAG.getRange(CST().STOCK_AG_LOT_NO),
                     link);
  
  toast("Création du lot "+value("inputLotNo")+" effectuée avec succès");
}


//======== applicative functions =============
/**
* copy the template to a new sheet and insert the value
* return null if fails, the Sheet if succesfull
*/
function createAGLotWithTypeAndLotNo(type, lotNo){
  
  s = SpreadsheetApp.getActiveSpreadsheet();
  // 0 check if name already exist
  const newName = makeAGSheetName(type, lotNo)
  if (s.getSheetByName(newName) != null) {
    //throw "existing newName";
    return null; 
  }
  // 1 find template and make a copy 
  try{
    agTemplate = copyTemplateTo(s.getSheetByName(CST().PREPA_AG_TEMPLATE), newName);
  }catch(error){
   // throw "copy failed";
    return null;  
  }
  // 2 rename it
  try {
    agTemplate = agTemplate.setName(newName);
  }catch(error){
    s.deleteSheet(agTemplate);
    //throw "rename failed";
    return null;
  }
  // 3 set the value for lot and type 
  agTemplate.getRange(CST().AG_TEMPLATE_LOT_NO).setValue(lotNo);
  agTemplate.getRange(CST().AG_TEMPLATE_TYPE).setValue(type);
  
  return agTemplate;
}



// ==== CUSTOM FUNCTION ====== utilisable dans les formules de calcul

/**
*  Make the name for an antigene lot sheet 
* @param {string} type : the type of antigene (MM, MG, ...)
* @param {string} lotNr : the lot identifier, as "12.34.56"
* @return {string} : a valid name for the sheet, in the form of "AG_MM_12.34.56.78"
*/
function makeAGSheetName(type, lotNr){
  return "Prepa_AG_"+type+"_"+lotNr;  
}


function AGSheetName(type, noLot) { 
  return makeAGSheetName(type,noLot);
}


/**
*  Check validity of prepa AG lot number
* @param {string} lotNumber : the lot identifier, as "12.34.56"
* @return {string} : true if it is in a form of 3 double digit number separated by dot
*/
function isValidAGLotNo(lotNumber) {
  if ( typeof lotNumber != "string" ) { return false };
  return lotNumber.match(/\d+\.\d\d\.\d\d/) != null;
}



/* Note sur la navigation entre feuille
un lien hypertexte "#gid=4566" avec l'ID de la feuille
permet d'accéder à la feuille directement (sans ouvrir une nouvelle feuille)
mais ne marche pas si la feuille est masquée
*/

// Note : déclarer des constantes globales ne marche pas car chaque fonction 
// est exécutée par le serveur dans un nouveau contexte lorsqu'on clique sur
// un bouton, ce qui entraine une redéclaration de la constante
// => à la place, on utilise une fonction retournant un objet contenant les constantes (fichier Parametres)



//======== Temporary functions =============
function test(){
  s = SpreadsheetApp.getActiveSpreadsheet();
  a = s.getActiveSheet();
  alert(a.getName());
  alert(s.getSheetByName("a faire") === a); 
}

function testGetLink(){
 var s =  SpreadsheetApp.getActiveSpreadsheet()
 alert(getLinkToSheet(s.getSheetByName("a faire")));
}
