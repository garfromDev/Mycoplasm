// appellée depuis le bouton "Lancer la production" de simulation production AG
function prepareAntigeneColores() {
    
  // 2 vérifier la présence des données
  var type = value(CST().INPUT_ANTIGENE_TYPE);
  if(type == ""){
    alert("Choisir un type d'antigène colorés");
    return;
  }
  const isSGP = /SGP/.test(type);
  if( (! value(CST().ANTIGENE_COLORES_VALID_INPUTS)) ||
    ( isSGP && (! value(CST().ANTIGENE_COLORES_VALID_INPUTS_SGP)) )
  ){
      // for SGP, need both condition true
    alert("Vérifier la sélection des lots à utiliser");
    return;
  }
  
  const lotAntigene = getValueForRange(CST().ANTIGENE_COLORES_INPUT_LOT_TO_CREATE);
  if(lotAntigene=="" ){
    alert("Saisir le n° de lot d'antigène colorés à lancer en fabrication");
    return;
  }else if(isInvalidAntigeneColoresLot(lotAntigene)){
    alert("<"+lotAntigene+"> n'est pas un identifiant de lot valide");
    return;
  }
  
  // 3 Création du template et copie des données
  const templateName = templateNameForAntigeneColoreType(type);
  const sheetNameToCreate = createAntigeneColoresLotIdForType(type, lotAntigene);
  toast("création du lot dantigène colorés "+sheetNameToCreate+" en cours...");
  var datas = [type, lotAntigene,
              getValueForRange(CST().ANTIGENE_COLORES_VOL1_LOCATION),
              getValueForRange(CST().ANTIGENE_COLORES_VOL2_LOCATION),
              getValueForRange(CST().ANTIGENE_COLORES_VOL3_LOCATION),
              getValueForRange(CST().ANTIGENE_COLORES_FLACON1_LOCATION),
              getValueForRange(CST().ANTIGENE_COLORES_FLACON2_LOCATION),
              getValueForRange(CST().ANTIGENE_COLORES_FLACON3_LOCATION)
                       ];
  var newSheet;
  datas = addSelectedLotNumber(datas,myRange(CST().INPUT_LOT_SELECTION));
  var SGPStandardData=[];
  var locations = [CST().ANTIGENE_COLORES_TYPE_LOCATION_IN_TEMPLATE,
                     CST().ANTIGENE_COLORES_LOT_LOCATION_IN_TEMPLATE,
                     CST().ANTIGENE_COLORES_VOL1_LOCATION_IN_TEMPLATE,
                     CST().ANTIGENE_COLORES_VOL2_LOCATION_IN_TEMPLATE,
                     CST().ANTIGENE_COLORES_VOL3_LOCATION_IN_TEMPLATE,
                     CST().ANTIGENE_COLORES_FLACON1_LOCATION_IN_TEMPLATE,
                     CST().ANTIGENE_COLORES_FLACON2_LOCATION_IN_TEMPLATE,
                     CST().ANTIGENE_COLORES_FLACON3_LOCATION_IN_TEMPLATE,
                     CST().ANTIGENE_COLORES_LOT1_PREPA_AG_LOCATION_IN_TEMPLATE,
                     CST().ANTIGENE_COLORES_LOT2_PREPA_AG_LOCATION_IN_TEMPLATE,
                     CST().ANTIGENE_COLORES_LOT3_PREPA_AG_LOCATION_IN_TEMPLATE,
                     CST().ANTIGENE_COLORES_LOT4_PREPA_AG_LOCATION_IN_TEMPLATE];
  try{
    newSheet = CopyTemplateWithData(getSheet(templateName),
                                        sheetNameToCreate, datas, locations)
    if(isSGP){
      SGPStandardData = addSelectedLotNumber(SGPStandardData,
                                             myRange(CST().INPUT_LOT_SELECTION_SGP));
      var SGPLocations = [ CST().ANTIGENE_COLORES_LOT1_SGP_STANDARD_LOCATION_IN_TEMPLATE,
                          CST().ANTIGENE_COLORES_LOT2_SGP_STANDARD_LOCATION_IN_TEMPLATE,
                          CST().ANTIGENE_COLORES_LOT3_SGP_STANDARD_LOCATION_IN_TEMPLATE,
                          CST().ANTIGENE_COLORES_LOT4_SGP_STANDARD_LOCATION_IN_TEMPLATE];
      transferDataTo(SGPStandardData, SGPLocations, newSheet);
    }
  }catch(error){
      alert("La creation de la feuille "+sheetNameToCreate+" à partir du template "+ templateName +" a échoué : "+error.message);
      return;
  }
  
  
  // 4) revenir sur stock antigene colorés et insérer une ligne dans le tableau
  getSheet(CST().STOCK_ANTIGENE_COLORE_NAME).activate();
  try{
    insertLineInTable(CST().ANTIGENE_COLORES_FORMULA_FROM_ZONE, CST().ANTIGENE_COLORES_FORMULA_TO_ZONE );
  }catch(error){ // in case of failure, we display a message and delete the created copy
    alert("Erreur lors de l'insertion d'une ligne au niveau de "+CST().DILUANT_FORMULA_FROM_ZONE+" <"+error.message+">");
    newSheet.getParent().deleteSheet(newSheet);
    return;
  }
  
  //5) y écrire les données et le lien hypertexte
  locations = [CST().ANTIGENE_COLORES_TYPE_LOCATION,
              CST().ANTIGENE_COLORES_LOT_LOCATION];
  datas = [type, lotAntigene];
  try{
    transferDataTo(datas, locations, activeSheet());
    addHyperlinkToCell(activeSheet().getRange(CST().ANTIGENE_COLORES_LOT_LOCATION),
                     getLinkToSheet(newSheet));
  }catch (error) { // in case of failure, we display a message and delete the created copy and the line
    alert("Erreur lors de l'écriture des données dans le tableau <"+error.message+">");
    newSheet.getParent().deleteSheet(newSheet);
    activeSheet().deleteRow(getRowFromA1(CST().ANTIGENE_COLORES_FORMULA_TO_ZONE));
    return;
  }
  toast("création du lot d'antigène colorés "+sheetNameToCreate+" effectué avec succès...");
}






// appellée depuis le bouton "Mise à jour des stocks" de la feuille du lot d'antigène coloré (copie du template)
/*
AG Template :
Diluant 6 lots de T17 à U19
Prépa antigène 1 lot T21

Template SGP:
Prépa antigène SGP standard : 4 lots de T17 à U18
Prépa antigène SGP variant :  4 lots de T21 à U22
*/
function updateStockWithVolume(){
  // must not be called from template sheet
  if(isTemplate(activeSheet())){
    alert("Cette feuille est un template, ne pas la modifier");
    return;
  }
  
  if( checkStockAlreadyUpdated()){
    alert("Le stock a déjà été mis à jour!");
    return;    
  }
  
  const typePrepaAG = getValueForRange(CST().ANTIGENE_COLORES_TYPE_LOCATION_IN_TEMPLATE);
  const antigene_colores=activeSheet().getName();
  const isSGP = /SGP/.test(typePrepaAG);
  var volumeInvalid;
  //alert("updating stock step 1 - isSGP:"+isSGP);
  if(isSGP){ //SGP variant
    [lot1,lot2, lot3, lot4,  vol1, vol2, vol3, vol4] = 
                [ CST().ANTIGENE_COLORES_LOT1_PREPA_AG_LOCATION_IN_TEMPLATE,
                  CST().ANTIGENE_COLORES_LOT2_PREPA_AG_LOCATION_IN_TEMPLATE,
                  CST().ANTIGENE_COLORES_LOT3_PREPA_AG_LOCATION_IN_TEMPLATE,
                  CST().ANTIGENE_COLORES_LOT4_PREPA_AG_LOCATION_IN_TEMPLATE,
                  CST().ANTIGENE_COLORES_VOLUME1_USED_LOCATION_IN_TEMPLATE, 
                  CST().ANTIGENE_COLORES_VOLUME2_USED_LOCATION_IN_TEMPLATE, 
                  CST().ANTIGENE_COLORES_VOLUME3_USED_LOCATION_IN_TEMPLATE,
                  CST().ANTIGENE_COLORES_VOLUME4_USED_LOCATION_IN_TEMPLATE 
                       ].map(getValueForRange);
    volumeInvalid = isInvalidVolumeUsed(lot1, lot2, lot3, lot4, "", "",
                         vol1, vol2, vol3, vol4, 0, 0);
    
  }else{ //Lot prépa AG
      [lot1, vol1] = [ CST().ANTIGENE_COLORES_LOT1_PREPA_AG_LOCATION_IN_TEMPLATE,
                         CST().ANTIGENE_COLORES_VOLUME1_USED_LOCATION_IN_TEMPLATE    
                       ].map(getValueForRange);
    volumeInvalid = (lot1 != "" && vol1 <= 0) || lot1 == "";
  }
  // 1 check volume information
  if(volumeInvalid){
    var produit = isSGP ? "SGP variant" : "prépa antigène";
    alert("Les volumes de "+produit+" utilisés ne sont pas correctement renseignés");
    return;
  }
  //lotD1, .. contient directement le nom de la feuille du flacon de diluant 2_1_2018-10-04_Diluant MM5
  // ou le nO de lot de SGP Standard
  const [lotD1, lotD2, lotD3, lotD4, lotD5, lotD6,
         volD1, volD2, volD3, volD4, volD5, volD6] = [ CST().ANTIGENE_COLORES_USED_FLACON1_LOCATION_IN_TEMPLATE,
                                 CST().ANTIGENE_COLORES_USED_FLACON2_LOCATION_IN_TEMPLATE,
                                 CST().ANTIGENE_COLORES_USED_FLACON3_LOCATION_IN_TEMPLATE,               
                                 CST().ANTIGENE_COLORES_USED_FLACON4_LOCATION_IN_TEMPLATE,
                                 CST().ANTIGENE_COLORES_USED_FLACON5_LOCATION_IN_TEMPLATE,
                                 CST().ANTIGENE_COLORES_USED_FLACON6_LOCATION_IN_TEMPLATE,               
                                 CST().ANTIGENE_COLORES_DILUANT_VOLUME_USED1_LOCATION_IN_TEMPLATE,
                                 CST().ANTIGENE_COLORES_DILUANT_VOLUME_USED2_LOCATION_IN_TEMPLATE,
                                 CST().ANTIGENE_COLORES_DILUANT_VOLUME_USED3_LOCATION_IN_TEMPLATE,
                                 CST().ANTIGENE_COLORES_DILUANT_VOLUME_USED4_LOCATION_IN_TEMPLATE,
                                 CST().ANTIGENE_COLORES_DILUANT_VOLUME_USED5_LOCATION_IN_TEMPLATE,
                                 CST().ANTIGENE_COLORES_DILUANT_VOLUME_USED6_LOCATION_IN_TEMPLATE,
                                               ].map(getValueForRange);
  if(isInvalidVolumeUsed(lotD1, lotD2, lotD3, lotD4, lotD5, lotD6,
                         volD1, volD2, volD3, volD4, volD5, volD6)){  
    alert("Les volumes de diluant ou SGP standard utilisés ne sont pas correctement renseignés"); 
    return;
  }
  
  // 2 record prepa antigene lot usage or SGP variant for SGP
  toast("Enregistrement des volumes utilisés en cours");
  try{
    if(isSGP){ //SGP variant, il faut mapper le lot vers le nom de feuille complet incluant le type 
      recordDiluantUsage([lot1, lot2, lot3, lot4].map(
        function(l){return "Prepa_AG_SGP variant_"+l}),
                       [vol1, vol2, vol3, vol4],
                       antigene_colores, CST().PREPA_AG_TEMPLATE_VOLUME_USAGE_LOCATION);                                                
    }else{
      recordLotUsage(makeAGSheetName(typePrepaAG, lot1), vol1, antigene_colores);
    }
  }catch(error){
    alert("Erreur lors de la mise à jour des volumes pour "+lot1+" - "+error.message);
    return;
  }

  // 2 record diluant usage (or SGP Standard for SGP)
  // CAUTION : template for SGP must have SGP standard at the same position than diluant in non-SGP standard !!!! 
  try{   
    if(isSGP){
      recordDiluantUsage([lotD1, lotD2, lotD4, lotD5].map(
        function(l){return "Prepa_AG_SGP standard_"+l}),
                       [volD1, volD2, volD4, volD5],
                       antigene_colores, CST().PREPA_AG_TEMPLATE_VOLUME_USAGE_LOCATION);     
    }else{
      recordDiluantUsage([lotD1, lotD2, lotD3, lotD4, lotD5, lotD6],
                       [volD1, volD2, volD3, volD4, volD5, volD6],
                       antigene_colores, CST().DILUANT_TEMPLATE_VOLUME_USAGE_LOCATION);
    }
  }catch(error){
    produit = isSGP ? "SGP standard" : "diluant";
    alert("Erreur lors de la mise à jour des volumes de "+produit+" - "+error.message);
    return;
  }
  
  activeSheet().getRange(CST().PREPA_AG_TEMPLATE_STOCK_UPDATE_LOCATION).setValue(new Date())
  toast("Enregistrement des volumes utilisés terminé avec succès");
}



// appellée depuis le bouton du stock antigene colorés
/*
 Archive le lot d'antigene coloré de la ligne sur laquelle est le curseur dans le stock
*/
function archiveAntigeneColores(){
  archiveFromTable(CST().ANTIGENE_COLORES_FORMULA_FROM_ZONE,
                   getValueForRange(CST().ANTIGENE_COLORES_LINE_TO_ARCHIVE),
                   getValueForRange(CST().ANTIGENE_COLORES_LOT_TO_ARCHIVE) );
}  
  


// ============== Applicative FUnction ========================
// true if contains a number, a dot, a 2 digit number, a dot, a 2 digit number
function isInvalidAntigeneColoresLot(lotAntigene){
  return !(/^\d+\.\d\d\.\d\d/.test(lotAntigene));
}

function checkStockAlreadyUpdated(){
  var past = new Date("2010/06/06");
  var stockUpdateDate = getValueForRange(CST().PREPA_AG_TEMPLATE_STOCK_UPDATE_LOCATION)
  return stockUpdateDate.valueOf() > past.valueOf()
}

function recordLotUsage(lot, volume, usedByLot){
  //alert("record usage of lot "+lot+" with volume "+volume+" from lot "+usedByLot);
  const targetSheet = getSheet(lot);
  if(targetSheet==null){
    throw new Error("lot " + lot + " non trouvé");
  }
  writeLotUsage(targetSheet, volume, usedByLot,CST().PREPA_AG_TEMPLATE_VOLUME_USAGE_LOCATION );
}


/**
* write the volume of diluant used for each flask at which date in the diluant flask sheet
* @param {[String]} flacons : array of flask reference (full name, 2_1_2018-10-01_dilunat MM5
* @param {[Float]} volumes : the value that must be written (volume information) for each flask
* @param {String} usedByLot : the identification of the antigene colores lot that has used this volume
* NOTE : no protection against overflow, may continue to write data lower than table end
*/
function recordDiluantUsage(flacons, volumes, usedByLot, usageLocation ){
  //alert("record diluant of "+flacons.length+" diluants used by "+usedByLot);

  for(i=0;i<flacons.length;i++){
    //alert("flacon no "+i+" : "+flacons[i]+" volume "+volumes[i]);
    if(flacons[i]==""){ continue;}
    targetSheet = null;
    try {
      targetSheet = getSheet(flacons[i]);
    }catch(e){
      continue; //on ignore l'erreur et on passe au lot suivant
    }
     if(targetSheet==null){
       continue; //feuille inexistante possible car pour SGP le nom associé à un lot inexistant n'est pas vide
        //throw new Error("Feuille " + flacons[i] + " non trouvée pour mise à jour du stock - utilisation par le lot "+usedByLot+" pour un volume de "+volumes[i]);
     }
     writeLotUsage(targetSheet, volumes[i], usedByLot, usageLocation);
  } // end for
  
}


/**
* write the volume used at which date in the target position of the target sheet
* @param {Sheet} targetSheet : the sheet into which must be written
* @param {Float} vol : the value that must be written (volume information)
* @param {String} usedByLot : the identification of the antigene colores lot that has used this volume
* @param {String} where : A1 notation of range where data must be written at first empty line
* NOTE : no protection against overflow, may continue to write data lower than table end
*/
function writeLotUsage(targetSheet, vol, usedByLot, where){

  var originalProtection = unprotectSheet(targetSheet);
  // find first empty position 
  var r = targetSheet.getRange(where);
  while(r.getValue() != ""){
    r=r.offset(1,0);
  }
  // write the datas
 // alert("writeLotUsage of lot  "+usedByLot+" vol: "+vol);
  r.setValue(Utilities.formatDate(new Date(), getUserTimeZone() , "yyyy-MM-dd"));
  r.offset(0,1).setValue(usedByLot);
  r.offset(0,2).setValue(vol);
  
  restoreProtection(originalProtection);
}


/**
* return true when there is no lot defined, or a lot is defined without volume
* a lot is defined when the string is not empty
* volume is defined when > 0
* @param {String} lotx the lot number (as a string)
* @param {Float} volx the volume
* @returns {Bool} true if volume invalid
*/
function isInvalidVolumeUsed(lot1, lot2, lot3, lot4, lot5, lot6, vol1, vol2, vol3, vol4, vol5, vol6){
    return ((lot1 != "" && vol1 <= 0) || (lot2 != "" && vol2 <= 0) ||
           (lot3 != "" && vol3 <= 0) || (lot4 != "" && vol4 <= 0) ||
           (lot5 != "" && vol5 <= 0) || (lot6 != "" && vol6 <= 0) ) ||
           ( lot1 == "" && lot2 == "" && lot3 == "" && lot4 == "" && lot5 == "" && lot6 == "") ;
}
 

/**
* @param {String} type
* @return {String} sheetName for template according type
*/
function templateNameForAntigeneColoreType(type){
 return "AG Template_"+type 
}



/**
* @param {String} type
* @return {String} lot ID according current date and type
*/
function createAntigeneColoresLotIdForType(type, lot){
  return "AG_" + type+"_"+lot;
}


/**
* @param {Range} inputs
* @return {Bool} true if there is a line with box checked and at least one column is empty
*/
function lotDataMissing(inputs){
  for(l=0;l<inputs.getNumRows();l++){
    if(inputs.getCell(l,4).getValue() && (
       inputs.getCell(l,1).getValue() =="" ||
       inputs.getCell(l,2).getValue() =="" ||
         inputs.getCell(l,3).getValue() ==""))
         {
           return true;
         }
  }
  return false;  
}


/**
* @param {[]} datas
* @param {Range} inputs
* @return {[]} array completed with value of first column of the line with box checked 
*/
function addSelectedLotNumber(datas,inputs){
  const checkCol = inputs.getNumColumns();
  for(l=1;l<=inputs.getNumRows();l++){
    if(inputs.getCell(l,checkCol).getValue()){
      datas.push(inputs.getCell(l,1).getValue());
    }
  }
  return datas;
}


//======================== PROV =========================
function test_ID(){
  alert(createAntigeneColoresLotIdForType("MG"));
}

function test(){
  var fromFormula = CST().ANTIGENE_COLORES_FORMULA_FROM_ZONE;
  alert(Number(fromFormula.substring(5)));
}

function testDate(){
  var d = new Date();
  alert(Utilities.formatDate(new Date(), getUserTimeZone() , "yyyy-MM-dd  HH:mm"));
}

function testGetValue(){
 Logger.log(getValueForRange("D13")); 
}

function testWriteLotUsage(){
  writeLotUsage(getSheet("Prepa_AG_MG_68.05.18"), 50, "68.05.18", CST().PREPA_AG_TEMPLATE_VOLUME_USAGE_LOCATION);
}

function testThrow(){
  try{
    throw new Error("toto");
  }catch(error){
    alert(error.message);
  }
}

function testaddSelectedLotNumber(){
 var datas=[]; 
 datas = addSelectedLotNumber(datas,myRange(CST().INPUT_LOT_SELECTION)); 
 Logger.log(datas);
}

