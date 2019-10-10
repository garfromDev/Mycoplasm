/**
* Appellé par le bouton de la feuille stock diluant
* "nouvelle commande"
* must be called from "Commande diluant" sheet
*/
function recordDiluantOrder(){
  var stockDiluant = activeSheet();
  
  // 1) vérifier les données d'entrée
  if(! isValidInputDiluant()){
    alert("Saisisser un type de diluant et une date de commande valide");
    return;
  }
  var diluantType = getValueForRange(CST().USER_INPUT_DILUANT_TYPE);
  var diluantOrderdate = getValueForRange(CST().USER_INPUT_DILUANT_ORDER_DATE);
  var volume = getValueForRange(CST().USER_INPUT_DILUANT_VOLUME);
  var orderNo = createOrderNo(diluantType, diluantOrderdate);
  
  toast("création de la commande diluant <"+diluantType+"> du "+Utilities.formatDate(diluantOrderdate, getUserTimeZone(), "dd/MM/yyyy")+" en cours...");
  
  try{
    insertLineInTable(CST().DILUANT_ORDER_FORMULA_FROM_ZONE, CST().DILUANT_ORDER_FORMULA_TO_ZONE );
  }catch(error){ // in case of failure, we display a message and delete the created copy
    alert("Erreur lors de l'insertion d'une ligne au niveau de "+CST().DILUANT_FORMULA_FROM_ZONE+" <"+error.message+">");
    return;
  }
  
  //5) y écrire les données 
  var datas = [orderNo, diluantType, diluantOrderdate, volume, false];
  locations = [CST().DILUANT_ORDER_NO_LOCATION,
              CST().DILUANT_ORDER_TYPE_LOCATION,
              CST().DILUANT_ORDER_DATE_LOCATION,
              CST().DILUANT_ORDER_VOLUME_LOCATION,
              CST().DILUANT_ORDER_RECEPTION_STATUS];
  try{
    transferDataTo(datas, locations, stockDiluant);
  }catch (error) { // in case of failure, we display a message and delete the created copy and the line
    alert("Erreur lors de l'écriture des données dans le tableau <"+error.message+">");
    stockDiluant.deleteRow(Number(CST().DILUANT_FORMULA_TO_ZONE.substring(1,2)));
    return;
  }
  toast("commande de diluant "+diluantType+" datée du "+Utilities.formatDate(diluantOrderdate, getUserTimeZone(), "dd/MM/yyyy")+" pour un volume de "+volume+" crée avec succès!");
}




//======== applicative functions =============

function isValidInputDiluant(){
  return getValueForRange(CST().VALID_DILUANT_ORDER);
}


/**
* Order No is coimposed of incremental number for this date
* followed by date (dd-MM-yyyy) followed by type
* @param {String} diluantType
* @param {Date} diluantOrderdate (in GMT)
* @return {String} in the form of 2_2018-05-01_MS
* SIDE EFFECT : change value of user properties MYCOPLASM_ORDER_DATE and MYCOPLASM_ORDER_INCREMENTAL 
*/
function createOrderNo(diluantType, diluantOrderdate){
  var previousDate = PropertiesService.getUserProperties().getProperty("MYCOPLASM_ORDER_DATE");
  var newDate = Utilities.formatDate(diluantOrderdate, "GMT", "yyyy-MM-dd");
  var incremental = 1;
  if(newDate==previousDate){
    incremental = Number(PropertiesService.getUserProperties().getProperty("MYCOPLASM_ORDER_INCREMENTAL"))+1;
  }else{
    PropertiesService.getUserProperties().setProperty("MYCOPLASM_ORDER_DATE", newDate);
  }
  PropertiesService.getUserProperties().setProperty("MYCOPLASM_ORDER_INCREMENTAL", incremental);
  
  return incremental + "_" + newDate + "_" + diluantType;
}



//======== Temporary functions =============
