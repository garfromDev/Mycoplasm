/**
* Appellé par le bouton de la feuille stock diluant
* "reception de flacon"
* must be called from "Stock diluant" sheet
*/

function receptionFlaconDiluant(){
  if(!getValueForRange(CST().DILUANT_VALID_INPUT)){
   alert("Saisir le n° de commande, la date de réception et le volume pour réceptionner ce flacon");
   return; 
  }
  // get the input information
  var [ diluantType, orderNo, diluantOrderDate, flaconNo, receptionDate, volume] = [ 
      CST().INPUT_DILUANT_TYPE,CST().INPUT_DILUANT_ORDER_NO,CST().INPUT_DILUANT_ORDER_DATE,
      CST().DILUANT_FLACON_NO, CST().DILUANT_FLACON_RECEPTION_DATE, 
      CST().INPUT_DILUANT_VOLUME
  ].map(getValueForRange);
  // increment the flask value if same order
  if(flaconNo==""){
    flaconNo = 1;
  }else{
    flaconNo=Number(flaconNo)+1
  }

  // create name for the new sheet for the flacon
  var newName = diluantSheetName(orderNo, flaconNo)
  
  // 2) copier le template vers une nouvelle feuille
  toast("Création du flacon de diluant "+diluantType+" n° "+flaconNo +" pour un volume de "+volume);
  try{
    var newSheet = copyTemplateTo( //protection is handled in copyTemplateTo
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CST().DILUANT_TEMPLATE),
      newName);
  }catch(error) {
    alert("erreur lors de la copie de "+CST().DILUANT_TEMPLATE+" vers "+newName+" <"+error.message+">");
    return;
  }
    
  // 3) transférer les données vers la nouvelle feuille
  var datas = [ diluantType, diluantOrderDate, receptionDate, volume, flaconNo];
  var locations = [ CST().DILUANT_TYPE_LOCATION_IN_TEMPLATE,
                   CST().DILUANT_ORDER_DATE_LOCATION_IN_TEMPLATE,
                   CST().DILUANT_RECEPTION_DATE_LOCATION_IN_TEMPLATE,
                    CST().DILUANT_VOLUME_LOCATION_IN_TEMPLATE,
                    CST().DILUANT_FLACON_NO_LOCATION_IN_TEMPLATE
                  ];
  try{
    transferDataTo(datas, locations, newSheet); //protection is handled in transferDataTo()
  }catch(error) { // in case of failure, we display a message and delete the created copy
    alert("Erreur lors du transfert des données vers "+newName+" <"+error.message+">");
    newSheet.getParent().deleteSheet(newSheet);
    return;
  }
  
 
  // 4) insérer une ligne dans le tableau
  try{
    insertLineInTable(CST().DILUANT_FORMULA_FROM_ZONE, CST().DILUANT_FORMULA_TO_ZONE );
  }catch(error){ // in case of failure, we display a message and delete the created copy
    alert("Erreur lors de l'insertion d'une ligne au niveau de "+CST().DILUANT_FORMULA_FROM_ZONE+" <"+error.message+">");
    newSheet.getParent().deleteSheet(newSheet);
    return;
  }
  
  //5) y écrire les données et le lien hypertexte
  datas = [ orderNo, flaconNo];
  locations = [CST().DILUANT_ORDER_NO, CST().DILUANT_TABLE_FLACON_NO];
  try{
    transferDataTo(datas, locations, activeSheet());
    addHyperlinkToCell(activeSheet().getRange(CST().DILUANT_ORDER_NO_LOCATION),
                     getLinkToSheet(newSheet));
  }catch (error) { // in case of failure, we display a message and delete the created copy and the line
    alert("Erreur lors de l'écriture des données dans le tableau <"+error.message+">");
    newSheet.getParent().deleteSheet(newSheet);
    activeSheet().deleteRow(Number(CST().DILUANT_FORMULA_TO_ZONE.substring(1,2)));
    return;
  }
  toast("Flacon de diluant "+diluantType+" n° "+flaconNo +" pour un volume de "+volume+" réceptionné avec succès!");
}



// appelé par le bouton "Archiver" de la feuille stock diluant
function archiveDiluant(){
 // 1 récupére le diluant correspondant à la ligne sélectionnée
  var sheetName = getValueForRange(CST().DILUANT_FLASK_TO_ARCHIVE);
  var lineToArchive = getValueForRange(CST().DILUANT_LINE_NO_TO_ARCHIVE);

  archiveFromTable(CST().DILUANT_FORMULA_FROM_ZONE, lineToArchive, sheetName);

}



// ==== CUSTOM FUNCTION ====== utilisable dans les formules de calcul

function diluantSheetName(orderNo, flaconNo) {
  return flaconNo+"_"+orderNo;
}

function getFormula(l,c){
  r=activeSheet().getRange(l,c);
 rep= r.getFormula(); 
  if( rep==""){ return "VIDE"}
  return rep;
}

// ==== temporary function ====
function convertHyp(){
  for(i=1;i<=20;i++){
    var val=activeSheet().getRange(i,3).getValue();
    var form = activeSheet().getRange(i,5).getValue();
    var gid=form.split('"')[1];
    var out='=HYPERLINK("'+gid+'";"'+val+'")';
    activeSheet().getRange(i,10).setFormula(out);
  }
}