// this function must be installed as trigger by the owner  from Edition/Declencheur du projet actuel/En cours de modification
// do not use <après modification>, which is executed after inserting line, ...
// a simple onEdit() would be executed with current user permission, so cannot change protection
// trigger installed from owner are executed with owner's protection
//
// the execution is triggered by mean of a check box that change the value of underlying cell
// the script to execute is written in the note of the cell as "@@@functionName" without the parenthesis
function InstalledOnEdit(e) {
  // force refresh before launching script to avoid cached value
  SpreadsheetApp.flush();
   //refreshSheet(e.source, e.range.getSheet()); // !!! commenter cette ligne si problèmes !!!
   var cmd = getValidCommand(e.range);
   if(cmd==null){return}
   // in every case, protection of current sheet is deprotected and reprotected after execution of script 
  var originalUnprotected = unprotectSheet(activeSheet());
 
  try{
   cmd();  
  }catch(e){
    alert("erreur : "+e.message);
    return;
  }finally{
     e.range.setValue(false); //check box back to empty state
     restoreProtection(originalUnprotected);
     SpreadsheetApp.flush(); //make sure every changes applied to the sheet before user start next operation
  }
}
  


/**
* return the command (i.e. a function) to be executed
* from the given range
* conditions are :
* range is one cell, with true value, with the name of the function to execute
* stored in the note prefixed with "@@@"
* @param {Range} range
* @return {Function} null if condition are not meet, otherwise the function to execute
*/
function getValidCommand(range){
  if(range.getNumRows()>1){return null}
  if(range.getNumColumns()>1){return null}
  if(range.getValue()!=true){return null}
  var cmd = getCommandFromNote(range.getNote());
  if(cmd==""){return null}
  return new Function(cmd); //create a function with the command as body
}


/**
* @param {String} note
* @return {String} the code of the command (the function call f() )
*/
function getCommandFromNote(note){
  if(note.length<4) {return ""} //il faut @@@ + nom de fonction => 4 mini
  if(note.substring(0,3)!="@@@") {return ""}
  return note.substring(3)+"()";
}



// =========== DISCUSSION =============
// another method would be to deploy this script as webApp
// and call it by script running as user
// see https://stackoverflow.com/questions/42809202/google-apps-script-temporarily-run-script-as-another-user/42820377?noredirect=1#comment72767832_42820377


// ============ PROV ===================
function testFromString(){
  var str = "change()";
  var tmpfunc = new Function(str);
  tmpfunc();
}


function change(){
  if(getSheet("toto")!=null){
     SpreadsheetApp.getActiveSpreadsheet().deleteSheet(getSheet("toto"));
  }
  var a = activeSheet();
  activeSheet().getRange("C1").setValue("démarage");
  activeSheet().getRange("C2").setValue(false);
  return;
  
  
  var tmpl = getSheet("Diluant_Template");
  var newSheet = copyTemplateTo(tmpl, "toto");
  copyProtectionFromSheetToSheet(tmpl, newSheet);
  a.getRange("C3").setValue("modifié à " +new Date());
  retun;
  
  
  
  var protection = tmpl.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
    a.getRange("C1").setValue("protection template de type "+protection.getProtectionType());
  //Utilities.sleep(400);
  var newProtection =  newSheet.protect();
  a.getRange("C1").setValue("protection crée de type "+newProtection.getProtectionType());
  Utilities.sleep(400);
  var ur = protection.getUnprotectedRanges();
  alert("range non protégés "+ur);
  alert("can edit new protection  : "+newProtection.canEdit());
  var targetUr=[];
  for(i=0;i<ur.length;i++){
    var r = ur[i].getA1Notation();
    targetUr.push(newSheet.getRange(r));
  }
  alert("target range : "+targetUr);
  newProtection.setUnprotectedRanges(targetUr);
  a.getRange("C1").setValue("unprotected ajoutés");
  a.getRange("C1").setValue("unprotected ajoutés nb : "+newProtection.getUnprotectedRanges().length);
  
  Utilities.sleep(800);   
  alert("editor template : "+protection.getEditors());
  newProtection.removeEditors(newProtection.getEditors());  
  newProtection.addEditors(protection.getEditors())
  alert("new file : "+newProtection.getEditors());
  a.getRange("C1").setValue("editor ajouté  "+newProtection.getEditors());
  Utilities.sleep(400);    
    newProtection.setDescription(newSheet.getSheetName())
  a.getRange("C1").setValue("description ajouté");   
  a.getRange("C3").setValue("modifié à " +new Date());
}


