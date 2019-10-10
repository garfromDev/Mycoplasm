// Ces paramètres permettent d'adapter les scripts à des modifications de lignes et colonne des
// tableaux
function CST(){
  return {
// ============= Global ==============    
// the Name of the folder where archiving must take place
    AG_ARCHIVE_FOLDER_ID : "Archive_Mycoplasmes_Biovac",

// ============= Sheet "Etat des lieux" ==============
// name of the sheet with the état des lieux follow up
    ETAT_DES_LIEUX_NAME : "Etat des lieux",

// ============= Sheet "Stock prepa antigène" ==============    
// name of the sheet with the prepa Antigen stock follow up
    STOCK_AG_NAME : "Stock prepa antigène",
// the index of the first line of the stock table
    STOCK_AG_FIRST_LINE : 8,
// the portion of the stock table with formula that needs to be copied on the new inserted line
    STOCK_AG_FORMULA_FROM_ZONE : "A8:N8",
// the position where the formula must be pasted after the insertion of the new line
    STOCK_AG_FORMULA_TO_ZONE : "A7:N7",
// the cell where lot number must be written
    STOCK_AG_LOT_NO : "C7",
// the cell where type must be written
    STOCK_AG_TYPE : "B7",
// the cell where the lot to archive is choosen
    STOCK_AG_LOT_NO_TO_ARCHIVE : "K3",
// the cell whith the line number corresponding to the lot to be archived    
    STOCK_AG_LINE_NO_OF_LOT_TO_ARCHIVE : "N3",
    
// ============= Sheet "prepa AG template" ==============   
// name of the sheet with template for prepa AG lot
    PREPA_AG_TEMPLATE : "Prepa AG Template",    
// the cell where type is written in prepa AG template
    AG_TEMPLATE_TYPE : "A2",
// the cell where lot no is written in prepa AG template
    AG_TEMPLATE_LOT_NO : "B2",
// the first cell where volume usage is written in prepa AG template
// this is the top left cell of date column, below the title 
    PREPA_AG_TEMPLATE_VOLUME_USAGE_LOCATION : "S6",
// cell where the date of stock update is written
    PREPA_AG_TEMPLATE_STOCK_UPDATE_LOCATION : "M14",
    
// ============= Sheet "Stock diluants" ==============
// name of the sheet with the Diluant stock folluw up
    STOCK_DILUANT_NAME : "Stock diluants",
// cell to input the diluant order no in "Stock diluant" sheet
    INPUT_DILUANT_ORDER_NO : "D4",
// cell to input the diluant volume for the flacon in "Stock diluant" sheet
    INPUT_DILUANT_VOLUME : "E4",
// cell where the date for previous reception is in table "Stock diluant"
    DILUANT_PREVIOUS_DATE : "C7",
// cell where the type is query result in stock diluant
    INPUT_DILUANT_TYPE : "G4",
// cell where user input reception date for the flask in sheet Stock diluant
    DILUANT_FLACON_RECEPTION_DATE : "C4",
// cell where the order date is query result in stock diluant
    INPUT_DILUANT_ORDER_DATE : "H4",
// cell where flacon No associated to curent order is in stock diluant
    DILUANT_FLACON_NO : "H3",
// cell where user input reception date for flacon in "Stock diluant" sheet
    DILUANT_FLACON_RECEPTION_DATE : "C4",
// cell where order no is in stock diluant table_DILUANT
    DILUANT_ORDER_NO : "B7",
// cell where flask no is in stock diluant table_DILUANT
    DILUANT_TABLE_FLACON_NO : "E7",
// the cell where diluant usage must be written in diluant sheet
    DILUANT_VOLUME_USAGE_LOCATION : "G6",    
// the cell where SGP usage must be written in SGP sheet
    SGP_VOLUME_USAGE_LOCATION : "R6",
// the portion of the stock table with formula that needs to be copied on the new inserted line
    DILUANT_FORMULA_FROM_ZONE : "A8:L8",
// the position where the formula must be pasted after the insertion of the new line
    DILUANT_FORMULA_TO_ZONE : "A7:L7",
// the cell where a formula validate the input for diluant flacon reception    
    DILUANT_VALID_INPUT : "G2",
// the cell that gives the flask selected for archiving
    DILUANT_FLASK_TO_ARCHIVE : "L3",
// the cell that give the line number corresponding to the selected flask to archive
    DILUANT_LINE_NO_TO_ARCHIVE : "N3",    
    
// ============= Sheet "Commande diluant" ==============    
// name of the sheet with the diluant order follow up    
    COMMANDE_DILUANT_NAME : "Commande diluant",
// cell which formula is true if valid data for order of diluant
    VALID_DILUANT_ORDER : "F4",    
// cell where user input order date in commande diluant
    USER_INPUT_DILUANT_ORDER_DATE : "B4",
    USER_INPUT_DILUANT_TYPE : "D4",
    USER_INPUT_DILUANT_VOLUME : "E4",
// the portion of the order list table with formula that needs to be copied on the new inserted line
    DILUANT_ORDER_FORMULA_FROM_ZONE : "A8:F8",
// the position where the formula must be pasted after the insertion of the new line
    DILUANT_ORDER_FORMULA_TO_ZONE : "A7:F7",
// the cell where diluant type must be written in order list
    DILUANT_ORDER_TYPE_LOCATION : "D7", 
// the cell where diluant order date must be written in order list
    DILUANT_ORDER_DATE_LOCATION : "C7", 
// the cell where diluant volume must be written in order list
    DILUANT_ORDER_VOLUME_LOCATION : "E7",
// the cell where reception status must be written in order list
    DILUANT_ORDER_RECEPTION_STATUS : "F7",    
// the cell where diluant usage must be written in order list
    DILUANT_ORDER_NO_LOCATION : "B7",
    
// ============= Sheet "Diluant_template" =============
// name of the sheet with template for diluant order
    DILUANT_TEMPLATE : "Diluant_template",
// the cell where diluant type must be written in template
    DILUANT_TYPE_LOCATION_IN_TEMPLATE : "B2",
// the cell where order date type must be written in template
    DILUANT_ORDER_DATE_LOCATION_IN_TEMPLATE : "A2",
// the cell where reception date type must be written in template
    DILUANT_RECEPTION_DATE_LOCATION_IN_TEMPLATE : "D2",
// the cell where volume must be written in template
    DILUANT_VOLUME_LOCATION_IN_TEMPLATE : "G2",
// the cell where flacon no must be written in template
    DILUANT_FLACON_NO_LOCATION_IN_TEMPLATE : "C2",
// the first cell where volume usage is written in diluant template
    DILUANT_TEMPLATE_VOLUME_USAGE_LOCATION : "G6",    

// ============= Sheet "Stock antigènes colorés" =============
// name of the sheet with the Antigene colorés stock follow up
    STOCK_ANTIGENE_COLORE_NAME : "Stock antigènes colorés",
// the portion of the stock table with formula that needs to be copied on the new inserted line
    ANTIGENE_COLORES_FORMULA_FROM_ZONE : "A7:BB7",
// the position where the formula must be pasted after the insertion of the new line
    ANTIGENE_COLORES_FORMULA_TO_ZONE : "A6:BB6",
// the cell where antigene colorés type must be written in stock
    ANTIGENE_COLORES_TYPE_LOCATION : "C6", 
// the column where antigene colorés sheetName is written
    ANTIGENE_COLORES_SHEETNAME_COLUMN : "A", 
// the cell where antigène colorés lot nr must be written in stock
    ANTIGENE_COLORES_LOT_LOCATION : "D6",
// the cell where the lot of antigene colores to archive is
    ANTIGENE_COLORES_LOT_TO_ARCHIVE : "L2",
//the cell where the line to archive is 
    ANTIGENE_COLORES_LINE_TO_ARCHIVE : "M1",

// ============= Sheet "Simulation production AG" =============    
// name of the sheet with the Simulation production AG follow up
    SIMULATION_PRODUCTION_AG_NAME : "Simulation production AG",    
// name of the range with the type    
    INPUT_ANTIGENE_TYPE : "inputTypeAntigeneColores",
// name of the range with the selection of lot to use   
    INPUT_LOT_SELECTION : "inputLotSelection",
// name of the range with the selection of lot to use   
    INPUT_LOT_SELECTION_SGP : "inputLotSelectionSGP",
// the cell that is true when inputs are valid
    ANTIGENE_COLORES_VALID_INPUTS : "AntigeneColoresValidInput",
// the cell that is true when inputs are valid for SGP Standard
    ANTIGENE_COLORES_VALID_INPUTS_SGP : "AntigeneColoresValidInputSGP",
// the cell where the user inputs the lot of antigene colore to launch
    ANTIGENE_COLORES_INPUT_LOT_TO_CREATE : "G12",
// the cell where antigène colorés volume of flask to prepare is in stock
    ANTIGENE_COLORES_VOL1_LOCATION : "F5",
    ANTIGENE_COLORES_VOL2_LOCATION : "F6",
    ANTIGENE_COLORES_VOL3_LOCATION : "F7",
// the cell where antigène colorés nb of flask to prepare is in stock
    ANTIGENE_COLORES_FLACON1_LOCATION : "G5",
    ANTIGENE_COLORES_FLACON2_LOCATION : "G6",
    ANTIGENE_COLORES_FLACON3_LOCATION : "G7",
    
// ============= Sheet "AG Template xx" =============    
// the cell where antigène colorés lot number is in template
    ANTIGENE_COLORES_LOT_LOCATION_IN_TEMPLATE : "B4",
// the cell where antigène colorés type is in template
    ANTIGENE_COLORES_TYPE_LOCATION_IN_TEMPLATE : "B3",
// the cell where antigène colorés volume to prepare is in template
    ANTIGENE_COLORES_VOL1_LOCATION_IN_TEMPLATE : "R11",
    ANTIGENE_COLORES_VOL2_LOCATION_IN_TEMPLATE : "R12",
    ANTIGENE_COLORES_VOL3_LOCATION_IN_TEMPLATE : "R13",
// the cell where antigène colorés nb of flacons to prepare is in template
    ANTIGENE_COLORES_FLACON1_LOCATION_IN_TEMPLATE : "T11",
    ANTIGENE_COLORES_FLACON2_LOCATION_IN_TEMPLATE : "T12",
    ANTIGENE_COLORES_FLACON3_LOCATION_IN_TEMPLATE : "T13",
// the cell where prépa AG lot / SGP Variant used is in antigène colorés template
    ANTIGENE_COLORES_LOT1_PREPA_AG_LOCATION_IN_TEMPLATE : "T21",
    ANTIGENE_COLORES_LOT2_PREPA_AG_LOCATION_IN_TEMPLATE : "T22",
    ANTIGENE_COLORES_LOT3_PREPA_AG_LOCATION_IN_TEMPLATE : "U21",
    ANTIGENE_COLORES_LOT4_PREPA_AG_LOCATION_IN_TEMPLATE : "U22", 
// the cell where prépa AG lot / SGP Standard used is in antigène colorés template
    ANTIGENE_COLORES_LOT1_SGP_STANDARD_LOCATION_IN_TEMPLATE : "T17",
    ANTIGENE_COLORES_LOT2_SGP_STANDARD_LOCATION_IN_TEMPLATE : "T18",
    ANTIGENE_COLORES_LOT3_SGP_STANDARD_LOCATION_IN_TEMPLATE : "U17",
    ANTIGENE_COLORES_LOT4_SGP_STANDARD_LOCATION_IN_TEMPLATE : "U18",
// the cell where prépa AG / SGP Standard  volume used is in antigène colorés template
    ANTIGENE_COLORES_VOLUME1_USED_LOCATION_IN_TEMPLATE : "V21",
    ANTIGENE_COLORES_VOLUME2_USED_LOCATION_IN_TEMPLATE : "V22",
    ANTIGENE_COLORES_VOLUME3_USED_LOCATION_IN_TEMPLATE : "W21",
    ANTIGENE_COLORES_VOLUME4_USED_LOCATION_IN_TEMPLATE : "W22",
// the cell where diluant /SGP Variant lot  used is in antigène colorés template
    ANTIGENE_COLORES_USED_FLACON1_LOCATION_IN_TEMPLATE : "T17",
    ANTIGENE_COLORES_USED_FLACON2_LOCATION_IN_TEMPLATE : "T18",
    ANTIGENE_COLORES_USED_FLACON3_LOCATION_IN_TEMPLATE : "T19",
    ANTIGENE_COLORES_USED_FLACON4_LOCATION_IN_TEMPLATE : "U17",
    ANTIGENE_COLORES_USED_FLACON5_LOCATION_IN_TEMPLATE : "U18",
    ANTIGENE_COLORES_USED_FLACON6_LOCATION_IN_TEMPLATE : "U19",
// the cell where diluant / SGP STANDARD volume used is in antigène colorés template
    ANTIGENE_COLORES_DILUANT_VOLUME_USED1_LOCATION_IN_TEMPLATE : "V17",
    ANTIGENE_COLORES_DILUANT_VOLUME_USED2_LOCATION_IN_TEMPLATE : "V18",
    ANTIGENE_COLORES_DILUANT_VOLUME_USED3_LOCATION_IN_TEMPLATE : "V19",
    ANTIGENE_COLORES_DILUANT_VOLUME_USED4_LOCATION_IN_TEMPLATE : "W17",
    ANTIGENE_COLORES_DILUANT_VOLUME_USED5_LOCATION_IN_TEMPLATE : "W18", 
    ANTIGENE_COLORES_DILUANT_VOLUME_USED6_LOCATION_IN_TEMPLATE : "W19", 
  };
}