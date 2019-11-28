function ziploc(zip1, zip2) {
 
  // zip codes 84009 and 84129 are not recognized by api.zippopotam.us
  if(zip1 === 84009){
    zip1 = 84095 
  } else if (zip1 === 84129){
    zip1 = 84123
  } 
  
  
  if (zip2 === 84009){
    zip2 = 84095
  } else if (zip2 === 84129){
    zip2 = 84123
  }
  
  
  // 1st Zip Code API call
  
  // to learn more about the zippopotam.us API, visit
  var response = UrlFetchApp.fetch("http://api.zippopotam.us/US/" + zip1, {muteHttpExceptions: true});
  
  if (String(response.getResponseCode())[0] === '4'){
    return "Zip code not found"
  }
  
  var a = JSON.parse(response.getContentText());
  
  var lat1 = a.places[0].latitude 
  var lon1 = a.places[0].longitude
  
  
  // 2nd Zip Code API call
  
    // to learn more about the zippopotam.us API, visit
  var response = UrlFetchApp.fetch("http://api.zippopotam.us/US/" + zip2, {muteHttpExceptions: true});
  
  if (String(response.getResponseCode())[0] === '4'){
    return "Zip code not found"
  }
  
  var b = JSON.parse(response.getContentText());
  var lat2 = b.places[0].latitude 
  var lon2 = b.places[0].longitude
  
  return calcCrow(lat1, lon1, lat2, lon2)
}

//This function takes in latitude and longitude of two location and returns the distance between them as the crow flies (in km)
function calcCrow(lat1, lon1, lat2, lon2) 
{
  var R = 6371; // km
  var dLat = toRad(lat2-lat1);
  var dLon = toRad(lon2-lon1);
  var lat1 = toRad(lat1);
  var lat2 = toRad(lat2);

  var a = Math.sin(dLat/2) * Math.sin(dLat/2) +
    Math.sin(dLon/2) * Math.sin(dLon/2) * Math.cos(lat1) * Math.cos(lat2); 
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)); 
  var d = R * c;  
  return d;
}

// Converts numeric degrees to radians
function toRad(Value) 
{
  return (Value * Math.PI / 180) * 0.621371;
}

function onEdit(e){
  var range = e.range;
  var columnOfCellEdited = range.getColumn();
  var rowOfCellEdited = range.getRow();
  var spreadsheet = SpreadsheetApp.getActive();
  
  // only run with zip code or filters is changed
  if (columnOfCellEdited === 5 && (rowOfCellEdited === 3 || rowOfCellEdited === 4  || rowOfCellEdited === 6)) { 
    if (rowOfCellEdited === 3 && spreadsheet.getRange('C4').isChecked()){
      redoFilter()
      copyNames()
      //      lookupZip()
    } else if (rowOfCellEdited === 4 && spreadsheet.getRange('C5').isChecked()){
      redoFilter()
      copyNames()
      //      lookupCity()
    } else if (rowOfCellEdited === 6){
      redoFilter()
      copyNames()
    }
  } else if (columnOfCellEdited === 3 && (rowOfCellEdited === 4 || rowOfCellEdited === 5)){
    toggleCheckboxes(rowOfCellEdited)
  }  
}

function redoFilter() {
  clearFilters()
  
  // Recreate filter
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D8').activate();
  spreadsheet.getRange('D8:G15').createFilter();
  
  // Get Radius in C2
  spreadsheet.getRange('E6').activate();
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  var val = cell.getValue();
  
  var criteria = SpreadsheetApp.newFilterCriteria().whenNumberLessThanOrEqualTo(val).build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(6, criteria);
  
  // Sort Last Lead Received from oldest to youngest
  spreadsheet.getActiveSheet().getFilter().sort(7, true);
};

function clearFilters() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  spreadsheet.getActiveSheet().getFilter().remove();
  spreadsheet.getRange('E5').activate();
  
  // clear column h cells
  spreadsheet.getRange('K1:K25').clear({contentsOnly: true, skipFilteredRows: false});
  
  // clear dropdown
  spreadsheet.getRange('G5').clear({contentsOnly: true})
};

function copyNames(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D8').activate();
  var currentCell = spreadsheet.getCurrentCell()
  spreadsheet.getRange('D8:D24').copyTo(spreadsheet.getRange('K1'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('K2').copyTo(spreadsheet.getRange('G5'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  return spreadsheet.getRange('G5').activate();
}

function updateAgentTimeStamp(){
  var spreadsheet = SpreadsheetApp.getActive();
  if(spreadsheet.getRange('G5').isBlank()){
    return 
  } else {
    spreadsheet.getRange('G5').activate()
    
    findAgent()
    
    spreadsheet.getActiveRange().offset(0,3).activate() 
    spreadsheet.getActiveCell().setValue(new Date());
    
    // Sort Last Lead Received from oldest to youngest
    spreadsheet.getActiveSheet().getFilter().sort(5, true);
    
    copyNames()
  }
  
  spreadsheet.getRange('G5').activate()
}

function findAgent() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D9').activate()
  while (spreadsheet.getActiveRange().getValue() !== spreadsheet.getRange('G5').getValue()){
    spreadsheet.getActiveRange().offset(1,0).activate()
  }
}

function toggleCheckboxes(rowOfCellEdited){
  var spreadsheet = SpreadsheetApp.getActive();
  if (rowOfCellEdited === 4){
    if (spreadsheet.getRange('C4').isChecked()){
      lightenZip()
      spreadsheet.getRange('C5').setValue(false)
    } else {
      lightenCity()
      spreadsheet.getRange('C5').setValue(true)
    }
  } else {
    if (spreadsheet.getRange('C5').isChecked()){
      lightenCity()
      spreadsheet.getRange('C4').setValue(false)
    } else {
      lightenZip()
      spreadsheet.getRange('C4').setValue(true)
    }
  }
} 

//function lookupCity() {
//  var spreadsheet = SpreadsheetApp.getActive();
//  spreadsheet.getRange('E4').activate();
//  spreadsheet.getCurrentCell().setFormula('=VLOOKUP(E5,\'Utah Zip Codes\'!B2:D345,3,false)');
//}
//
//function lookupZip(){
//  var spreadsheet = SpreadsheetApp.getActive();
//  spreadsheet.getRange('E5').activate();
//  spreadsheet.getCurrentCell().setFormula('=VLOOKUP(E4,\'Utah Zip Codes\'!A2:B345,2,false)'); 
//}

function lightenCity() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D4:E4').setFontColor('#cccccc')
  spreadsheet.getRange('D5:E5').setFontColor('#000000')
  spreadsheet.getRange('E5').activate()
}

function lightenZip() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D5:E5').setFontColor('#cccccc')
  spreadsheet.getRange('D4:E4').setFontColor('#000000')
  spreadsheet.getRange('E4').activate()
}