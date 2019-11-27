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

  // Set a comment on the edited cell to indicate when it was changed.

  var range = e.range;
  var sheet = range.getSheet();

  redoFilter()
  copyNames()
}

function redoFilter() {
  clearFilters()
  
  // Recreate filter
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A4').activate();
  spreadsheet.getRange('A4:D11').createFilter();
  
  // Get Radius in B2
  spreadsheet.getRange('B2').activate();
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  var val = cell.getValue();
  
  var criteria = SpreadsheetApp.newFilterCriteria()
  .whenNumberLessThanOrEqualTo(val)
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(3, criteria);
};

function clearFilters() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  spreadsheet.getActiveSheet().getFilter().remove();
  spreadsheet.getRange('B1').activate();
  
  // clear column h cells
  spreadsheet.getRange('H1:H20').clear({contentsOnly: true, skipFilteredRows: true});
  
  // clear dropdown
  spreadsheet.getRange('E1').clear({contentsOnly: true, skipFilteredRows: true})
};

function copyNames(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A4').activate();
  var currentCell = spreadsheet.getCurrentCell()
  spreadsheet.getRange('A4:A20').copyTo(spreadsheet.getRange('H1'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  return spreadsheet.getRange('H1').activate();
}

function updateAgentTimeStamp(){
  var spreadsheet = SpreadsheetApp.getActive();
  if(spreadsheet.getRange('E1').isBlank()){
    return 
  } else {
    spreadsheet.getRange('E1').activate()
    
    findAgent()
    
    spreadsheet.getActiveRange().offset(0,3).activate() 
    spreadsheet.getActiveCell().setValue(new Date());
  }
}

function findAgent() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A5').activate()
  while (spreadsheet.getActiveRange().getValue() !== spreadsheet.getRange('E1').getValue()){
    spreadsheet.getActiveRange().offset(1,0).activate()
  }
}


