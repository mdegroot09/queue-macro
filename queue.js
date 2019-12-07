function zipIt(zip1, zip2) {
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
  
  // Get lat and lon of zip from zippopotam.us API
  var response = UrlFetchApp.fetch("http://api.zippopotam.us/US/" + zip1, {muteHttpExceptions: true});
  
  if (String(response.getResponseCode())[0] === '4'){
    return "Zip code not found"
  }
  
  var a = JSON.parse(response.getContentText());
  
  var lat1 = a.places[0].latitude 
  var lon1 = a.places[0].longitude
  
  
  // 2nd Zip Code API call
  
    // Get lat and lon of zip from zippopotam.us API
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
  return d
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
  var zip = spreadsheet.getRange('E5').getValue()
  var ui = SpreadsheetApp.getUi()
  
  // only run with zip code or filters is changed
  if (columnOfCellEdited === 5 && (rowOfCellEdited === 4 || rowOfCellEdited === 5 || rowOfCellEdited === 6)) { 
    if (rowOfCellEdited === 4){
      
      // change zip code to match entered city
      zip = lookupZip()
      spreadsheet.getRange('E5').setValue(zip)
      
      // if City name is changed AND selected
      redoFormulas(zip)
      redoFilter()
      copyNames()
      
    } else if (rowOfCellEdited === 5){
      
      // change city name to match entered zip code
      var city = lookupCity()
      spreadsheet.getRange('E4').setValue(city)
      
      // if Zip Code is changed AND selected
      redoFormulas(zip)
      redoFilter()
      copyNames()
            
    } else if (rowOfCellEdited === 6){
      
      // if Mile Radius is changed
      redoFormulas(zip)
      redoFilter()
      copyNames()
      
    }
  } else if (columnOfCellEdited === 8 && rowOfCellEdited === 9){
    
    // Check if buyer name is filled out
    if (spreadsheet.getRange('H9').getValue() === 'Assign' && (!spreadsheet.getRange('I5').getValue() || (!spreadsheet.getRange('I6').getValue() && !spreadsheet.getRange('I7').getValue()))){
      spreadsheet.getRange('H9').setValue('')
      if (!spreadsheet.getRange('I5').getValue()) {
        errorBox('I5')
      }
      if (!spreadsheet.getRange('I6').getValue() && !spreadsheet.getRange('I7').getValue()) {
        errorBox('I6')
        errorBox('I7')
      }
//      ui.alert('Please fill out the buyer info.') 
    } else {
      // When H9 is changed to 'ASSIGN', change
      if(spreadsheet.getRange('H9').getValue() === 'Assign' && !spreadsheet.getRange('E9').isBlank()){
        updateAgentTimeStamp()
        spreadsheet.getRange('H9').setValue('')
      }
    } 
  }
}

function redoFormulas(zip){
  var spreadsheet = SpreadsheetApp.getActive();
  
  // clear dropdown
  spreadsheet.getRange('E9').clear({contentsOnly: true})
  
  // add two calls for each agent
  spreadsheet.getRange('Y14').setValue("=zipIt(VLOOKUP(D14,'Agent Team List'!$A$2:$C$8,3,false),"+zip+")")
  spreadsheet.getRange('Y15').setValue("=zipIt(VLOOKUP(D15,'Agent Team List'!$A$2:$C$8,3,false),"+zip+")")
  spreadsheet.getRange('Y16').setValue("=zipIt(VLOOKUP(D16,'Agent Team List'!$A$2:$C$8,3,false),"+zip+")")
  spreadsheet.getRange('Y17').setValue("=zipIt(VLOOKUP(D17,'Agent Team List'!$A$2:$C$8,3,false),"+zip+")")
  spreadsheet.getRange('Y18').setValue("=zipIt(VLOOKUP(D18,'Agent Team List'!$A$2:$C$8,3,false),"+zip+")")
  spreadsheet.getRange('Y19').setValue("=zipIt(VLOOKUP(D19,'Agent Team List'!$A$2:$C$8,3,false),"+zip+")")
  spreadsheet.getRange('Y20').setValue("=zipIt(VLOOKUP(D20,'Agent Team List'!$A$2:$C$8,3,false),"+zip+")")

  spreadsheet.getRange('Z14').setValue("=zipIt("+zip+",VLOOKUP(D14,'Agent Team List'!$A$2:$C$8,3,false))")
  spreadsheet.getRange('Z15').setValue("=zipIt("+zip+",VLOOKUP(D15,'Agent Team List'!$A$2:$C$8,3,false))")
  spreadsheet.getRange('Z16').setValue("=zipIt("+zip+",VLOOKUP(D16,'Agent Team List'!$A$2:$C$8,3,false))")
  spreadsheet.getRange('Z17').setValue("=zipIt("+zip+",VLOOKUP(D17,'Agent Team List'!$A$2:$C$8,3,false))")
  spreadsheet.getRange('Z18').setValue("=zipIt("+zip+",VLOOKUP(D18,'Agent Team List'!$A$2:$C$8,3,false))")
  spreadsheet.getRange('Z19').setValue("=zipIt("+zip+",VLOOKUP(D19,'Agent Team List'!$A$2:$C$8,3,false))")
  spreadsheet.getRange('Z20').setValue("=zipIt("+zip+",VLOOKUP(D20,'Agent Team List'!$A$2:$C$8,3,false))")
  
  // select distance result that isn't an error
  if (spreadsheet.getRange('Y14').getValue() !== "#ERROR!"){
    spreadsheet.getRange('G14').setValue(spreadsheet.getRange('Y14').getValue())
  } else {
    spreadsheet.getRange('G14').setValue(spreadsheet.getRange('Z14').getValue())
  }
  
  if (spreadsheet.getRange('Y15').getValue() !== "#ERROR!"){
    spreadsheet.getRange('G15').setValue(spreadsheet.getRange('Y15').getValue())
  } else {
    spreadsheet.getRange('G15').setValue(spreadsheet.getRange('Z15').getValue())
  }
  
  if (spreadsheet.getRange('Y16').getValue() !== "#ERROR!"){
    spreadsheet.getRange('G16').setValue(spreadsheet.getRange('Y16').getValue())
  } else {
    spreadsheet.getRange('G16').setValue(spreadsheet.getRange('Z16').getValue())
  }
  
  if (spreadsheet.getRange('Y17').getValue() !== "#ERROR!"){
    spreadsheet.getRange('G17').setValue(spreadsheet.getRange('Y17').getValue())
  } else {
    spreadsheet.getRange('G17').setValue(spreadsheet.getRange('Z17').getValue())
  }
  
  if (spreadsheet.getRange('Y18').getValue() !== "#ERROR!"){
    spreadsheet.getRange('G18').setValue(spreadsheet.getRange('Y18').getValue())
  } else {
    spreadsheet.getRange('G18').setValue(spreadsheet.getRange('Z18').getValue())
  }
  
  if (spreadsheet.getRange('Y19').getValue() !== "#ERROR!"){
    spreadsheet.getRange('G19').setValue(spreadsheet.getRange('Y19').getValue())
  } else {
    spreadsheet.getRange('G19').setValue(spreadsheet.getRange('Z19').getValue())
  }
  
  if (spreadsheet.getRange('Y20').getValue() !== "#ERROR!"){
    spreadsheet.getRange('G20').setValue(spreadsheet.getRange('Y20').getValue())
  } else {
    spreadsheet.getRange('G20').setValue(spreadsheet.getRange('Z20').getValue())
  }
}

function redoFilter() {
  clearFilters()
  
  // Recreate filter
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D13:I20').createFilter();
  
  // Get Radius in E6
  var val = spreadsheet.getRange('E6').getValue()
  
  var criteria = SpreadsheetApp.newFilterCriteria().whenNumberLessThanOrEqualTo(val).build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(7, criteria);
  
  // Sort Last Lead Received from oldest to youngest
  spreadsheet.getActiveSheet().getFilter().sort(8, true);
  
  // Sort 7-Day Total from least to most
  spreadsheet.getActiveSheet().getFilter().sort(9, true);
};

function clearFilters() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().getFilter().remove();
  
  // clear column h cells
  spreadsheet.getRange('L1:L25').clear({contentsOnly: true, skipFilteredRows: false});
  
  // clear dropdown
  spreadsheet.getRange('E9').clear({contentsOnly: true})
};

function copyNames(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D13:D33').copyTo(spreadsheet.getRange('L1'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('L2').copyTo(spreadsheet.getRange('E9'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  return;
}

function updateAgentTimeStamp(){
  var spreadsheet = SpreadsheetApp.getActive();
  if(spreadsheet.getRange('E9').isBlank()){
    return 
  } else {
    
    var spreadsheet = SpreadsheetApp.getActive();
    var agentName = spreadsheet.getRange('E9').getValue()
    var buyerName = spreadsheet.getRange('I5').getValue()
    var buyerPhone = spreadsheet.getRange('I6').getValue()
    var buyerEmail = spreadsheet.getRange('I7').getValue()
    var zip = spreadsheet.getRange('E5').getValue()
    
    spreadsheet.getSheetByName(agentName).insertRowsBefore(9,1)
    spreadsheet.getSheetByName(agentName).getRange('A9').setValue(new Date());
    spreadsheet.getSheetByName(agentName).getRange('A9').setNumberFormat('m"/"d" "h":"mma/p');
    spreadsheet.getSheetByName(agentName).getRange('B9').setValue(buyerName)
    spreadsheet.getSheetByName(agentName).getRange('C9').setValue(buyerPhone)
    spreadsheet.getSheetByName(agentName).getRange('D9').setValue(buyerEmail)
    
    // clear Buyer Info inputs and redo formatting
    spreadsheet.getRange('I5:I7').clear({contentsOnly: true})
    spreadsheet.getRange('I5:I7').setBackground('#fff2cc')
    .setBorder(true, true, true, true, true, true, '#ffe599', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
    spreadsheet.getRange('H5:I7').setBorder(true, true, true, true, null, null, '#58dbc2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    // Sort Last Lead Received from oldest to youngest
    spreadsheet.getActiveSheet().getFilter().sort(8, true);
    
    // Sort 7-Day Total from least to most
    spreadsheet.getActiveSheet().getFilter().sort(9, true);

    redoFormulas(zip)
    copyNames()
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

function lookupCity() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E4').setFormula('=VLOOKUP(E5,\'Utah Zip Codes\'!B2:D390,3,false)');
  return spreadsheet.getRange('E4').getValue()
}

function lookupZip(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E5').setFormula('=VLOOKUP(E4,\'Utah Zip Codes\'!A2:B390,2,false)'); 
  return spreadsheet.getRange('E5').getValue()
}

function lightenCity() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D4:E4').setFontColor('#cccccc')
  spreadsheet.getRange('E4').setBackground('#fefefe')
  spreadsheet.getRange('E4').setBorder(true, true, true, true, null, null, '#f5f5f5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  spreadsheet.getRange('D5:E5').setFontColor('#3e494c')
  spreadsheet.getRange('E5').setBackground('#fff2cc')
  spreadsheet.getRange('E5').setBorder(true, true, true, true, null, null, '#ffe599', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
}

function lightenZip() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D5:E5').setFontColor('#cccccc')
  spreadsheet.getRange('E5').setBackground('#fefefe')
  spreadsheet.getRange('E5').setBorder(true, true, true, true, null, null, '#f5f5f5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  spreadsheet.getRange('D4:E4').setFontColor('#3e494c')
  spreadsheet.getRange('E4').setBackground('#fff2cc')
  spreadsheet.getRange('E4').setBorder(true, true, true, true, null, null, '#ffe599', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
}

function errorBox(cell) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange(cell).setBackground('#f4cccc')
  .setBorder(true, true, true, true, null, null, '#ea9999', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('H5:I7').setBorder(true, true, true, true, null, null, '#58dbc2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('H4:I4').setBorder(true, true, true, true, null, null, '#58dbc2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}