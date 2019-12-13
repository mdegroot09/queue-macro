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
    if (rowOfCellEdited === 4 && spreadsheet.getRange('E4').getValue()){
      
      // if city name is changed
      agentCellTurnGray()
      
      // change zip code to match entered city
      zip = lookupZip()
      spreadsheet.getRange('E5').setValue(zip)
      
      // if City name is changed AND selected
      redoFormulas(zip)
      redoFilter()
      copyNames()
      
      agentCellTurnOrange()
      
    } else if (rowOfCellEdited === 5 && spreadsheet.getRange('E5').getValue()){
      
      // if zip code is changed
      agentCellTurnGray()
      
      // change city name to match entered zip code
      var city = lookupCity()
      spreadsheet.getRange('E4').setValue(city)
      
      // if Zip Code is changed AND selected
      redoFormulas(zip)
      redoFilter()
      copyNames()
      
      agentCellTurnOrange()
            
    } else if (rowOfCellEdited === 6){
      
      // if Mile Radius is changed
      agentCellTurnGray()
      // redoFormulas(zip)
      // redoFilter()
      
      // Get Radius in E6
      var val = spreadsheet.getRange('E6').getValue()
      
      var criteria = SpreadsheetApp.newFilterCriteria().whenNumberLessThanOrEqualTo(val).build();
      spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(7, criteria);
      
      // Sort Last Lead Received from oldest to youngest
      spreadsheet.getActiveSheet().getFilter().sort(8, true);
      
      // Sort 7-Day Total from least to most
      spreadsheet.getActiveSheet().getFilter().sort(9, true)
      
      copyNames()
      
      agentCellTurnOrange()
      
    }
//  } else if (columnOfCellEdited === 6 && rowOfCellEdited === 11){
    
//    // Check if buyer name is filled out
//    if (spreadsheet.getRange('F11').getValue() === 'Assign' && (!spreadsheet.getRange('I5').getValue() || (!spreadsheet.getRange('I6').getValue() && !spreadsheet.getRange('I7').getValue()))){
//      spreadsheet.getRange('F11').setValue('')
//      if (!spreadsheet.getRange('I5').getValue()) {
//        errorBox('I5')
//      }
//      if (!spreadsheet.getRange('I6').getValue() && !spreadsheet.getRange('I7').getValue()) {
//        errorBox('I6')
//        errorBox('I7')
//      }
//    //      ui.alert('Please fill out the buyer info.') 
//    } 
//    else {
//      // When F11 is changed to 'ASSIGN', change
//      if(spreadsheet.getRange('F11').getValue() === 'Assign' && !spreadsheet.getRange('E8').isBlank()){
//        updateAgentTimeStamp()
//        spreadsheet.getRange('F11').setValue('')
//      }
//    } 
  }
}

function redoFormulas(zip){
  var spreadsheet = SpreadsheetApp.getActive();
  
  // clear dropdown
  spreadsheet.getRange('E8').clear({contentsOnly: true})
  agentCellTurnGray()
  
  // add two calls for each agent
  spreadsheet.getRange('Y14').setValue("=zipIt(F14,E5)")
  spreadsheet.getRange('Y15').setValue("=zipIt(F15,E5)")
  spreadsheet.getRange('Y16').setValue("=zipIt(F16,E5)")
  spreadsheet.getRange('Y17').setValue("=zipIt(F17,E5)")
  spreadsheet.getRange('Y18').setValue("=zipIt(F18,E5)")
  spreadsheet.getRange('Y19').setValue("=zipIt(F19,E5)")
  spreadsheet.getRange('Y20').setValue("=zipIt(F20,E5)")
  spreadsheet.getRange('Y21').setValue("=zipIt(F21,E5)")
  spreadsheet.getRange('Y22').setValue("=zipIt(F22,E5)")
  spreadsheet.getRange('Y23').setValue("=zipIt(F23,E5)")
  spreadsheet.getRange('Y24').setValue("=zipIt(F24,E5)")

  spreadsheet.getRange('Z14').setValue("=zipIt(E5,F14)")
  spreadsheet.getRange('Z15').setValue("=zipIt(E5,F15)")
  spreadsheet.getRange('Z16').setValue("=zipIt(E5,F16)")
  spreadsheet.getRange('Z17').setValue("=zipIt(E5,F17)")
  spreadsheet.getRange('Z18').setValue("=zipIt(E5,F18)")
  spreadsheet.getRange('Z19').setValue("=zipIt(E5,F19)")
  spreadsheet.getRange('Z20').setValue("=zipIt(E5,F20)")
  spreadsheet.getRange('Z21').setValue("=zipIt(E5,F21)")
  spreadsheet.getRange('Z22').setValue("=zipIt(E5,F22)")
  spreadsheet.getRange('Z23').setValue("=zipIt(E5,F23)")
  spreadsheet.getRange('Z24').setValue("=zipIt(E5,F24)")
  
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
  
  if (spreadsheet.getRange('Y21').getValue() !== "#ERROR!"){
    spreadsheet.getRange('G21').setValue(spreadsheet.getRange('Y21').getValue())
  } else {
    spreadsheet.getRange('G21').setValue(spreadsheet.getRange('Z21').getValue())
  }
  
  if (spreadsheet.getRange('Y22').getValue() !== "#ERROR!"){
    spreadsheet.getRange('G22').setValue(spreadsheet.getRange('Y22').getValue())
  } else {
    spreadsheet.getRange('G22').setValue(spreadsheet.getRange('Z22').getValue())
  }
  
  if (spreadsheet.getRange('Y23').getValue() !== "#ERROR!"){
    spreadsheet.getRange('G23').setValue(spreadsheet.getRange('Y23').getValue())
  } else {
    spreadsheet.getRange('G23').setValue(spreadsheet.getRange('Z23').getValue())
  }
  
  if (spreadsheet.getRange('Y24').getValue() !== "#ERROR!"){
    spreadsheet.getRange('G24').setValue(spreadsheet.getRange('Y24').getValue())
  } else {
    spreadsheet.getRange('G24').setValue(spreadsheet.getRange('Z24').getValue())
  }
}

function redoFilter() {
  clearFilters()
  
  // Recreate filter
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D13:I24').createFilter();
  
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
  spreadsheet.getRange('E8').clear({contentsOnly: true})
  agentCellTurnGray()
};

function copyNames(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D13:D33').copyTo(spreadsheet.getRange('L1'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('L2').copyTo(spreadsheet.getRange('E8'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  agentCellTurnOrange()
  return;
}

function updateAgentTimeStamp(){
  
  var spreadsheet = SpreadsheetApp.getActive()
  
  // Check if buyer info is filled out
  if (!spreadsheet.getRange('I5').getValue() || (!spreadsheet.getRange('I6').getValue() && !spreadsheet.getRange('I7').getValue())){
    if (!spreadsheet.getRange('I5').getValue()) {
      errorBox('I5')
    }
    if (!spreadsheet.getRange('I6').getValue() && !spreadsheet.getRange('I7').getValue()) {
      errorBox('I6')
      errorBox('I7')
    }
  //      ui.alert('Please fill out the buyer info.') 
  } else if(spreadsheet.getRange('E8').isBlank()){
    
    errorBox('E8:F9') 
  
  } else {
    
    buyerInfoTurnGray()
    agentCellTurnGray()
    
    if (spreadsheet.getRange('I8').getValue()){
      updateMaster()
    }
    
    var spreadsheet = SpreadsheetApp.openById('1Cqy5-CySvFJtWtkkli8UNGnjSSOX9DeZz_5FKpkmXlM')
    var queue = spreadsheet.getSheetByName('Queue')
    
    var buyerAgent = queue.getRange('E8').getValue()
    var buyerName = queue.getRange('I5').getValue()
    var buyerPhone = queue.getRange('I6').getValue()
    var buyerEmail = queue.getRange('I7').getValue()
    var listingAgent = queue.getRange('I8').getValue()
    var source = queue.getRange('I9').getValue()
    var tags = queue.getRange('I10').getValue()
    var notes = queue.getRange('I11').getValue()
    var zip = queue.getRange('E5').getValue()
    
    var buyerAgentSheet = spreadsheet.getSheetByName(buyerAgent)
    
    if (buyerAgentSheet) {    
      spreadsheet.getSheetByName(buyerAgent).insertRowsBefore(9,1)
      spreadsheet.getSheetByName(buyerAgent).getRange('A9').setValue(new Date());
      spreadsheet.getSheetByName(buyerAgent).getRange('A9').setNumberFormat('m"/"d" "h":"mma/p');
      spreadsheet.getSheetByName(buyerAgent).getRange('B9').setValue(buyerName)
      spreadsheet.getSheetByName(buyerAgent).getRange('C9').setValue(buyerPhone)
      spreadsheet.getSheetByName(buyerAgent).getRange('D9').setValue(buyerEmail)
    }
    
    // clear Buyer Info inputs and redo formatting
    agentCellTurnOrange();
    spreadsheet.getRange('I5:I11').clear({contentsOnly: true})
    buyerInfoTurnOrange()
    
    // clear city, zip, and miles distances for each agent
    spreadsheet.getRange('E4:E5').clear({contentsOnly: true})
    spreadsheet.getRange('E6').setValue(15)
    spreadsheet.getRange('G14:G24').clear({contentsOnly: true})
    
    // clear the miles radius filter
    spreadsheet.getActiveSheet().getFilter().removeColumnFilterCriteria(7)
    
    // Sort Last Lead Received from oldest to youngest
    spreadsheet.getActiveSheet().getFilter().sort(8, true);
    
    // Sort 7-Day Total from least to most
    spreadsheet.getActiveSheet().getFilter().sort(9, true);

    // redoFormulas(zip)
    
    copyNames()
    agentCellTurnOrange()
  }
}


function updateMaster(){
  var spreadsheet = SpreadsheetApp.getActive();
  var buyerAgent = spreadsheet.getRange('E8').getValue()
  var buyerName = spreadsheet.getRange('I5').getValue()
  var buyerPhone = spreadsheet.getRange('I6').getValue()
  var buyerEmail = spreadsheet.getRange('I7').getValue()
  var listingAgent = spreadsheet.getRange('I8').getValue()
  var source = spreadsheet.getRange('I9').getValue()
  var tags = spreadsheet.getRange('I10').getValue()
  var notes = spreadsheet.getRange('I11').getValue()
  var zip = spreadsheet.getRange('E5').getValue()
    
  var master = SpreadsheetApp.openById('1jHTJbt4FM4WGbHSy0nGF8OEpArik44Qmj0Ba7GfMOnE')
  var referrals = master.getSheetByName('Referrals')
  
  referrals.insertRowsBefore(referrals.getRange('4:4').getRow(), 1);  
  referrals.getRange('A4').setValue(buyerName)
  referrals.getRange('B4').setValue(listingAgent)
  referrals.getRange('C4').setValue('Lead')
  referrals.getRange('D4').setValue(600)
  referrals.getRange('E4').setValue(source)
  referrals.getRange('G4').setValue(buyerAgent)
  referrals.getRange('H4').setValue('Open')
  referrals.getRange('K4').setFormula('=IF(B4="","",VLOOKUP(B4,Setting!A:B,2,false))')
  referrals.getRange('L4').setValue(tags)
  referrals.getRange('M4').setValue(notes)
  referrals.getRange('N4').setFormula('=IF(F4="","",IFS(F4="TBD","TBD",MONTH(F4)=1,"January",MONTH(F4)=2,"February",MONTH(F4)=3,"March",MONTH(F4)=4,"April",MONTH(F4)=5,"May",MONTH(F4)=6,"June",MONTH(F4)=7,"July",MONTH(F4)=8,"August",MONTH(F4)=9,"September",MONTH(F4)=10,"October",MONTH(F4)=11,"November",MONTH(F4)=12,"December"))');
  referrals.getRange('O4').setFormula('=IF(F4="","",IF(F4="TBD","TBD",year(F4)))');
  referrals.getRange('P4').setFormula('=IFS(N4="TBD","TBD",N4="","",N4>0,O4&" "&N4)');
  referrals.getRange('Q4').setValue('=TODAY()')
  referrals.getRange('Q4').setNumberFormat('m"/"d"/"yy')
  var date = referrals.getRange('Q4').getValue()
  referrals.getRange('Q4').setValue(date)
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
  if (spreadsheet.getRange('E5').getValue()){
    spreadsheet.getRange('E4').setFormula('=VLOOKUP(E5,\'Utah Zip Codes\'!B2:D391,3,false)');
    return spreadsheet.getRange('E4').getValue()
  } else {
    return ''
  }
}

function lookupZip(){
  var spreadsheet = SpreadsheetApp.getActive();
  if (spreadsheet.getRange('E4').getValue()){
    spreadsheet.getRange('E5').setFormula('=VLOOKUP(E4,\'Utah Zip Codes\'!A2:B391,2,false)'); 
    return spreadsheet.getRange('E5').getValue()
  } else {
    return ''
  }
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
  spreadsheet.getRange('H5:I11').setBorder(true, true, true, true, null, null, '#58dbc2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('H4:I4').setBorder(true, true, true, true, null, null, '#58dbc2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function agentCellTurnGray(){
  var spreadsheet = SpreadsheetApp.getActive()
  spreadsheet.getRange('E4:E6').setBackground('#f3f3f3')
  .setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  
  spreadsheet.getRange('E8:F9').setBackground('#f3f3f3')
  .setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
}

function agentCellTurnOrange(){
  var spreadsheet = SpreadsheetApp.getActive()
  
  spreadsheet.getRange('E4:E6').setBackground('#fff2cc')
  .setBorder(true, true, true, true, true, true, '#ffe599', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontSize(13)
  .setFontFamily('Arial')
  
  spreadsheet.getRange('E8:F9').setBackground('#fff2cc')
  .setBorder(true, true, true, true, true, true, '#ffe599', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontSize(17)
  .setFontFamily('Arial')
  
  var value = spreadsheet.getRange('D13').offset(1, 0).getValue()
  spreadsheet.getRange('A1').setValue(value)
  spreadsheet.getRange('A1')
}

function buyerInfoTurnGray(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('I5:I11').setBackground('#f3f3f3')
  .setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('H5:I11').setBorder(true, true, true, true, null, null, '#58dbc2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function buyerInfoTurnOrange(){
  var spreadsheet = SpreadsheetApp.getActive();  
  spreadsheet.getRange('I5:I11').setBackground('#fff2cc')
  .setBorder(true, true, true, true, true, true, '#ffe599', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setHorizontalAlignment('left')
  .setVerticalAlignment('middle')
  .setFontSize(11)
  .setFontFamily('Arial');
  spreadsheet.getRange('H5:I11').setBorder(true, true, true, true, null, null, '#58dbc2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}