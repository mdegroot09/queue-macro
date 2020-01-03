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
  var ss = SpreadsheetApp.getActive();
  var zip = ss.getRange('E5').getValue()
  var ui = SpreadsheetApp.getUi()
  
  // only run with zip code or filters is changed
  if (columnOfCellEdited === 5 && (rowOfCellEdited === 4 || rowOfCellEdited === 5 || rowOfCellEdited === 6)) { 
    if (rowOfCellEdited === 4 && ss.getRange('E4').getValue()){
      
      // if city name is changed
      agentCellTurnGray()
      
      // Capitalize city name
      var word = ss.getRange('E4').getValue()
      var wordUppercase = word.charAt(0).toUpperCase() + word.slice(1)
      ss.getRange('E4').setValue(wordUppercase)
      
      // change zip code to match entered city
      zip = lookupZip()
      ss.getRange('E5').setValue(zip)
      
      // if City name is changed AND selected
      redoFormulas(zip)
      redoFilter()
      copyNames()
      
      agentCellTurnOrange()
      
    } else if (rowOfCellEdited === 5 && ss.getRange('E5').getValue()){
      
      // if zip code is changed
      agentCellTurnGray()
      
      // change city name to match entered zip code
      var city = lookupCity()
      ss.getRange('E4').setValue(city)
      
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
      var val = ss.getRange('E6').getValue()
      
      var criteria = SpreadsheetApp.newFilterCriteria().whenNumberLessThanOrEqualTo(val).build();
      ss.getActiveSheet().getFilter().setColumnFilterCriteria(7, criteria);
      
      // Sort Last Lead Received from oldest to youngest
      ss.getActiveSheet().getFilter().sort(8, true);
      
      // Sort 7-Day Total from least to most
      ss.getActiveSheet().getFilter().sort(9, true)
      
      copyNames()
      
      agentCellTurnOrange()
      
    }
  }
}

function redoFormulas(zip){
  var ss = SpreadsheetApp.getActive();
  
  // clear radius filter if active
  if (ss.getActiveSheet().getFilter().getColumnFilterCriteria(7)){
    ss.getActiveSheet().getFilter().removeColumnFilterCriteria(7)
  }
  
  // clear dropdown
  ss.getRange('E8').clear({contentsOnly: true})
  agentCellTurnGray()
  
  // add two calls for each agent
  ss.getRange('Y14').setValue("=zipIt(F14,E5)")
  ss.getRange('Y15').setValue("=zipIt(F15,E5)")
  ss.getRange('Y16').setValue("=zipIt(F16,E5)")
  ss.getRange('Y17').setValue("=zipIt(F17,E5)")
  ss.getRange('Y18').setValue("=zipIt(F18,E5)")
  ss.getRange('Y19').setValue("=zipIt(F19,E5)")
  ss.getRange('Y20').setValue("=zipIt(F20,E5)")
  ss.getRange('Y21').setValue("=zipIt(F21,E5)")
  ss.getRange('Y22').setValue("=zipIt(F22,E5)")
  ss.getRange('Y23').setValue("=zipIt(F23,E5)")
  ss.getRange('Y24').setValue("=zipIt(F24,E5)")
  
  ss.getRange('Z14').setValue("=zipIt(E5,F14)")
  ss.getRange('Z15').setValue("=zipIt(E5,F15)")
  ss.getRange('Z16').setValue("=zipIt(E5,F16)")
  ss.getRange('Z17').setValue("=zipIt(E5,F17)")
  ss.getRange('Z18').setValue("=zipIt(E5,F18)")
  ss.getRange('Z19').setValue("=zipIt(E5,F19)")
  ss.getRange('Z20').setValue("=zipIt(E5,F20)")
  ss.getRange('Z21').setValue("=zipIt(E5,F21)")
  ss.getRange('Z22').setValue("=zipIt(E5,F22)")
  ss.getRange('Z23').setValue("=zipIt(E5,F23)")
  ss.getRange('Z24').setValue("=zipIt(E5,F24)")
  
  // select distance result that isn't an error
  if (ss.getRange('Y14').getValue() !== "#ERROR!"){
    ss.getRange('G14').setValue(ss.getRange('Y14').getValue())
  } else {
    ss.getRange('G14').setValue(ss.getRange('Z14').getValue())
  }
  
  if (ss.getRange('Y15').getValue() !== "#ERROR!"){
    ss.getRange('G15').setValue(ss.getRange('Y15').getValue())
  } else {
    ss.getRange('G15').setValue(ss.getRange('Z15').getValue())
  }
  
  if (ss.getRange('Y16').getValue() !== "#ERROR!"){
    ss.getRange('G16').setValue(ss.getRange('Y16').getValue())
  } else {
    ss.getRange('G16').setValue(ss.getRange('Z16').getValue())
  }
  
  if (ss.getRange('Y17').getValue() !== "#ERROR!"){
    ss.getRange('G17').setValue(ss.getRange('Y17').getValue())
  } else {
    ss.getRange('G17').setValue(ss.getRange('Z17').getValue())
  }
  
  if (ss.getRange('Y18').getValue() !== "#ERROR!"){
    ss.getRange('G18').setValue(ss.getRange('Y18').getValue())
  } else {
    ss.getRange('G18').setValue(ss.getRange('Z18').getValue())
  }
  
  if (ss.getRange('Y19').getValue() !== "#ERROR!"){
    ss.getRange('G19').setValue(ss.getRange('Y19').getValue())
  } else {
    ss.getRange('G19').setValue(ss.getRange('Z19').getValue())
  }
  
  if (ss.getRange('Y20').getValue() !== "#ERROR!"){
    ss.getRange('G20').setValue(ss.getRange('Y20').getValue())
  } else {
    ss.getRange('G20').setValue(ss.getRange('Z20').getValue())
  }
  
  if (ss.getRange('Y21').getValue() !== "#ERROR!"){
    ss.getRange('G21').setValue(ss.getRange('Y21').getValue())
  } else {
    ss.getRange('G21').setValue(ss.getRange('Z21').getValue())
  }
  
  if (ss.getRange('Y22').getValue() !== "#ERROR!"){
    ss.getRange('G22').setValue(ss.getRange('Y22').getValue())
  } else {
    ss.getRange('G22').setValue(ss.getRange('Z22').getValue())
  }
  
  if (ss.getRange('Y23').getValue() !== "#ERROR!"){
    ss.getRange('G23').setValue(ss.getRange('Y23').getValue())
  } else {
    ss.getRange('G23').setValue(ss.getRange('Z23').getValue())
  }
  
  if (ss.getRange('Y24').getValue() !== "#ERROR!"){
    ss.getRange('G24').setValue(ss.getRange('Y24').getValue())
  } else {
    ss.getRange('G24').setValue(ss.getRange('Z24').getValue())
  }
}

function redoFilter() {
  clearFilters()
  
  // Recreate filter
  var ss = SpreadsheetApp.getActive();
  ss.getRange('D13:I24').createFilter();
  
  // Get Radius in E6
  var val = ss.getRange('E6').getValue()
  
  var criteria = SpreadsheetApp.newFilterCriteria().whenNumberLessThanOrEqualTo(val).build();
  ss.getActiveSheet().getFilter().setColumnFilterCriteria(7, criteria);
  
  // Sort Last Lead Received from oldest to youngest
  ss.getActiveSheet().getFilter().sort(8, true);
  
  // Sort 7-Day Total from least to most
  ss.getActiveSheet().getFilter().sort(9, true);
};

function clearFilters() {
  var ss = SpreadsheetApp.getActive();
  ss.getActiveSheet().getFilter().remove();
  
  // clear column h cells
  ss.getRange('L1:L25').clear({contentsOnly: true, skipFilteredRows: false});
  
  // clear dropdown
  ss.getRange('E8').clear({contentsOnly: true})
  agentCellTurnGray()
};

function copyNames(){
  var ss = SpreadsheetApp.getActive();
  ss.getRange('D13:D33').copyTo(ss.getRange('L1'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  ss.getRange('L2').copyTo(ss.getRange('E8'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  agentCellTurnOrange()
  return;
}

function assignAgent(){
  
  var ss = SpreadsheetApp.getActive()
  
  // Check if buyer info is filled out
  if (!ss.getRange('I5').getValue() || (!ss.getRange('I6').getValue() && !ss.getRange('I7').getValue())){
    if (!ss.getRange('I5').getValue()) {
      errorBox('I5')
    }
    if (!ss.getRange('I6').getValue() && !ss.getRange('I7').getValue()) {
      errorBox('I6')
      errorBox('I7')
    }
    //      ui.alert('Please fill out the buyer info.') 
  } else if(ss.getRange('E8').isBlank()){
    
    errorBox('E8:F9') 
    
  } else {
    
    buyerInfoTurnGray()
    agentCellTurnGray()
    
    updateAgentSS()
        
    var ss = SpreadsheetApp.getActive()
    var queue = ss.getSheetByName('Queue')
    
    var buyerAgent = queue.getRange('E8').getValue()
    var buyerName = queue.getRange('I5').getValue()
    var buyerPhone = queue.getRange('I6').getValue()
    var buyerEmail = queue.getRange('I7').getValue()
    var listingAgent = queue.getRange('I8').getValue()
    var source = queue.getRange('I9').getValue()
    var tags = queue.getRange('I10').getValue()
    var notes = queue.getRange('I11').getValue()
    var zip = queue.getRange('E5').getValue()
    
    ss.getSheetByName('Raw Data').insertRowsBefore(ss.getSheetByName('Raw Data').getRange('4:4').getRow(), 1);  
    ss.getSheetByName('Raw Data').getRange('A4').setValue(buyerName)
    //  hotWarmLeads.getRange('B4').setValue()
    //  hotWarmLeads.getRange('C4').setValue()
    ss.getSheetByName('Raw Data').getRange('D4').setValue(buyerPhone)
    ss.getSheetByName('Raw Data').getRange('E4').setValue(buyerEmail)
    ss.getSheetByName('Raw Data').getRange('F4').setValue(listingAgent)
    ss.getSheetByName('Raw Data').getRange('G4').setValue('New Lead')
    //    hotWarmLeads.getRange('H4').setValue(600)
    ss.getSheetByName('Raw Data').getRange('I4').setValue(source)
    //  hotWarmLeads.getRange('J4').setValue()
    ss.getSheetByName('Raw Data').getRange('K4').setValue(buyerAgent)
    ss.getSheetByName('Raw Data').getRange('L4').setValue('Open')
    ss.getSheetByName('Raw Data').getRange('O4').setFormula('=IF(B4="","",VLOOKUP(B4,Setting!A:B,2,false))')
    ss.getSheetByName('Raw Data').getRange('P4').setValue(tags)
    ss.getSheetByName('Raw Data').getRange('Q4').setFormula('=IF(J4="","",IFS(J4="TBD","TBD",MONTH(J4)=1,"January",MONTH(J4)=2,"February",MONTH(J4)=3,"March",MONTH(J4)=4,"April",MONTH(J4)=5,"May",MONTH(J4)=6,"June",MONTH(J4)=7,"July",MONTH(J4)=8,"August",MONTH(J4)=9,"September",MONTH(J4)=10,"October",MONTH(J4)=11,"November",MONTH(J4)=12,"December"))');
    ss.getSheetByName('Raw Data').getRange('R4').setFormula('=IF(J4="","",IF(J4="TBD","TBD",year(J4)))');
    ss.getSheetByName('Raw Data').getRange('S4').setFormula('=IFS(N4="TBD","TBD",N4="","",N4>0,O4&" "&N4)');
    ss.getSheetByName('Raw Data').getRange('AA4').setValue('=TODAY()')
    ss.getSheetByName('Raw Data').getRange('AA4').setNumberFormat('m"/"d" "h":"mma/p')
    var date = ss.getSheetByName('Raw Data').getRange('AA4').getValue()
    ss.getSheetByName('Raw Data').getRange('AA4').setValue(date)
    ss.getSheetByName('Raw Data').getRange('AB4').setValue(notes)
    
    // clear Buyer Info inputs and redo formatting
    ss.getRange('I5:I11').clear({contentsOnly: true})
    
    // clear city, zip, and miles distances for each agent
    ss.getRange('E4:E5').clear({contentsOnly: true})
    ss.getRange('E6').setValue(20)
    ss.getRange('G14:G24').clear({contentsOnly: true})
    
    // clear the miles radius filter
    ss.getActiveSheet().getFilter().removeColumnFilterCriteria(7)
    
    // Sort Last Lead Received from oldest to youngest
    ss.getActiveSheet().getFilter().sort(8, true);
    
    // Sort 7-Day Total from least to most
    ss.getActiveSheet().getFilter().sort(9, true);
    
    // redoFormulas(zip)
    
    copyNames()
    ss.getRange('E8').clear({contentsOnly: true})
    agentCellTurnOrange()
    buyerInfoTurnOrange()
  }
}

function updateAgentSS(){
  var ss = SpreadsheetApp.getActive();
  var buyerAgent = ss.getRange('E8').getValue()
  var buyerName = ss.getRange('I5').getValue()
  var buyerPhone = ss.getRange('I6').getValue()
  var buyerEmail = ss.getRange('I7').getValue()
  var listingAgent = ss.getRange('I8').getValue()
  var source = ss.getRange('I9').getValue()
  var tags = ss.getRange('I10').getValue()
  var notes = ss.getRange('I11').getValue()
  var zip = ss.getRange('E5').getValue()
  
  // Get the sheet URL specific to the Buyer Agent assigned
  var agentSS = getAgentSheet(buyerAgent)
  
  // If the buyer agent has a sheet, get the New/Warm Leads tab
  if (agentSS){
    var hotWarmLeads = agentSS.getSheetByName('New/Warm Leads')
    
    hotWarmLeads.insertRowsBefore(hotWarmLeads.getRange('4:4').getRow(), 1);  
    hotWarmLeads.getRange('A4').setValue(buyerName)
    //  hotWarmLeads.getRange('B4').setValue()
    //  hotWarmLeads.getRange('C4').setValue()
    hotWarmLeads.getRange('D4').setValue(buyerPhone)
    hotWarmLeads.getRange('E4').setValue(buyerEmail)
    hotWarmLeads.getRange('F4').setValue(listingAgent)
    hotWarmLeads.getRange('G4').setValue('New Lead')
    //    hotWarmLeads.getRange('H4').setValue(600)
    hotWarmLeads.getRange('I4').setValue(source)
    //  hotWarmLeads.getRange('J4').setValue()
    hotWarmLeads.getRange('K4').setValue(buyerAgent)
    hotWarmLeads.getRange('L4').setValue('Open')
    hotWarmLeads.getRange('O4').setFormula('=IF(B4="","",VLOOKUP(B4,Setting!A:B,2,false))')
    hotWarmLeads.getRange('P4').setValue(tags)
    hotWarmLeads.getRange('Q4').setFormula('=IF(J4="","",IFS(J4="TBD","TBD",MONTH(J4)=1,"January",MONTH(J4)=2,"February",MONTH(J4)=3,"March",MONTH(J4)=4,"April",MONTH(J4)=5,"May",MONTH(J4)=6,"June",MONTH(J4)=7,"July",MONTH(J4)=8,"August",MONTH(J4)=9,"September",MONTH(J4)=10,"October",MONTH(J4)=11,"November",MONTH(J4)=12,"December"))');
    hotWarmLeads.getRange('R4').setFormula('=IF(J4="","",IF(J4="TBD","TBD",year(J4)))');
    hotWarmLeads.getRange('S4').setFormula('=IFS(N4="TBD","TBD",N4="","",N4>0,O4&" "&N4)');
    hotWarmLeads.getRange('AA4').setValue('=TODAY()')
    hotWarmLeads.getRange('AA4').setNumberFormat('m"/"d" "h":"mma/p')
    var date = hotWarmLeads.getRange('AA4').getValue()
    hotWarmLeads.getRange('AA4').setValue(date)
    hotWarmLeads.getRange('AB4').setValue(notes)
  }
}

function getAgentSheet(buyerAgent){
  if (buyerAgent === 'Allison Timothy'){
    return 'https://docs.google.com/spreadsheets/d/1HtynDRCk0GyavYUM0enaTnz0weOYC7nwMfbJKfqjLRo/edit'
  }
  else if (buyerAgent === 'Ben Ellis'){
    return 'https://docs.google.com/spreadsheets/d/1RUOVZfM-434oK64nLQTCBwaNTEmAPByRrWUzp5Fjw-I/edit'
  }
  else if (buyerAgent === 'David Greenwood'){
    return 'https://docs.google.com/spreadsheets/d/1XkcKnG-HVjA2M1Rq1rEJnRWIwpoMO06wbzFvJT45SFE/edit'
  }
  else if (buyerAgent === 'Eric Nelson'){
    return 'https://docs.google.com/spreadsheets/d/1SQU6zAsGGvbWaeX1C1AYdHhh37HZB9ADA1aSo075pCw/edit'
  }
  else if (buyerAgent === 'Jake Richins'){
    return 'https://docs.google.com/spreadsheets/d/1m46W2QJNyehTyd8sm_yLX2aCXd1NzkIk3AVtcZrNESY/edit'
  }
  else if (buyerAgent === 'Jamie Johnson'){
    return 'https://docs.google.com/spreadsheets/d/1pie4IfWLlLLxZzR4EMXQpbf58LcBEiUIJTnNQ8YceX0/edit'
  }
  else if (buyerAgent === 'Jeremy Doggett'){
    return 'https://docs.google.com/spreadsheets/d/1e3ex5sBlKqA4BtQqqKuL6_A0KtHRTVpB0Exf85iHfTg/edit'
  }
  else if (buyerAgent === 'JoAnn Ortega-Petty'){
    return 'https://docs.google.com/spreadsheets/d/1X2Sns8dJr01Lv0eBb7RLYBr0L4gTS_p6DDof66LAQ4U/edit'
  }
  else if (buyerAgent === 'Juan Gomez'){
    return 'https://docs.google.com/spreadsheets/d/10DjiST0kxblpoCkyRWJAcYwCnwyBJUqqW3QNTmNwGaM/edit'
  }
  else if (buyerAgent === 'Kodi Paulson'){
    return 'https://docs.google.com/spreadsheets/d/1cRU8l4_miy0kski3XMxSQ8t6ZMzcdxtyVwpImY_qt1U/edit'
  }
  else if (buyerAgent === 'Mike Pembroke'){
    return 'https://docs.google.com/spreadsheets/d/1TIf55QXFo1QO73_qIKS7pGAPcQG49xNy0i0fNdxryXI/edit'
  }
  else if (buyerAgent === 'Taryn Nielsen'){
    return 'https://docs.google.com/spreadsheets/d/1tTjhvHG2Ut-eMlK8s0c_Q4wBotoZ9BDOOnTr5DOqUZg/edit'
  }
  else if (buyerAgent === 'Wyatt Koeven'){
    return 'https://docs.google.com/spreadsheets/d/1_cCisCkYlmX5FL9cHdI_k6RVBjQ12_qPhOjAUDMWo_0/edit'
  }
  else return
}

function lookupCity() {
  var ss = SpreadsheetApp.getActive();
  if (ss.getRange('E5').getValue()){
    ss.getRange('E4').setFormula('=VLOOKUP(E5,\'Utah Zip Codes\'!B2:D391,3,false)');
    return ss.getRange('E4').getValue()
  } else {
    return ''
  }
}

function lookupZip(){
  var ss = SpreadsheetApp.getActive();
  if (ss.getRange('E4').getValue()){
    ss.getRange('E5').setFormula('=VLOOKUP(E4,\'Utah Zip Codes\'!A2:B391,2,false)'); 
    return ss.getRange('E5').getValue()
  } else {
    return ''
  }
}

function lightenCity() {
  var ss = SpreadsheetApp.getActive();
  ss.getRange('D4:E4').setFontColor('#cccccc')
  ss.getRange('E4').setBackground('#fefefe')
  ss.getRange('E4').setBorder(true, true, true, true, null, null, '#f5f5f5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  ss.getRange('D5:E5').setFontColor('#3e494c')
  ss.getRange('E5').setBackground('#fff2cc')
  ss.getRange('E5').setBorder(true, true, true, true, null, null, '#ffe599', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
}

function lightenZip() {
  var ss = SpreadsheetApp.getActive();
  ss.getRange('D5:E5').setFontColor('#cccccc')
  ss.getRange('E5').setBackground('#fefefe')
  ss.getRange('E5').setBorder(true, true, true, true, null, null, '#f5f5f5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  ss.getRange('D4:E4').setFontColor('#3e494c')
  ss.getRange('E4').setBackground('#fff2cc')
  ss.getRange('E4').setBorder(true, true, true, true, null, null, '#ffe599', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
}

function errorBox(cell) {
  var ss = SpreadsheetApp.getActive();
  ss.getRange(cell).setBackground('#f4cccc')
  .setBorder(true, true, true, true, null, null, '#ea9999', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  ss.getRange('H5:I11').setBorder(true, true, true, true, null, null, '#58dbc2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  ss.getRange('H4:I4').setBorder(true, true, true, true, null, null, '#58dbc2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function agentCellTurnGray(){
  var ss = SpreadsheetApp.getActive()
  ss.getRange('E4:E6').setBackground('#f3f3f3')
  .setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  
  ss.getRange('E8:F9').setBackground('#f3f3f3')
  .setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
}

function agentCellTurnOrange(){
  var ss = SpreadsheetApp.getActive()
  
  ss.getRange('E4:E6').setBackground('#fff2cc')
  .setBorder(true, true, true, true, true, true, '#ffe599', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontSize(13)
  .setFontFamily('Arial')
  
  ss.getRange('E8:F9').setBackground('#fff2cc')
  .setBorder(true, true, true, true, true, true, '#ffe599', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontSize(17)
  .setFontFamily('Arial')
}

function buyerInfoTurnGray(){
  var ss = SpreadsheetApp.getActive();
  ss.getRange('I5:I11').setBackground('#f3f3f3')
  .setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  ss.getRange('H5:I11').setBorder(true, true, true, true, null, null, '#58dbc2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function buyerInfoTurnOrange(){
  var ss = SpreadsheetApp.getActive();  
  ss.getRange('I5:I11').setBackground('#fff2cc')
  .setBorder(true, true, true, true, true, true, '#ffe599', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setHorizontalAlignment('left')
  .setVerticalAlignment('middle')
  .setFontSize(11)
  .setFontFamily('Arial');
  ss.getRange('H5:I11').setBorder(true, true, true, true, null, null, '#58dbc2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function clearParameters() {
  agentCellTurnGray()
  var ss = SpreadsheetApp.getActive();
  ss.getRange('E4:E5').clear({contentsOnly: true})
  ss.getRange('E6').setValue(20)
  ss.getRange('E8').clear({contentsOnly: true})
  ss.getRange('G14:G24').clear({contentsOnly: true})
  ss.getActiveSheet().getFilter().removeColumnFilterCriteria(7)
  //  MailApp.sendEmail('mike.degroot@homie.com', 'test', 'This is a test email')
  agentCellTurnOrange()
}