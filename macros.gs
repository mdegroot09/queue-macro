function IR() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getRange('4:4').getRow(), 1);
  spreadsheet.getRange('C4').setValue('Lead')
  spreadsheet.getRange('D4').setValue(600)
  spreadsheet.getRange('H4').setValue('Open')
  spreadsheet.getRange('K4').setFormula('=IF(B4="","",VLOOKUP(B4,Setting!A:B,2,false))');
  spreadsheet.getRange('N4').setFormula('=IF(F4="","",IFS(F4="TBD","TBD",MONTH(F4)=1,"January",MONTH(F4)=2,"February",MONTH(F4)=3,"March",MONTH(F4)=4,"April",MONTH(F4)=5,"May",MONTH(F4)=6,"June",MONTH(F4)=7,"July",MONTH(F4)=8,"August",MONTH(F4)=9,"September",MONTH(F4)=10,"October",MONTH(F4)=11,"November",MONTH(F4)=12,"December"))');
  spreadsheet.getRange('O4').setFormula('=IF(F4="","",IF(F4="TBD","TBD",year(F4)))');
  spreadsheet.getRange('P4').setFormula('=IFS(N4="TBD","TBD",N4="","",N4>0,O4&" "&N4)');
  spreadsheet.getRange('Q4').setValue('=TODAY()')
  spreadsheet.getRange('Q4').setNumberFormat('m"/"d"/"yy')
  var date = spreadsheet.getRange('Q4').getValue()
  spreadsheet.getRange('Q4').setValue(date)
  spreadsheet.getRange('A4').activate()
}