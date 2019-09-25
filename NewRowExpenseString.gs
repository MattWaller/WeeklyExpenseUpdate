function WeeklyExpenses() {
  // comes from Monthly Expenses
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SPP Order Sheet");
  var ssLR = ss.getLastRow()+1;
  var service = ss.getRange(ssLR,3).setValue('Monthly Expenses (Prorated Weekly)')
  var price = ss.getRange(ssLR,4).setFormula("=-SUM('Monthly Expenses'!E:E)")
  var Rebuilding_Fund = ss.getRange(ssLR,5).setFormula('=if(INDIRECT(\"R[0]C[-2]\", FALSE)<>\"\",INDIRECT(\"R[0]C[-1]\", FALSE)*\'Sells Sheet\'!$J$2,VLOOKUP(INDIRECT(\"R[0]C[-2]\", FALSE),\'Sells Sheet\'!C:J,8))')
  var Jonathan_Payable = ss.getRange(ssLR,6).setFormula('=if(INDIRECT(\"R[0]C[-3]\", FALSE)<>\"\",INDIRECT(\"R[0]C[-2]\", FALSE)*\'Sells Sheet\'!$M$2,VLOOKUP(INDIRECT(\"R[0]C[-3]\", FALSE),\'Sells Sheet\'!C:M,11))')
  var Leith_Payable = ss.getRange(ssLR,7).setFormula('=if(INDIRECT(\"R[0]C[-4]\", FALSE)<>\"\",INDIRECT(\"R[0]C[-3]\", FALSE)*\'Sells Sheet\'!$L$2,VLOOKUP(INDIRECT(\"R[0]C[-4]\", FALSE),\'Sells Sheet\'!C:M,10))')
  var Date = ss.getRange(ssLR,8).setFormula('=ROUNDDOWN(indirect("r[0]c[-7]",FALSE),0)')
  var today = ss.getRange(ssLR,1).setFormula('=today()')
  // To stringify data
 // Logger.log(ssLR);
  var Str_date = ss.getRange(ssLR,8).getValue();
  var Str_price = ss.getRange(ssLR,4).getValue();
  var Str_today = ss.getRange(ssLR,1).getValue();
  var Str_RF = ss.getRange(ssLR,5).getValue();
  var Str_JP = ss.getRange(ssLR,6).getValue();
  var Str_LP = ss.getRange(ssLR,7).getValue();
  
  // setting stringed values into sheet
  ss.getRange(ssLR,8).setValue(Str_date);
  ss.getRange(ssLR,4).setValue(Str_price);
  ss.getRange(ssLR,1).setValue(Str_today);
  ss.getRange(ssLR,5).setValue(Str_RF);
  ss.getRange(ssLR,6).setValue(Str_JP);
  ss.getRange(ssLR,7).setValue(Str_LP);
  //Logger.log(Str_date);
}
