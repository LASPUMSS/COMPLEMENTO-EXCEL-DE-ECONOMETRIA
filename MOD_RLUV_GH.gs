function generarHojaEvalMod() {
  var ss = SpreadsheetApp.getActive();
  var shetModelo = ss.getActiveSheet();
  var nFilCelInc = shetModelo.getLastRow()-34;
  var rangoBetas = shetModelo.getRange(nFilCelInc,1,6,2).getA1Notation();
  var rangoDeter = shetModelo.getRange(nFilCelInc,10,19,3).getA1Notation();
  var rangoSRC = shetModelo.getRange(nFilCelInc-3,15).getA1Notation();
  var rangoSEC = shetModelo.getRange(nFilCelInc-3,17).getA1Notation();
  var shetHojaEval = ss.insertSheet(1);

  shetModelo.getRange(`\'${shetModelo.getName()}\'!${rangoBetas}`).copyTo(shetHojaEval.getRange("A1"), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  shetModelo.getRange(`\'${shetModelo.getName()}\'!${rangoDeter}`).copyTo(shetHojaEval.getRange("A8"), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  shetModelo.getRange(`\'${shetModelo.getName()}\'!${rangoSRC}`).copyTo(shetHojaEval.getRange("B28"), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  shetModelo.getRange(`\'${shetModelo.getName()}\'!${rangoSEC}`).copyTo(shetHojaEval.getRange("B29"), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  shetHojaEval.getRange("A28").setValue("SRC");
  shetHojaEval.getRange("A29").setValue("SEC");

  shetHojaEval.getRange("E1").setValue("PRUEBAS DE HIPOTESIS, TABLAS RESUMEN, CUADROS ANOVA, P VALOR Y INTERVALOS DE CONFIANZA");

  ss.setActiveSheet(shetHojaEval, true);
  
  
}
