function recogerNombreCeldas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangoCeldas = ss.getActiveRange().getA1Notation();
  return rangoCeldas;
};

function include( nameHtml ){
  return HtmlService.createHtmlOutputFromFile( nameHtml ).getContent();
};