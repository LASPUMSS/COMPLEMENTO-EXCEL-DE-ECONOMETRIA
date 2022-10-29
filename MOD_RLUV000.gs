function numeracionObs(){

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var nOb = ss.getLastRow();
  for(var i = 2; i < nOb+1; i++){
     ss.getRange(i,1).setValue(i-1);
  }
  
};

function copiarDatos(rangoY, rangoX){

  var spreadsheet = SpreadsheetApp.getActive();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssFuenteDatos = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var ssCopiaDatos = ss.insertSheet(1);
  var nObs = ssFuenteDatos.getLastRow();
  var diferencia = nObs - ssCopiaDatos.getMaxRows();

  if(diferencia>0){
    var incremento = diferencia + 500
  }

  if(nObs>999){
    ss.setActiveSheet(ssCopiaDatos, true);
    spreadsheet.getActiveSheet().insertRowsAfter(spreadsheet.getActiveSheet().getMaxRows(), incremento);
  }

  ssFuenteDatos.getRange(`\'${ssFuenteDatos.getName()}\'!${rangoY}`).copyTo(ssCopiaDatos.getRange("C2"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  ssFuenteDatos.getRange(`\'${ssFuenteDatos.getName()}\'!${rangoX}`).copyTo(ssCopiaDatos.getRange("B2"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
  ss.setActiveSheet(ssCopiaDatos, true);
  numeracionObs();

};

// FORMATOS CELDAS

function formatoTitulos(nCol, titulo, seCom, comentario){

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ss.getRange(1,nCol).setValue(titulo)
      .setBackground('#6fa8dc')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  if(seCom){
    if(comentario==""){
      comentario = "Aqui se argumentara el significado de la variable correspondiente."
    }
    ss.getRange(1,nCol).setNote(comentario);
  }

};

function formatoEtiquetas(nFil , nCol, titulo){

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ss.getRange(nFil,nCol).setValue(titulo)
      .setFontColor('#000000')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

};

function formatoFormulas(nCol, formula){

   var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    ss.getRange(2,nCol).setFormulaR1C1(formula);
    ss.getRange(2,nCol).setNumberFormat('#,##0.00');

};

function formatoFormulas2(nCol, formula, colXY, sumFil){

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var filaX = ss.getLastRow() + sumFil;
  ss.getRange(2,nCol).setFormulaR1C1(formula + filaX + 'C' + colXY);
  ss.getRange(2,nCol).setNumberFormat('#,##0.000000');

};

function formatoFormulaYestimada(){

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var filaX = ss.getLastRow();
  var filBeta1 = filaX + 9
  var filBeta2 = filaX + 8
  ss.getRange(2,10).setFormulaR1C1('=R' + filBeta1 + 'C2+R' + filBeta2 + 'C2*R[0]C[-8]');
  ss.getRange(2,10).setNumberFormat('#,##0.000000');

};

function formatoFormulaDesvYestimada(){

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var filaX = ss.getLastRow();
  var filaY = filaX + 2
  ss.getRange(2,14).setFormulaR1C1('=R[0]C[-4]-R' + filaY + 'C3');
  ss.getRange(2,14).setNumberFormat('#,##0.000000');

};

function formatoSumasVariables(fila, nCol){

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdaSuma = ss.getRange(fila + 1, nCol).activate();

    // Titulo Fila
    ss.getRange(fila + 1,1).setValue("SUMA")
      .setBackground('#6fa8dc')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Formula
    celdaSuma.setFormulaR1C1('=SUM(R[-' + (fila - 1) + ']C[0]:R[-1]C[0])');
    celdaSuma.setNumberFormat('#,##0.00');
    
    celdaSuma
      .setBackground('#6fa8dc')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    
};

function formatoPromedioVariables(fila, nCol){

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdaSuma = ss.getRange(fila + 1, nCol).activate();

    // Titulo Fila
    ss.getRange(fila + 1,1).setValue("PROMEDIO")
      .setBackground('#6fa8dc')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Formula
    celdaSuma.setFormulaR1C1('=AVERAGE(R[-' + (fila - 1) + ']C[0]:R[-2]C[0])');
    celdaSuma.setNumberFormat('#,##0.00');
    
    celdaSuma
      .setBackground('#6fa8dc')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    
};


