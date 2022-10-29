// ############################################################################
//             FUNCION PRINCIPAL
// ############################################################################
function correrModeloRLVV(rangoY, rangoX){

  try{

  rangoX = rangoX.replace(" ","");
  rangoY = rangoY.replace(" ","");

  var shet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var nColY= shet.getRange(rangoY).getNumColumns();
  var nFilY = shet.getRange(rangoY).getNumRows();
  var nColX = shet.getRange(rangoX).getNumColumns();
  var nFilX = shet.getRange(rangoX).getNumRows();

    if((nColY==1) && (nFilX==nFilY)){

      SpreadsheetApp.getUi().alert("CORRER MODELO");

    } else if (nColY>1){
      SpreadsheetApp.getUi().alert("El numero de columnas del rango Y es inadecuado.");
    } else if (nFilY != nFilX){
      SpreadsheetApp.getUi().alert("Los rangos deben tener el mismo numero de filas.");
    } 

  } 
  catch(err){

    SpreadsheetApp.getUi().alert(err.message);

  }
 
};
