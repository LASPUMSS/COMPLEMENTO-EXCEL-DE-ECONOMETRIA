// ############################################################################
//             FUNCION PRINCIPAL
// ############################################################################
function correrModeloRLUV(rangoY, rangoX, sigY, sigX){

  try{

  rangoX = rangoX.replace(" ","");
  rangoY = rangoY.replace(" ","");

  var shet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var nColY= shet.getRange(rangoY).getNumColumns();
  var nFilY = shet.getRange(rangoY).getNumRows();
  var nColX = shet.getRange(rangoX).getNumColumns();
  var nFilX = shet.getRange(rangoX).getNumRows();

    if((nColX==nColY) && (nFilX==nFilY)){
      copiarDatos(rangoY, rangoX);
      generacionTablaRLUV();
      determinacionBetas(sigY, sigX);
      SpreadsheetApp.getActive().getActiveSheet().setFrozenRows(1);
    } else if ((nColX!=1) || (nColY!=1)){
      SpreadsheetApp.getUi().alert("El numero de columnas en los rangos es inadecuado.");
    } else if (nFilY != nFilX){
      SpreadsheetApp.getUi().alert("Los rangos deben tener el mismo numero de filas.");
    } 

  } 
  catch(err){

    SpreadsheetApp.getUi().alert(err.message);

  }
 
};


// ############################################################################
//             GENERACION DE LA TABLA RLUV
// ############################################################################

// Funcion principal
function generacionTablaRLUV(){

  titulosIniciales();
  formulasCampos();

};

function titulosIniciales() {

 //ESTOS SON SOLO TODOS LOS TITULOS NECESARIOS PARA REALIZAR LA REGRECION
 //POR LA FORMULA CLASICA Y LA FORMULA POR DESVIOS

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    formatoTitulos(1, "N°",false,"");
    formatoTitulos(2,"X",true,"Este campo representa los valores que tomo o asumio la variable independiente.");
    formatoTitulos(3,"Y",true,"Este campo representa los valores que tomo o asumio la variable dependiente.");
    formatoTitulos(4,"XY",true,"Este campo representa el producto de la variable independiente respecto a la variable dependiente");
    formatoTitulos(5,"X²",true,"Este campo representa el cuadrado de los valores que tomo o asumio la variable independiente.");
    formatoTitulos(6,"xᵢ",true,"Este campo representa las desviaciones (diferencia), de los valores que asumio la independiente respecto a su valor promedio.");
    formatoTitulos(7,"yᵢ",true,"Este campo representa las desviaciones (diferencia), de los valores que asumio la dependiente respecto a su valor promedio.");
    formatoTitulos(8,"xᵢyᵢ",true,"Este campo representa el producto de las desviaciones de la variables independiente y la variable dependiente");
    formatoTitulos(9,"xᵢ²",true,"Este campo representa el cuadrado de las desviaciones, de los valores que tomo o asumio la variable independiente.");
    formatoTitulos(10,"Ŷ",true,"Este campo representa la variable dependiente estimada.");
    formatoTitulos(11,"e",true,"Este campo representa el error que resulta de la diferencia de la variable dependiente estimada respecto la variable dependiente observada.");
    formatoTitulos(12,"Ŷe",true,"Este campo representa el producto de la variable dependiente estimada respecto su error.");
    formatoTitulos(13,"Xe",true,"Este campo representa el producto de la variable independiente estimada respecto su error.");
    formatoTitulos(14,"ŷᵢ",true,"Este campo representa las desviaciones (diferencia), de los valores que asumio la dependiente esperada respecto a su valor promedio.");
    formatoTitulos(15,"e²",true,"Este campo representa el cuadrado del error de estimación.");
    formatoTitulos(16,"yᵢ²",true,"Este campo representa el cuadrado de las desviaciones (diferencia), de los valores que asumio la dependiente respecto a su valor promedio.");
    formatoTitulos(17,"ŷᵢ²",true,"Este campo representa el cuadrado de las las desviaciones (diferencia), de los valores que asumio la dependiente estiamada respecto a su valor promedio.");
    formatoTitulos(18,"Y²",true,"Este campo representa el cuadrado de los valores que tomo o asumio la variable dependiente.");
    formatoTitulos(19,"yᵢŷᵢ",true,"Este campo representa el producto, de las desviaciones de la variable dependiente respecto a las desviaciones de la variable dependiente estimada.");

    ss.getRange('B:C').setNumberFormat('#,##0.00');
    ss.getRange('A:A').setNumberFormat('#,##0').setHorizontalAlignment('center');


};

function formulasCampos(){

  var ss = SpreadsheetApp.getActive();

  // REPRESENTA EL PRODUCTO DE LA VARIABLE INDEPENDIENTE Y DEPENDIENTE
  formatoFormulas(4,'R[0]C[-2]*R[0]C[-1]');
  // REPRESENTA EL PRODUCTO CUADRADO DE LA VARIABLE INDEPENDIENTE
  formatoFormulas(5,'R[0]C[-3]^2');
  // SE DETERMINA LOS DEVIOS DE LAS OBSERVACIONES EN "X", RESPECTO A SU PROMEDIO
  formatoFormulas2(6,'R[0]C[-4]-R',2,2);
  // SE DERTERMINA LOS DESVIOS DE LAS OBSERVACIONES EN "Y", RESPECTO A SU PROMEDIO
  formatoFormulas2(7,'R[0]C[-4]-R',3,2);
  // SE DETERMINARA EL PRODUCTO DE LOS DESVIOS DE "X" Y "Y"
  formatoFormulas(8,'R[0]C[-2]*R[0]C[-1]');
  // SE DETERMINA EL CUADRADO DE LOS DESVIOS EN X
  formatoFormulas(9,'R[0]C[-3]^2');
  // SE DETERMINA Y ESTIMADA
  formatoFormulaYestimada();
  // SE DETERMINA EL ERROR DE LA ESTIMACION
  formatoFormulas(11,'=R[0]C[-8]-R[0]C[-1]');
  // SE DETERMINA EL PRODUCTO DE Y ESTIMADA RESPECTO AL ERROR
  formatoFormulas(12,'=R[0]C[-1]*R[0]C[-2]');
  // SE DETERMINA EL PRODUCTO DE X RESPECTO AL ERROR
  formatoFormulas(13,'=R[0]C[-11]*R[0]C[-2]');
  // SE DETERMINA LOS DESVIOS DE Y ESTIMADA
  formatoFormulaDesvYestimada();
  // SE DETERMINA EL ERROR AL CUADRADA
  formatoFormulas(15,'=R[0]C[-4]^2');
  // SE DETERMINA LOS DEVIOS DE Y ESTIMADA AL CUADRADO
  formatoFormulas(16,'=R[0]C[-9]^2');
  // SE DETERMINA LOS DEVIOS DE Y AL CUADRADO
  formatoFormulas(17,'=R[0]C[-3]^2');
  // SE DETERMINA Y AL CUADRADO
  formatoFormulas(18,'=R[0]C[-15]^2');
  // SE DETERMINA EL PRODUCTO DE DESVIOS DE Y RESPECTO A DESVIOS DE Y ESTIMADA
  formatoFormulas(19,'=R[0]C[-12]*R[0]C[-5]');

  ss.getRange('D2:S2').activate();
  ss.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  ss.getRange(`A2:S${ss.getLastRow()}`).setBorder(null, true, false, true, true, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // SUMAS TOTALES DE LOS DIFERENTES CAMPOS
  var nfila = ss.getLastRow();
  for(var i=2; i<20; i++){
    formatoSumasVariables(nfila,i);
  }

  // PROMEDIO DE LOS CAPOS CORRESPONDIENETES
  var nfila = ss.getLastRow();
  for(var i=2; i<4; i++){
    formatoPromedioVariables(nfila,i);
  }
  
  // ETIQUETAS SRC, STC, SEC
  var nfila = ss.getLastRow();
  formatoEtiquetas(nfila, 15,"SRC");
  formatoEtiquetas(nfila, 16,"STC");
  formatoEtiquetas(nfila, 17,"SEC");
  
};









