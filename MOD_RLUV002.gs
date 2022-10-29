// #############################################################################
//                            DETERMINACION DE BETAS
// #############################################################################
function determinacionBetas(sigY, sigX){
  titulosBetas(sigY, sigX);
  detFormaClasica();
  detPorDesvios();
};

function titulosBetas(sigY, sigX) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    ss.getRange(1,1).activate();
  var celdaInicio = ss.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(2,0).activate();

    celdaInicio.setValue("DETERMINACIÓN DE BETAS");
    celdaInicio.setFontWeight('bold').setHorizontalAlignment('left');    
    celdaInicio.offset(1,0).setValue("Forma Clasica:");
    celdaInicio.offset(1,0).setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(2,0).setValue("Numerador");
    celdaInicio.offset(2,0).setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(3,0).setValue("Denominador");
    celdaInicio.offset(3,0).setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(4,0).setValue("Beta 2");
    celdaInicio.offset(4,0).setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(5,0).setValue("Beta 1");
    celdaInicio.offset(5,0).setFontWeight('bold').setHorizontalAlignment('left');

    celdaInicio.offset(1,3).setValue("Por Desvios:");
    celdaInicio.offset(1,3).setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(2,3).setValue("Beta 2");
    celdaInicio.offset(2,3).setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(3,3).setValue("Beta 1");
    celdaInicio.offset(3,3).setFontWeight('bold').setHorizontalAlignment('left');

    celdaInicio.offset(4,3).setValue("OBJETIVO");
    celdaInicio.offset(4,3).setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(5,3).setValue("Explicar:");
    celdaInicio.offset(5,3).setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(6,3).setValue("Mediante:");
    celdaInicio.offset(6,3).setFontWeight('bold').setHorizontalAlignment('left');
    
    celdaInicio.offset(5,4).setValue(sigY);
    celdaInicio.offset(5,4).setHorizontalAlignment('left');
    celdaInicio.offset(6,4).setValue(sigX);
    celdaInicio.offset(6,4).setHorizontalAlignment('left');

    // ############################################################################

    celdaInicio.offset(0,9).setValue("DETERMINACIÓN DE VARIANZA Y DESVIACIÓN ESTANDAR.")
      .setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(1,9).setValue("PARA LAS DESVIACIONES:")
      .setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(2,9).setValue("n-k")
      .setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(3,9).setValue("Varianza u");
    celdaInicio.offset(4,9).setValue("Err. Estandar");
    celdaInicio.offset(5,9).setValue("PARA BETA 2:")
      .setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(6,9).setValue("Varianza");
    celdaInicio.offset(7,9).setValue("Err. Estandar");
    celdaInicio.offset(8,9).setValue("PARA BETA 1:")
      .setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(9,9).setValue("Varianza");
    celdaInicio.offset(10,9).setValue("Err. Estandar");

    celdaInicio.offset(12,9).setValue("DETERMINACIÓN DEL COEFICIENTE DE DETERMINACIÓN:")
      .setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(13,9).setValue("R^2");

    celdaInicio.offset(15,9).setValue("DETERMINACIÓN DEL COEFICIENTE DE CORRELACIÓN:")
      .setFontWeight('bold').setHorizontalAlignment('left');
    celdaInicio.offset(16,9).setValue("r");
    celdaInicio.offset(17,9).setValue("r");
    celdaInicio.offset(18,9).setValue("r");
    celdaInicio.offset(16,11).setValue("Formula clasica.");
    celdaInicio.offset(17,11).setValue("Formula por desvios.");
    celdaInicio.offset(18,11).setValue("Atajo desde R^2.");

    // ############################################################################

     celdaInicio.offset(7 ,0).setValue("PROPIEDADES:")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(8 ,0).setValue("1ERA PROPIEDAD: Que el promedio de Y observada es igual a Y estimada.")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(13 ,0).setValue("2DA PROPIEDAD: Que el promedio de Y observada es igual al promedio de Y estimada.")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(14 ,0).setValue("O su equivalente que la suma de Y observada es igual  la suma de Y estmada.")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(19 ,0).setValue("3RA PROPIEDAD: Que el promedio de los residuos es igual a cero.")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(25 ,0).setValue("4TA PROPIEDAD: Que la suma de los productos resultantes de Y estimada por sus desviaciones correspondientes, es igual a cero.")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(26 ,0).setValue("Los residuos no tienen niguna relación con el valor estimado.")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(31 ,0).setValue("5TA PROPIEDAD: Que la suma de los productos resultantes de X observada por sus desviaciones correspondientes, es igual a cero.")
      .setFontWeight('bold').setHorizontalAlignment('left');

     celdaInicio.offset(10 ,1).setValue("Prom. Y.")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(16 ,1).setValue("Suma Y")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(22 ,1).setValue("Suma de e.")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(28 ,1).setValue("Suma de Y estimada por e.")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(33 ,1).setValue("Suma de X estimada por e.")
      .setFontWeight('bold').setHorizontalAlignment('left');

     celdaInicio.offset(10 ,2).setValue("Y estimada.")
      .setFontWeight('bold').setHorizontalAlignment('left');
     celdaInicio.offset(16 ,2).setValue("Suma Y estimada")
      .setFontWeight('bold').setHorizontalAlignment('left');

};

function detFormaClasica(){
  // SE CALCULA LOS BETAS POR EL METODO CLASICO.

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      ss.getRange(1,1).activate();
  var celInicial = ss.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(2,0).activate();
      celInicial.offset(2,1).setFormulaR1C1('R[-5]C[2]-R[-6]C[-1]*R[-4]C[0]*R[-4]C[1]')
        .setNumberFormat('#,##0.00')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');

      celInicial.offset(3,1).setFormulaR1C1('R[-6]C[3]-R[-7]C[-1]*R[-5]C[0]^2')
        .setNumberFormat('#,##0.00')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');

      celInicial.offset(4,1).setFormulaR1C1('R[-2]C[0]/R[-1]C[0]')
        .setNumberFormat('#,##0.000000')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');

      celInicial.offset(5,1).setFormulaR1C1('R[-7]C[1]-R[-1]C[0]*R[-7]C[0]')
        .setNumberFormat('#,##0.000000')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');

    // ############################################################################

      celInicial.offset(2,10).setFormulaR1C1("=R[-6]C[-10]-2")
        .setNumberFormat('#,##0')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');
      celInicial.offset(3,10).setFormulaR1C1("=R[-6]C[4]/R[-1]C[0]")
        .setNumberFormat('#,##0.00')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');
      celInicial.offset(4,10).setFormulaR1C1("=SQRT(R[-1]C[0])")
        .setNumberFormat('#,##0.00')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');

      celInicial.offset(6,10).setFormulaR1C1("=R[-3]C[0]/R[-9]C[-2]")
        .setNumberFormat('#,##0.000000')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');
      celInicial.offset(7,10).setFormulaR1C1("=SQRT(R[-1]C[0])")
        .setNumberFormat('#,##0.000000')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');

      celInicial.offset(9,10).setFormulaR1C1("=(R[-6]C[0]*R[-12]C[-6])/(R[-13]C[-10]*R[-12]C[-2])")
        .setNumberFormat('#,##0.000000')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');
      celInicial.offset(10,10).setFormulaR1C1("=SQRT(R[-1]C[0])")
        .setNumberFormat('#,##0.000000')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');

      celInicial.offset(13,10).setFormulaR1C1("=R[-16]C[6]/R[-16]C[5]")
        .setNumberFormat('#,##0.000000')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');

      celInicial.offset(16,10).setFormulaR1C1("=(R[-20]C[-10]*R[-19]C[-7]-R[-19]C[-9]*R[-19]C[-8])/(SQRT((R[-20]C[-10]*R[-19]C[-6]-(R[-19]C[-9])^2)*(R[-20]C[-10]*R[-19]C[7]-(R[-19]C[-8])^2)))")
        .setNumberFormat('#,##0.000000')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');
      celInicial.offset(17,10).setFormulaR1C1("=(R[-20]C[8])/(SQRT(R[-20]C[5]*R[-20]C[6]))")
        .setNumberFormat('#,##0.000000')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');
      celInicial.offset(18,10).setFormulaR1C1("=SQRT(R[-5]C[0])")
        .setNumberFormat('#,##0.000000')
        .setFontWeight('bold')
        .setFontColor('#4a86e8');

    // ############################################################################

      celInicial.offset(11 ,0).setFormulaR1C1('=IF(R[0]C[1]=R[0]C[2];"Se Cumple.";"No Cumple.")')
        .setHorizontalAlignment('left');
      celInicial.offset(17 ,0).setFormulaR1C1('=IF(R[0]C[1]=R[0]C[2];"Se Cumple.";"No Cumple.")')
        .setHorizontalAlignment('left');
      celInicial.offset(23 ,0).setFormulaR1C1('=IF(R[0]C[1]=0;"Se Cumple.";"No Cumple.")')
        .setHorizontalAlignment('left');
      celInicial.offset(29 ,0).setFormulaR1C1('=IF(R[0]C[1]=0;"Se Cumple.";"No Cumple.")')
        .setHorizontalAlignment('left');
      celInicial.offset(34 ,0).setFormulaR1C1('=IF(R[0]C[1]=0;"Se Cumple.";"No Cumple.")')
        .setHorizontalAlignment('left');

      celInicial.offset(11 ,1).setFormulaR1C1('=ROUND(R[-13]C[1];2)')
        .setHorizontalAlignment('center');
      celInicial.offset(17 ,1).setFormulaR1C1('=ROUND(R[-20]C[1];2)')
        .setHorizontalAlignment('center');
      celInicial.offset(23 ,1).setFormulaR1C1('=ROUND(R[-26]C[9];2)')
        .setHorizontalAlignment('center');
      celInicial.offset(29 ,1).setFormulaR1C1('=ROUND(R[-32]C[10];2)')
        .setHorizontalAlignment('center');
      celInicial.offset(34 ,1).setFormulaR1C1('=ROUND(R[-37]C[11];2)')
        .setHorizontalAlignment('center');

      celInicial.offset(11 ,2).setFormulaR1C1('=ROUND((R[-6]C[-1]+R[-7]C[-1]*R[-13]C[-1]);2)')
        .setHorizontalAlignment('center');
      celInicial.offset(17 ,2).setFormulaR1C1('=ROUND(R[-20]C[7];2)')
        .setHorizontalAlignment('center');

};

function detPorDesvios(){
  // SE CALCULA LOS BETAS POR LOS DESVIOS EN "X" Y "Y".

  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    ss.getRange(1,1).activate();
  var celInicial = ss.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).offset(2,0).activate();
    celInicial.offset(2,4).setFormulaR1C1('R[-5]C[3]/R[-5]C[4]');
    celInicial.offset(2,4).setNumberFormat('#,##0.000000');
    celInicial.offset(2,4).setFontWeight('bold');
    celInicial.offset(2,4).setFontColor('#4a86e8');

    celInicial.offset(3,4).setFormulaR1C1('R[-5]C[-2]-R[-1]C[0]*R[-5]C[-3]');
    celInicial.offset(3,4).setNumberFormat('#,##0.000000');
    celInicial.offset(3,4).setFontWeight('bold');
    celInicial.offset(3,4).setFontColor('#4a86e8');
};

