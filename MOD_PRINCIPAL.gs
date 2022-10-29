function onInstall(e){
  onOpen(e);
}

function onOpen(e){
  miMenu();
};

function miMenu() {
  var menu = SpreadsheetApp.getUi();
  menu.createMenu('Econometria')

  //  .addItem('ANALISIS RAPIDO','analisisRapido')

    .addSeparator()

    .addSubMenu(SpreadsheetApp.getUi().createMenu('A UNA VARIABLE INDEPENDIENTE')
      .addItem('Grafico De Dispersión','graficoDispersion')
      .addItem('Regresión Lineal MCO','regresionMCOUV')

      .addSeparator()

    //  .addSubMenu(SpreadsheetApp.getUi().createMenu('EVALUACION INDIVIDUAL DEL MODELO GENERADO')
    //    .addItem('Generar Hoja Para Evaluar El Modelo','generarHojaEvalMod')
    //    .addItem('Hipotesis Simple','xxxxxxx')
    //    .addItem('Intervalo De Confianza','xxxxxxx')
    //    .addItem('Cuadro ANOVA','xxxxxxx'))

    //  .addSubMenu(SpreadsheetApp.getUi().createMenu('EVALUACION GENERAL DEL MODELO GENERADO')
    //    .addItem('Generar Hoja Para Evaluar El Modelo','xxxxxxx')
    //    .addItem('Hipotesis Simple','xxxxxxx')
    //    .addItem('Intervalo De Confianza','xxxxxxx')
    //    .addItem('Cuadro ANOVA','xxxxxxx'))
    )

    //.addSeparator()

    //.addSubMenu(SpreadsheetApp.getUi().createMenu('A VARIAS VARIABLES INDEPENDIENTES')
    //  .addItem('Regresion Lineal MCO','regresionMCOVV'))

    .addToUi();
};

function analisisRapido(){

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('formAR').evaluate()
      .setTitle('Modelo MCO')
  );

}

function graficoDispersion(){

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('formGD').evaluate()
      .setTitle('Grafico De Dispersión')
  );

};

function regresionMCOUV(){

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('formRLUV').evaluate()
      .setTitle('Modelo MCO')
  );

};

function regresionMCOVV(){

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createTemplateFromFile('formRLVV').evaluate()
      .setTitle('Modelo MCO')
  );

};
