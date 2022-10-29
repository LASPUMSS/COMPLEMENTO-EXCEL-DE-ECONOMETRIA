function generarGraficoDispersion(rangoY, rangoX, sigY, sigX) {

  try{

  rangoX = rangoX.replace(" ","");
  rangoY = rangoY.replace(" ","");

  var shet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var nColY= shet.getRange(rangoY).getNumColumns();
  var nFilY = shet.getRange(rangoY).getNumRows();
  var nColX = shet.getRange(rangoX).getNumColumns();
  var nFilX = shet.getRange(rangoX).getNumRows();

    if((nColX==nColY) && (nFilX==nFilY)){

      var sheet = SpreadsheetApp.getActiveSheet();
      var chartBuilder = sheet.newChart();

      chartBuilder
        .addRange(sheet.getRange(rangoX))
        .addRange(sheet.getRange(rangoY))
        .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
        .setChartType(Charts.ChartType.SCATTER)
        .setOption('title', 'Grafico De Dispersión')
        .setOption('titleTextStyle.color', '#000000')
        .setOption('titleTextStyle.bold', true)
        .setOption('titleTextStyle.alignment', 'center')
        .setOption('trendlines.n.visibleInLegend', true)
        .setOption('hAxis.title', sigX)
        .setOption('vAxis.title', sigY)
        .setOption('trendlines.0.visibleInLegend', true)
        .setOption('trendlines.0.type', 'linear')
        .setPosition(5, 5, 0, 0);
      sheet.insertChart(chartBuilder.build());
      
      var spreadsheet = SpreadsheetApp.getActive();
      var sheet = spreadsheet.getActiveSheet();
      var charts = sheet.getCharts();
      var chart = charts[charts.length - 1];

      spreadsheet.moveChartToObjectSheet(chart);

    } else if ((nColX!=1) || (nColY!=1)){
      SpreadsheetApp.getUi().alert("El numero de columnas en los rangos es inadecuado.");
    } else if (nFilY != nFilX){
      SpreadsheetApp.getUi().alert("Los rangos deben tener el mismo numero de filas.");
    } 

  } 
  catch(err){

    SpreadsheetApp.getUi().alert(err.message);
    
  }

  

  
}
