function buscarValor(){
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja1 = libro.getSheetByName("Dashboard");
  var hojaBuscar = libro.getSheetByName("Respuestas de formulario 1");
  var ticket = hoja1.getRange('C71').getValue();
  var tablaBuscar = hojaBuscar.getRange('F1:M5805').getValues();
  // Logger.log(ticket);
  // Logger.log(tablaBuscar);
  var lista = tablaBuscar.map(function(fila){return fila[7]});
  // Logger.log(lista);
  var indice = lista.indexOf(ticket);
  // Logger.log(indice);

  var solicitud = tablaBuscar[indice][0];
  // Logger.log(solicitud);
  hoja1.getRange('C72:F72').setValue(solicitud);
  SpreadsheetApp.getActiveSpreadsheet().toast('Solicitud encontrada');
}