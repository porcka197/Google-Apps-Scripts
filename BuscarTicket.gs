function buscarValor(){
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja1 = libro.getSheetByName("Dashboard");
  var hojaBuscar = libro.getSheetByName("Respuestas de formulario 1");
  var ticket = hoja1.getRange('C62').getValue();
  var tablaBuscar = hojaBuscar.getRange('F3000:NO5805').getValues();
  // Logger.log(ticket);
  // Logger.log(tablaBuscar);
  var lista = tablaBuscar.map(function(fila){return fila[7]});
  // Logger.log(lista);
  var indice = lista.indexOf(ticket);
  // Logger.log(indice);

  var solicitud = tablaBuscar[indice][0];
  // Logger.log(solicitud);
  hoja1.getRange('C63:F63').setValue(solicitud);
  var atendidoPor = tablaBuscar[indice][8];
  hoja1.getRange('C64:F64').setValue(atendidoPor);
  var situacion = tablaBuscar[indice][9];
  hoja1.getRange('C65:F65').setValue(situacion);
  SpreadsheetApp.getActiveSpreadsheet().toast('Solicitud encontrada');
}

function limpiarDatos(){
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C63:F63").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C64:F64").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C65:F65").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C62").clearContent();
}