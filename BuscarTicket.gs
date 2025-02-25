/**-----------------------------------Buscar tickets ---------------------------------------- */
function buscarValor() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja1 = libro.getSheetByName("Dashboard");
  var hojaBuscar = libro.getSheetByName("Respuestas de formulario 1");
  var filaUlt = hojaBuscar.getLastRow();
  var ticket = hoja1.getRange('C62').getValue();
  var tablaBuscar = hojaBuscar.getRange('F1645:Q' + filaUlt).getValues();

  // Logger.log(typeof(ticket));
  // Logger.log(ticket);
  // Logger.log(tablaBuscar);
  var lista = tablaBuscar.map(function (fila) { return fila[8] });
  // Logger.log(lista);
  var indice = lista.indexOf(ticket);
  // Logger.log(indice);

  if (typeof (ticket) == "string") {
    SpreadsheetApp.getUi().alert('Debes colocar el número ticket para buscar');
    limpiarDatos2();
  }

  else if (typeof (ticket) == "number") {
    var solicitud = tablaBuscar[indice][0];
    // Logger.log(solicitud);
    hoja1.getRange('C63:F63').setValue(solicitud);
    var atendidoPor = tablaBuscar[indice][10];
    hoja1.getRange('C64:F64').setValue(atendidoPor);
    var situacion = tablaBuscar[indice][11];
    hoja1.getRange('C65:F65').setValue(situacion);
    SpreadsheetApp.getActiveSpreadsheet().toast('Solicitud encontrada');
  }
  else if (ticket == "") {
    SpreadsheetApp.getUi().alert('Debes colocar el número ticket para buscar');
    limpiarDatos2();
  }
}

/**-----------------------------------Borrar datos ---------------------------------------- */

function limpiarDatos() {
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C63:F63").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C64:F64").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C65:F65").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C62").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().toast('Nueva búsqueda');
}

/**-----------------------------------Borrar datos2 ---------------------------------------- */

function limpiarDatos2() {
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C63:F63").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C64:F64").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C65:F65").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("C62").clearContent();
  SpreadsheetApp.getActiveSpreadsheet().toast('Busca el ticket correctamente');
}
