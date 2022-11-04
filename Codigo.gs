//Función para agregar ticket a Formulario
function addSequenceNumber() {
  // Obtain the sheet where we save the answers
  var sheet = SpreadsheetApp.getActiveSheet();
  // Obtain the last row with data
  var row = SpreadsheetApp.getActiveSheet().getLastRow();
  // Sequence number (record) minus 1, this is due to the headers
  var record = row - 1;
  // Set (or write) the sequence number in the cell specified, change number 4 for the rigth column
  sheet.getRange(row, 12).setValue(record);
  // Return the sequence number
  return record;
}
function sequenceNumberOnFormSubmit(e) {
  // Call the function that generates the sequence number
  var folio = addSequenceNumber();
  var timestamp = e.values[0];
  var mail = e.values[1];
  var campus = e.values[2];
  var name = e.values[3];
  var solicitud = e.values[4];
  var descripcion = e.values[5];
  var captura = e.values[6];
  var origen = e.values[7];
  var destino = e.values[8];
  var alumno = e.values[9];
  var nombreAlumno = e.values[10];


  var subject = `¡Gracias por contactarnos - CSH/Ticket ` + '#' + folio;

  var plain_email_body =
    "Estimado(a) " + name +
    "\n\n" +
    "Resgistramos tu solicitud el día: " + timestamp +
    "\n\n" +
    "Gracias por contactar al Equipo de Soporte Humanitas" +
    "\n\n" +
    "Tu solicitud: " + solicitud +
    "\n\n" +
    "Descripción: " + descripcion +
    "\n\n" +
    "En breve nos pondremos en contacto";
  var html_body = `<center>
          <div style="text-align: center; background-color: #b79b72; width: 100%;">
          <a href ="www.humanitas.edu.mx" target="_blanck"><img style="display: block; margin-left: auto; margin-right: auto;" src="https://clases.universidadhumanitas.edu.mx/Respuestas_Tickets/header.png" alt="Gracias por Contactarnos" width="800px" /></a>
          </div>
          <p style="text-align: center; font-family: Verdana;">Estimado(a) <strong>${name}</strong></p>
          <p style="text-align: center; font-family: Verdana;">Campus: <strong>${campus}</strong></p>
          <p style="text-align: center; font-family: Verdana;">¡Gracias por contactar al Centro de Soporte Humanitas!</p>
          <p style="text-align: center; font-family: Verdana;">Registramos tu solicitud el día ${timestamp}</p>
          <p style="text-align: center; font-family: Verdana;">Tu solictud: <strong>${solicitud}</strong></p>
          <p style="text-align: center; font-family: Verdana;">Descripción: <strong>${descripcion}<br /></strong></p>
          <p style="text-align: center; font-family: Verdana;">¡En breve nos pondremos en Contacto!</p>
          <div style="text-align: center; padding-bottom: 5px; padding-top: 5px;">
          <div>
	        <div style="text-align: center; padding-top: 5px; padding-bottom: 5px;">
          <p style="text-align: center; font-size: small; font-family: Verdana;">Síguenos en:</p>
			    <a href = "https://qrco.de/bco5ww" target = "_blanck"><img src="https://clases.universidadhumanitas.edu.mx/Respuestas_Tickets/GoogleSites/qr.png" width="140px"/></a>
	        </div>           
          <div style="text-align: center; background-color: #b79b72;padding-top: 5px; padding-bottom: 5px;">
          <p style="text-align: center; font-size: x-small; font-family: Verdana;color: white;">Copyright © 2022, Universidad Humanitas, Todos los derechos Reservados.</p>
          </div>
          </div>
          </div>
          </center>`;

          var html_bodyCaptura = `<center>
          <div style="text-align: center; background-color: #b79b72; width: 100%;">
          <a href ="www.humanitas.edu.mx" target="_blanck"><img style="display: block; margin-left: auto; margin-right: auto;" src="https://clases.universidadhumanitas.edu.mx/Respuestas_Tickets/header.png" alt="Gracias por Contactarnos" width="800px" /></a>
          </div>
          <p style="text-align: center; font-family: Verdana;">Estimado(a) <strong>${name}</strong></p>
          <p style="text-align: center; font-family: Verdana;">Campus: <strong>${campus}</strong></p>
          <p style="text-align: center; font-family: Verdana;">¡Gracias por contactar al Centro de Soporte Humanitas!</p>
          <p style="text-align: center; font-family: Verdana;">Registramos tu solicitud el día ${timestamp}</p>
          <p style="text-align: center; font-family: Verdana;">Tu solictud: <strong>${solicitud}</strong></p>
          <p style="text-align: center; font-family: Verdana;">Descripción: <strong>${descripcion}<br /></strong></p>
          <p style="text-align: center; font-family: Verdana;">Descripción: <strong>${captura}<br /></strong></p>
          <p style="text-align: center; font-family: Verdana;">¡En breve nos pondremos en Contacto!</p>
          <div style="text-align: center; padding-bottom: 5px; padding-top: 5px;">
          <div>
	        <div style="text-align: center; padding-top: 5px; padding-bottom: 5px;">
          <p style="text-align: center; font-size: small; font-family: Verdana;">Síguenos en:</p>
			    <a href = "https://qrco.de/bco5ww" target = "_blanck"><img src="https://clases.universidadhumanitas.edu.mx/Respuestas_Tickets/GoogleSites/qr.png" width="140px"/></a>
	        </div>           
          <div style="text-align: center; background-color: #b79b72;padding-top: 5px; padding-bottom: 5px;">
          <p style="text-align: center; font-size: x-small; font-family: Verdana;color: white;">Copyright © 2022, Universidad Humanitas, Todos los derechos Reservados.</p>
          </div>
          </div>
          </div>
          </center>`;

  var html_bodyTraslados = `<center>
          <div style="text-align: center; background-color: #b79b72; width: 100%;">
          <a href ="www.humanitas.edu.mx" target="_blanck"><img style="display: block; margin-left: auto; margin-right: auto;" src="https://clases.universidadhumanitas.edu.mx/Respuestas_Tickets/header.png" alt="Gracias por Contactarnos" width="800px" /></a>
          </div>
          <p style="text-align: center; font-family: Verdana;">Estimado(a) <strong>${name}</strong></p>
          <p style="text-align: center; font-family: Verdana;">¡Gracias por contactar al Centro de Soporte Humanitas!</p>
          <p style="text-align: center; font-family: Verdana;">Registramos tu solicitud el día ${timestamp}</p>
          <p style="text-align: center; font-family: Verdana;">Tu solictud: <strong>${solicitud}</strong></p>
          <p style="text-align: center; font-family: Verdana;">De Campus: <strong>${origen}<br /></strong></p>
          <p style="text-align: center; font-family: Verdana;">A Campus: <strong>${destino}<br /></strong></p>
          <p style="text-align: center; font-family: Verdana;">Matricula: <strong>${alumno}<br /></strong></p>
          <p style="text-align: center; font-family: Verdana;">Nombre: <strong>${nombreAlumno}<br /></strong></p>
		      <p style="text-align: center; font-family: Verdana;">¡En breve nos pondremos en Contacto!</p>
          <div style="text-align: center; padding-bottom: 5px; padding-top: 5px;">
          <div>
	        <div style="text-align: center; padding-top: 5px; padding-bottom: 5px;">
          <p style="text-align: center; font-size: small; font-family: Verdana;">Síguenos en:</p>
			    <a href = "https://qrco.de/bco5ww" target = "_blanck"><img src="https://clases.universidadhumanitas.edu.mx/Respuestas_Tickets/GoogleSites/qr.png" width="140px"/></a>
	        </div>           
          <div style="text-align: center; background-color: #b79b72;padding-top: 5px; padding-bottom: 5px;">
          <p style="text-align: center; font-size: x-small; font-family: Verdana;color: white;">Copyright © 2022, Universidad Humanitas, Todos los derechos Reservados.</p>
          </div>
          </div>
          </div>
          </center>`;

  var advancedOpts = { cc: "csh@humanitas.edu.mx", name: "Centro de Soporte Humanitas", htmlBody: html_body };
  var advanceOpts1 = { cc: "csh@humanitas.edu.mx", name: "Centro de Soporte Humanitas", htmlBody: html_bodyTraslados };
  var advanceOpts2 = { cc: "csh@humanitas.edu.mx", name: "Centro de Soporte Humanitas", htmlBody: html_bodyCaptura };
  if (solicitud == "Traslados") {
    MailApp.sendEmail(mail, subject, plain_email_body, advanceOpts1);
  }
  else if (solicitud != "Traslados") {
    MailApp.sendEmail(mail, subject, plain_email_body, advancedOpts);
  }
  else if (captura != ""){MailApp.sendEmail(mail, subject, plain_email_body, advanceOpts2);}
}


//Función para enviar ticket asignado

function enviarAsignado() {
  const libro1 = SpreadsheetApp.getActiveSpreadsheet();
  libro1.setActiveSheet(libro1.getSheetByName("Respuestas de formulario 1"));
  const hoja1 = SpreadsheetApp.getActiveSheet();
  const filas1 = hoja1.getRange("A4000:Y5020").getValues();

  for (indiceFilas in filas1) {
    var tecnico3 = crearTecnico(filas1[indiceFilas]);
    enviarTicket(tecnico3);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Ticket asignado correctamente');
}

function crearTecnico(datosFila) {
  const tecnico4 = {
    fecha3: datosFila[0],
    mail3: datosFila[1],
    campus3: datosFila[2],
    name3: datosFila[3],
    solicitud3: datosFila[4],
    descripcion3: datosFila[5],
    captura3: datosFila[6],
    campusOrigen3: datosFila[7],
    campusDestino3: datosFila[8],
    matricula3: datosFila[9],
    nameAlumno3: datosFila[10],
    ticket3: datosFila[11],
    folio3: datosFila[12],
    tecnico3: datosFila[13],
    estatus3: datosFila[14],
    atencion3: datosFila[15],
    prior3: datosFila[16],
    colaboracion3: datosFila[17],
    solucion3: datosFila[18],
    enviar3: datosFila[19],
    etiqueta: datosFila[24],
  };
  return tecnico4;

}

function enviarTicket(tecnico4) {
  if (tecnico4.mail3 == "") return;

  if (tecnico4.tecnico3 == "Ricardo Porcayo" && tecnico4.solicitud3 != "Traslados" && tecnico4.enviar3 == "No") {
    const plantilla = HtmlService.createTemplateFromFile('asignado');
    plantilla.tecnico4 = tecnico4;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: "ricardo.porcayo@humanitas.edu.mx",
      subject: "Ticket #" + tecnico4.folio3,
      htmlBody: mensaje
    });
  }

  if (tecnico4.tecnico3 == "Angel Montes" && tecnico4.solicitud3 != "Traslados" && tecnico4.enviar3 == "No") {
    const plantilla = HtmlService.createTemplateFromFile('asignado');
    plantilla.tecnico4 = tecnico4;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: "angel.montes@humanitas.edu.mx",
      subject: "Ticket #" + tecnico4.folio3,
      htmlBody: mensaje
    });
  }

  if (tecnico4.tecnico3 == "Gerardo Omaña" && tecnico4.solicitud3 != "Traslados" && tecnico4.enviar3 == "No") {
    const plantilla = HtmlService.createTemplateFromFile('asignado');
    plantilla.tecnico4 = tecnico4;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: "gerardo.omana@humanitas.edu.mx",
      subject: "Ticket #" + tecnico4.folio3,
      htmlBody: mensaje
    });
  }

  if (tecnico4.tecnico3 == "Andrea Ayala" && tecnico4.solicitud3 != "Traslados" && tecnico4.enviar3 == "No") {
    const plantilla = HtmlService.createTemplateFromFile('asignado');
    plantilla.tecnico4 = tecnico4;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: "andrea.ayala@humanitas.edu.mx",
      subject: "Ticket #" + tecnico4.folio3,
      htmlBody: mensaje
    });
  }

  if (tecnico4.tecnico3 == "Victor Barrera" && tecnico4.solicitud3 != "Traslados" && tecnico4.enviar3 == "No") {
    const plantilla = HtmlService.createTemplateFromFile('asignado');
    plantilla.tecnico4 = tecnico4;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: "victor@humanitas.edu.mx",
      subject: "Ticket #" + tecnico4.folio3,
      htmlBody: mensaje
    });
  }

  //*****************************************************************************Traslados ****************************************************************************************************************/

  if (tecnico4.tecnico3 == "Ricardo Porcayo" && tecnico4.enviar3 == "No" && tecnico4.solicitud3 == "Traslados") {
    const plantilla = HtmlService.createTemplateFromFile('Traslados');
    plantilla.tecnico4 = tecnico4;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: "ricardo.porcayo@humanitas.edu.mx",
      subject: "Ticket #" + tecnico4.folio3,
      htmlBody: mensaje
    });
  }

  if (tecnico4.tecnico3 == "Angel Montes" && tecnico4.enviar3 == "No" && tecnico4.solicitud3 == "Traslados") {
    const plantilla = HtmlService.createTemplateFromFile('Traslados');
    plantilla.tecnico4 = tecnico4;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: "angel.montes@humanitas.edu.mx",
      subject: "Ticket #" + tecnico4.folio3,
      htmlBody: mensaje
    });
  }

  if (tecnico4.tecnico3 == "Gerardo Omaña" && tecnico4.enviar3 == "No" && tecnico4.solicitud3 == "Traslados") {
    const plantilla = HtmlService.createTemplateFromFile('Traslados');
    plantilla.tecnico4 = tecnico4;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: "gerardo.omana@humanitas.edu.mx",
      subject: "Ticket #" + tecnico4.folio3,
      htmlBody: mensaje
    });
  }

  if (tecnico4.tecnico3 == "Andrea Ayala" && tecnico4.enviar3 == "No" && tecnico4.solicitud3 == "Traslados") {
    const plantilla = HtmlService.createTemplateFromFile('Traslados');
    plantilla.tecnico4 = tecnico4;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: "andrea.ayala@humanitas.edu.mx",
      subject: "Ticket #" + tecnico4.folio3,
      htmlBody: mensaje
    });
  }

  if (tecnico4.tecnico3 == "Victor Barrera" && tecnico4.enviar3 == "No" && tecnico4.solicitud3 == "Traslados") {
    const plantilla = HtmlService.createTemplateFromFile('Traslados');
    plantilla.tecnico4 = tecnico4;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: "victor@humanitas.edu.mx",
      subject: "Ticket #" + tecnico4.folio3,
      htmlBody: mensaje
    });
  }

  //************************************************Fin de Mensaje de Traslados (Asignado)*************************************************************************************************************************/

  else { return; }
}

//Función para agregar Menú

function onOpen() {
  menu();
}
function menu() {
  var menu = SpreadsheetApp.getUi().createMenu("Enviar");
  menu.addItem("📤 Enviar Respuestas", "enviarCorreos");
  menu.addItem("✅ Enviar Ticket Asignado", "enviarAsignado");
  menu.addToUi();
}

//Función para Enviar Soluciones

function enviarCorreos() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  libro.setActiveSheet(libro.getSheetByName("Respuestas de formulario 1"));
  const hoja = SpreadsheetApp.getActiveSheet();
  const filas = hoja.getRange("A4000:Y5020").getValues();

  for (indiceFila in filas) {
    var user = crearUser(filas[indiceFila]);
    enviarCorreo2(user);

  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Correo enviado');
}

function crearUser(datosFila) {
  const user = {
    fecha2: datosFila[0],
    mail2: datosFila[1],
    campus2: datosFila[2],
    name2: datosFila[3],
    solicitud2: datosFila[4],
    descripcion2: datosFila[5],
    captura2: datosFila[6],
    campusOrigen2: datosFila[7],
    campusDestino2: datosFila[8],
    matricula2: datosFila[9],
    nameAlumno2: datosFila[10],
    ticket2: datosFila[11],
    folio2: datosFila[12],
    tecnico2: datosFila[13],
    estatus2: datosFila[14],
    atencion2: datosFila[15],
    prior2: datosFila[16],
    colaboracion2: datosFila[17],
    solucion2: datosFila[18],
    enviar2: datosFila[19],
    comentarios: datosFila[23],
    etiqueta2: datosFila[24],
  };
  return user;
}


function enviarCorreo2(user) {
  if (user.mail2 == "") return;

  /*****************************************************************Enviar en Espera **************************************************************************** */

  if (user.estatus2 == "En Espera" && user.enviar2 == "No") {
    const plantilla = HtmlService.createTemplateFromFile('en_espera');
    plantilla.user = user;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: user.mail2,
      subject: "Ticket #" + user.folio2 + " - " + user.estatus2,
      htmlBody: mensaje
    });
  }

  if (user.estatus2 == "Solucionado" && user.enviar2 == "No") {
    const plantilla = HtmlService.createTemplateFromFile('solucionado');
    plantilla.user = user;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: user.mail2,
      subject: "Ticket #" + user.folio2 + " - " + user.estatus2,
      htmlBody: mensaje
    });
  }

  if (user.estatus2 == "En Progreso" && user.enviar2 == "No" && user.solicitud2 != "Traslados") {
    const plantilla = HtmlService.createTemplateFromFile('progreso');
    plantilla.user = user;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: user.mail2,
      subject: "Ticket #" + user.folio2 + " - " + user.estatus2,
      htmlBody: mensaje
    });
  }

  if (user.estatus2 == "En Progreso" && user.enviar2 == "No" && user.solicitud2 == "Traslados") {
    const plantilla = HtmlService.createTemplateFromFile('progreso_traslado');
    plantilla.user = user;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: user.mail2,
      subject: "Ticket #" + user.folio2 + " - " + user.estatus2,
      htmlBody: mensaje
    });
  }

  if (user.estatus2 == "Cancelado por Tiempo" && user.enviar2 == "No") {
    const plantilla = HtmlService.createTemplateFromFile('cancelado_tiempo');
    plantilla.user = user;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: user.mail2,
      subject: "Ticket #" + user.folio2 + " - " + user.estatus2,
      htmlBody: mensaje
    });
  }

  if (user.estatus2 == "Cancelado" && user.enviar2 == "No") {
    const plantilla = HtmlService.createTemplateFromFile('cancelado');
    plantilla.user = user;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: user.mail2,
      subject: "Ticket #" + user.folio2 + " - " + user.estatus2,
      htmlBody: mensaje
    });
  }

  if (user.estatus2 == "Duplicado" && user.enviar2 == "No") {
    const plantilla = HtmlService.createTemplateFromFile('duplicado');
    plantilla.user = user;
    const mensaje = plantilla.evaluate().getContent();

    MailApp.sendEmail({
      name: "Centro de Soporte Humanitas",
      recipient: "soporte@humanitas.edu.mx",
      to: user.mail2,
      subject: "Ticket #" + user.folio2 + " - " + user.estatus2,
      htmlBody: mensaje
    });
  }

  else { return user; }
}

//Función para horas y fechas
function horaAsignado() {
  var incidencias = SpreadsheetApp.getActiveSpreadsheet();
  var hojaIncidencia = incidencias.getSheetByName("Respuestas de formulario 1");
  var activa = hojaIncidencia.getActiveCell();
  var dato = activa.getValue();
  var filaActiva = activa.getRow();
  var colActiva = activa.getColumn();

  if (filaActiva >= 2 && (colActiva == 14 || colActiva == 15) && incidencias.getActiveSheet().getName() == "Respuestas de formulario 1") {
    if (activa.offset(0, 7).getValue()) { return; }
    else { activa.offset(0, 7).setValue(new Date()); }
  }

  if (filaActiva >= 2 && (colActiva == 19) && incidencias.getActiveSheet().getName() == "Respuestas de formulario 1") {
    if (activa.offset(0, 4).getValue()) { return; }
    else { activa.offset(0, 4).setValue(new Date()); }
  }
}

function onEdit() {
  horaAsignado();
}