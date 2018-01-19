/Los nombres de las hojas de trabajo son en español (hojaEspañol), y lo demás será inglés (columnName, cellName, formulaAssignation, functionName, etc.).
//Hoja de trabajo Propuestas y Tesis sometidas (respuestas).
//Hoja de trabajo Asignación de Evaluadores (asignacionEvaluadores).

function onOpen(e) {
  //Al abrir el gSheet se genera el menú "SISTEMA", que permite asignar evaluador(es) y enviar correo a los evaluadores seleccionados.
  SpreadsheetApp.getUi()
  .createMenu('SISTEMA')
  .addItem('Asignar Evaluador', 'evaluatorAssignation')
  .addSeparator()
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Correo Manual')
              .addItem('Enviar a (1) evaluador', 'letterToEvaluators'))
  .addSeparator()
  .addItem('Enviar Respuesta Evaluación', 'manualDocumentEvaluation')
  .addToUi();
}

function evaluatorsNeeded() {
  //Se carga al inicio, y permite encontrar, de acuerdo al nivel y tipo de documento, la cantidad de evaluadores requerida.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var respuestas = ss.getSheetByName('Propuestas y Tesis sometidas');
  var educationLevel = respuestas.getRange(respuestas.getLastRow(), 2).getValue();
  var type = respuestas.getRange(respuestas.getLastRow(), 26).getValue();
  var cell = respuestas.getRange(respuestas.getLastRow(), 44).getCell(1,1);
  cell.setFormulaR1C1('=if(and(REGEXMATCH(R[0]C[-42],"Docto"), REGEXMATCH(R[0]C[-18],"final")),"3 evaluadores", if(and(REGEXMATCH(R[0]C[-42],"Docto"), REGEXMATCH(R[0]C[-18],"Proposal")),"2 evaluadores", if(and(REGEXMATCH(R[0]C[-42],"Maest"), REGEXMATCH(R[0]C[-18],"final")),"2 evaluadores", "1 evaluador")))');
  //var URL = respuestas.getRange(respuestas.getLastRow(), 32).getCell(1,1).getValue();
  //var urlCell = respuestas.getRange(respuestas.getLastRow(), 48).getCell(1,1);
}

function processBegin() {
  //Se carga al inicio, y permite encontrar, de acuerdo al nivel y tipo de documento, la cantidad de evaluadores requerida.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var respuestas = ss.getSheetByName('Propuestas y Tesis sometidas');
  var url = respuestas.getRange(respuestas.getLastRow(), 32).getValue();
  var shortUrl = UrlShortener.Url.insert({ longUrl: url });
  respuestas.getRange(respuestas.getLastRow(), 45).setValue(shortUrl.id);
  var dataRangeProceso = respuestas.getRange(respuestas.getLastRow(), 1, 1, respuestas.getLastColumn());
  var dataProceso = dataRangeProceso.getValues();
  var fechaSolicitud = dataProceso[0][0];
  var fechaCartaInicio = Utilities.formatDate(new Date(), "GMT", "dd - MM - yyyy");
  var programa = dataProceso[0][1];
  var nombreEstudiante = dataProceso[0][2];
  var apellidoEstudiante = dataProceso[0][3];
  var correoEstudiante = dataProceso[0][4];
  var director = dataProceso[0][9];
  var correoDirector = dataProceso[0][10];
  var titulo = dataProceso[0][20];
  var resumenTrabajo = dataProceso[0][21];
  var tipoTrabajo = dataProceso[0][25];
  var documento = dataProceso[0][28];
  var cuantosEvaluadores = dataProceso[0][43];
  var cartaEstudiante = ("ESTUDIANTE: " + nombreEstudiante + "-" + tipoTrabajo);
  var confirmacionIdEstudiante = DriveApp.getFileById('1rxH3W_WQ7jQ5pmmiDnScFyItuc0D0kV2z4miHUYV-xI').makeCopy(cartaEstudiante).getId();
  var confirmacionEstudiante = DocumentApp.openById(confirmacionIdEstudiante);
  var cuerpoConfirmacionEstudiante = confirmacionEstudiante.getActiveSection();
  cuerpoConfirmacionEstudiante.replaceText("%fechaCartaInicio%", fechaCartaInicio);
  cuerpoConfirmacionEstudiante.replaceText("%studentName%", nombreEstudiante + " " + apellidoEstudiante);
  cuerpoConfirmacionEstudiante.replaceText("%tipoTrabajo%", tipoTrabajo);
  cuerpoConfirmacionEstudiante.replaceText("%titulo%", titulo);
  cuerpoConfirmacionEstudiante.replaceText("%programa%", programa);
  cuerpoConfirmacionEstudiante.replaceText("%cuantosEvaluadores%", cuantosEvaluadores);
  cuerpoConfirmacionEstudiante.replaceText("%resumenTrabajo%", resumenTrabajo);
  confirmacionEstudiante.saveAndClose();
  var pdfConfirmacionEstudiante = DriveApp.createFile(confirmacionEstudiante.getAs("application/pdf"));
  var pdfConfirmacionId = pdfConfirmacionEstudiante.getId();
  var subjectEstudiante = "Inicio del proceso de evaluación de " + tipoTrabajo + ".";
  var bodyEstudiante = "Respetad@ " + nombreEstudiante +".\n\nLa coordinación del posgrado hace recepción de sus documentos para dar inicio al proceso de asignación de evaluadores de su " + tipoTrabajo + ": " + titulo + "."
  MailApp.sendEmail(correoEstudiante, subjectEstudiante , bodyEstudiante, {
    attachments: [pdfConfirmacionEstudiante.getAs(MimeType.PDF)]
  });
  var folderEvaluatorsRequest = DriveApp.getFolderById("1DIr8mXJCwmc_2prV1gIZzWcqyJdK4G7_");
  var fileEstudiante = DriveApp.getFileById(pdfConfirmacionEstudiante.getId());
  folderEvaluatorsRequest.addFile(fileEstudiante);
  var root = DriveApp.getRootFolder();
  root.removeFile(pdfConfirmacionEstudiante);
  DriveApp.getFileById(confirmacionIdEstudiante).setTrashed(true);
  var cartaDirector = ("DIRECTOR: " +  director + "-" + nombreEstudiante + "-" + programa);
  var confirmacionIdDirector = DriveApp.getFileById('1Mw9zzAu6YvO-gUsg-O4TsDJVbSyyOJEcLJFxCcOKl0w').makeCopy(cartaDirector).getId();
  var confirmacionDirector = DocumentApp.openById(confirmacionIdDirector);
  var cuerpoConfirmacionDirector = confirmacionDirector.getActiveSection();
  cuerpoConfirmacionDirector.replaceText("%fechaCartaInicio%", fechaCartaInicio);
  cuerpoConfirmacionDirector.replaceText("%director%", director);
  cuerpoConfirmacionDirector.replaceText("%tipoTrabajo%", tipoTrabajo);
  cuerpoConfirmacionDirector.replaceText("%titulo%", titulo);
  cuerpoConfirmacionDirector.replaceText("%studentName%", nombreEstudiante + " " + apellidoEstudiante);
  cuerpoConfirmacionDirector.replaceText("%programa%", programa);
  cuerpoConfirmacionDirector.replaceText("%cuantosEvaluadores%", cuantosEvaluadores);
  cuerpoConfirmacionDirector.replaceText("%resumenTrabajo%", resumenTrabajo);
  confirmacionDirector.saveAndClose();
  var pdfConfirmacionDirector = DriveApp.createFile(confirmacionDirector.getAs("application/pdf"));
  var pdfConfirmacionIdDirector = pdfConfirmacionDirector.getId();
  var subjectDirector = "Inicio del proceso de evaluación de " + tipoTrabajo + ".";
  var bodyDirector = "Respetad@ " + director +".\n\nLa coordinación del posgrado hace recepción de los documentos de " + nombreEstudiante + " " + apellidoEstudiante +" para dar inicio al proceso de asignación de evaluadores de su " + tipoTrabajo + ": " + titulo + "."
  MailApp.sendEmail(correoDirector, subjectDirector , bodyDirector, {
    attachments: [pdfConfirmacionDirector.getAs(MimeType.PDF)]
  });
  var folderEvaluatorsRequest = DriveApp.getFolderById("1DIr8mXJCwmc_2prV1gIZzWcqyJdK4G7_");
  var fileDirector = DriveApp.getFileById(pdfConfirmacionDirector.getId());
  folderEvaluatorsRequest.addFile(fileDirector);
  var root = DriveApp.getRootFolder();
  root.removeFile(pdfConfirmacionDirector);
  DriveApp.getFileById(confirmacionIdDirector).setTrashed(true);
} 

function evaluatorAssignation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var asignacionEvaluadores = ss.getSheetByName('Asignación de Evaluadores');
  asignacionEvaluadores.clear();
  var respuestas = ss.getSheetByName("Propuestas y Tesis sometidas");
  var evaluadores = ss.getSheetByName("Banco evaluadores");
  var querySheet = ("'Banco evaluadores'" + "!A2:S");
  var studentRow = respuestas.getActiveCell().getRow();
  respuestas.getRange(2, 1, 1, respuestas.getLastColumn()).copyValuesToRange(asignacionEvaluadores, 1, 45, 2, 2);
  respuestas.getRange(studentRow, 1, 1, respuestas.getLastColumn()).copyValuesToRange(asignacionEvaluadores, 1, 45, 3, 3);
  asignacionEvaluadores.deleteColumns(1, 1);
  asignacionEvaluadores.deleteColumns(5, 4);
  asignacionEvaluadores.deleteColumns(7, 9);
  asignacionEvaluadores.deleteColumns(8, 3);
  asignacionEvaluadores.deleteColumns(10, 27);
  var alerta = asignacionEvaluadores.getRange("C1");
  var alerta1 = asignacionEvaluadores.getRange("E1");
  alerta.setValue("La preselección de evaluadores consiste en el proceso manual de seleccionar una casilla que corresponda al estudiante (al cual es necesario preseleccionar evaluadores) en la hoja 'Respuestas Propuestas y Tesis'. Los investigadores son preseleccionados si alguna de las palabras claves del estudiante hacen match con alguna de las palabras claves del correspondiente investigador.  En las columnas 'D5:G' se obtiene la salida generada por el sistema.");
  alerta.setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true).setBackground("Coral");
  asignacionEvaluadores.setColumnWidth(3, 240);
  var palabras = asignacionEvaluadores.getRange(3, 8, 1, 1).getValues();
  var listaPalabras = palabras.toString().split(',');
  var longitud = listaPalabras.length;
  asignacionEvaluadores.getRange(4, 1, longitud, asignacionEvaluadores.getLastColumn());
  asignacionEvaluadores.getRange("E5:AZ").setBackground("WHITE").setFontColor("WHITE");
  for (var i = 0; i<longitud; i++){
    var keyWords = listaPalabras[i];
    var keyWordLimpia = asignacionEvaluadores.getRange(4+i, 1).setValue("'" + keyWords.trim() + "'");
    asignacionEvaluadores.getRange(5, 6+i).setFormula('=QUERY(' + querySheet + ', ' + '"SELECT S WHERE O CONTAINS ' + "'" + keyWordLimpia.getValue() + ' OR P CONTAINS ' + "'" + keyWordLimpia.getValue() + ' OR Q CONTAINS ' + "'" + keyWordLimpia.getValue() + ' ORDER BY S")');
    if (asignacionEvaluadores.getRange(5, 6+i).getValue() == "#N/A")
    {
      asignacionEvaluadores.getRange(5, 6+i).setValue("");
    }
  }
  var temp, columna = '';
  temp = (asignacionEvaluadores.getLastColumn() - 1) % 26;
  columna = String.fromCharCode(temp + 65) + columna;
  var fin = columna + asignacionEvaluadores.getLastRow();
  asignacionEvaluadores.getRange(20, 1, (asignacionEvaluadores.getLastRow() + 19)).setBackground("WHITE").setFontColor("WHITE");
  asignacionEvaluadores.getRange("A20").setFormula('=unique(transpose(split(ArrayFormula(concatenate(F5:' + fin + '& " "))," ")))');
  var cantidadEvaluadores = asignacionEvaluadores.getLastRow() - 19;
  var rangoEvaluadores = asignacionEvaluadores.getRange(20, 1, cantidadEvaluadores);
  var evaluadoresFinales = "A20:A" + cantidadEvaluadores;
  Logger.log(evaluadoresFinales);
  asignacionEvaluadores.getRange(evaluadoresFinales).copyTo(asignacionEvaluadores.getRange("D5"), {contentsOnly:true});
  var cantidad = asignacionEvaluadores.getRange("D1").setFormula("=COUNTA(D5:D)").setBackground("WHITE").setFontColor("WHITE");
  asignacionEvaluadores.getRange(evaluadoresFinales).clear();
  asignacionEvaluadores.getRange("D4:G4").setBackground("LIME");
  asignacionEvaluadores.getRange("D4").setValue("eMail Evaluador");
  asignacionEvaluadores.getRange("E4").setValue("Título");
  asignacionEvaluadores.getRange("F4").setValue("Nombre");
  asignacionEvaluadores.getRange("G4").setValue("Apellido");
  asignacionEvaluadores.getRange("E5:AZ").clear();
  asignacionEvaluadores.getRange("B4").setValue("Existen " + asignacionEvaluadores.getRange("D1").getValue() + " posibles evaluadores").setBackground("Orange");
  var cantidad = asignacionEvaluadores.getRange("D1").getValue();
  for (var i = 0; i< cantidad; i++){
    asignacionEvaluadores.getRange(5+i, 5).setFormula("=INDEX('Banco evaluadores'!L2:L,MATCH(D" + (i+5) + ",'Banco evaluadores'!S2:S,0))");
    asignacionEvaluadores.getRange(5+i, 6).setFormula("=INDEX('Banco evaluadores'!C2:C,MATCH(D" + (i+5) + ",'Banco evaluadores'!S2:S,0))");
    asignacionEvaluadores.getRange(5+i, 7).setFormula("=INDEX('Banco evaluadores'!D2:D,MATCH(D" + (i+5) + ",'Banco evaluadores'!S2:S,0))");
  }
  asignacionEvaluadores.getRange("D:D").setWrap(true);
  asignacionEvaluadores.getRange(3, 11, 1, 50).setBackground("WHITE").setFontColor("WHITE");
  respuestas.getRange(studentRow, 1, 1, respuestas.getLastColumn()).copyValuesToRange(asignacionEvaluadores, 11, 60, 3, 3);
  alerta1.setValue("El envío a 1 evaluador consiste en el proceso manual de seleccionar una casilla que corresponda al correo electrónico (D5:D) del evaluador. Se envía un mensaje cuyo cuerpo se puede editar (NO RECOMENDADO). Se adjunta el PDF que corresponde al documento del estudiante. ");
  alerta1.setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true).setBackground("PINK");
  asignacionEvaluadores.setColumnWidth(5, 300);
}

//ESTA ES LA FUNCION QUE TRABAJA DESDE EVALUADORES, seleccionando una casilla donde esté el correo electrónico, y envía la solicitud al seleccionado
function letterToEvaluators() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var evaluadores = ss.getSheetByName('Asignación de Evaluadores');
  var hojaCorreos = ss.getSheetByName('Envio de correos automáticos eva');
  var hojaMensaje = ss.getSheetByName("Correo solicitud de evaluación");
  var dataRangeProceso = evaluadores.getRange(3, 11, 1, 60);
  var dataProceso = dataRangeProceso.getValues();
  var fechaSolicitud = dataProceso[0][0];
  var fechaCartaInicio = Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy");
  var programa = dataProceso[0][1];
  var nombreEstudiante = dataProceso[0][2];
  var apellidoEstudiante = dataProceso[0][3];
  var correoEstudiante = dataProceso[0][4];
  var director = dataProceso[0][9];
  var correoDirector = dataProceso[0][10];
  var titulo = dataProceso[0][20];
  var resumenTrabajo = dataProceso[0][21];
  var keyWords = dataProceso[0][24];
  var tipoTrabajo = dataProceso[0][25];
  var documento = dataProceso[0][31];
  var cuantosEvaluadores = dataProceso[0][40];
  var URL = dataProceso[0][44];
  var aceptacion = 'https://docs.google.com/forms/d/e/1FAIpQLSdsWkc6QkKRPfIXFA1vMXPgRCVf2SfZiToKVdA9F0lov92nIQ/viewform?usp=pp_url&entry.406285137=' + URL + '&entry.1833441442=' + programa +'&entry.109767644=' + nombreEstudiante + ' ' + apellidoEstudiante +'&entry.1431544173=' + director + '&entry.1200499305=' + titulo + '&entry.2117079111=' + tipoTrabajo + '&entry.1942458416&entry.626264777&entry.593390219&entry.1650230059&entry.2114359596&entry.169738932&entry.142326488&entry.224835711&entry.38887212&entry.1827453043&entry.1227381213&entry.216539192&entry.2018541163&entry.1297196829&entry.1290095238&entry.1105524184&entry.160383615&entry.494482189&entry.154747419&entry.764672352&entry.1697461536&entry.974536642&entry.879997161&entry.206682334';
  var acepto1 = UrlShortener.Url.insert({ longUrl: aceptacion});
  var acepto = acepto1.id;
  var noAceptacion = 'https://docs.google.com/a/correounivalle.edu.co/forms/d/e/1FAIpQLSefF5FDBq7nv3vfAcs6J5dHhmSYSGOwQ6sgPvxMIegn3UIiKQ/viewform';
  var noAcepto1 = UrlShortener.Url.insert({ longUrl: noAceptacion});
  var noAcepto = noAcepto1.id;
  var cartaEvaluador = (nombreEstudiante + " " + apellidoEstudiante + " " + tipoTrabajo + " " + programa);
  var confirmacionId = DriveApp.getFileById('1MwsqqmCFKL_x4XANsIjhLHDA-bvrJwWaR82ON65F88k').makeCopy(cartaEvaluador).getId();
  var confirmacion = DocumentApp.openById(confirmacionId);
  var filaEvaluadores = evaluadores.getActiveCell().getRow();
  Logger.log(filaEvaluadores);
  var dataRangeProcesoCorreoEvaluador = evaluadores.getRange(filaEvaluadores, 4, 1, 4);
  var dataProcesoEvaluador = dataRangeProcesoCorreoEvaluador.getValues();
  var nombreEvaluador = (dataProcesoEvaluador[0][1] + " " + dataProcesoEvaluador[0][2] + " " + dataProcesoEvaluador[0][3]);
  var evaluatorsEmails = dataProcesoEvaluador[0][0];
  var cuerpoConfirmacion = confirmacion.getActiveSection();
  cuerpoConfirmacion.replaceText("%fechaCartaInicio%", fechaCartaInicio);
  cuerpoConfirmacion.replaceText("%nombreEvaluador%", nombreEvaluador);
  cuerpoConfirmacion.replaceText("%titulo%", titulo);
  cuerpoConfirmacion.replaceText("%programa%", programa);
  cuerpoConfirmacion.replaceText("%tipoTrabajo%", tipoTrabajo);
  cuerpoConfirmacion.replaceText("%resumenTrabajo%", resumenTrabajo);
  cuerpoConfirmacion.replaceText("%keyWords%", keyWords);
  cuerpoConfirmacion.replaceText("%studentName%", nombreEstudiante + " " + apellidoEstudiante);
  cuerpoConfirmacion.replaceText("%acepto%", acepto);
  cuerpoConfirmacion.replaceText("%noAcepto%", noAcepto);
  confirmacion.saveAndClose();
  var pdfConfirmacion = DriveApp.createFile(confirmacion.getAs("application/pdf"));
  var pdfConfirmacionId = pdfConfirmacion.getId();
  pdfConfirmacion.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var payment;
  var subject = hojaMensaje.getRange("B2").getValue();
  var body1 = hojaMensaje.getRange("B3").getValue();
  var body2 = hojaMensaje.getRange("B4").getValue();
  var body3 = hojaMensaje.getRange("B5").getValue();
  var payment;
  if (cuantosEvaluadores == "3 evaluadores")
  {
    payment = hojaMensaje.getRange("E9").getValue()
  } else if (cuantosEvaluadores == "2 evaluadores")
  {
    payment = hojaMensaje.getRange("E10").getValue()
  } else {
    payment = hojaMensaje.getRange("E11").getValue()
  }
  var body4 = hojaMensaje.getRange("B11").getValue() + payment;
  var body5 = hojaMensaje.getRange("B8").getValue()
  var firma = hojaMensaje.getRange("B9").getValue();
  var body = body1 + "\n" + body2 + "\n" + body3 + "\n" + body4 + "\n" + body5 + "\n"+ firma;
  MailApp.sendEmail(evaluatorsEmails, subject, body, {
    attachments: [pdfConfirmacion.getAs(MimeType.PDF)]
  });
  var folderEvaluatorsRequest = DriveApp.getFolderById("0B08tcqjuUlVbRGhZYzlXZXRMMnM");
  var file = DriveApp.getFileById(pdfConfirmacion.getId());
  folderEvaluatorsRequest.addFile(file);
  var root = DriveApp.getRootFolder();
  root.removeFile(pdfConfirmacion);
  DriveApp.getFileById(confirmacionId).setTrashed(true);
}

//ESTO ES LO QUE DEBE FUNCIONAR
function documentEvaluation(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var respuestas = ss.getSheetByName('Propuestas y Tesis sometidas');
  var respuestaEvaluadores = ss.getSheetByName('Respuestas de Evaluación');
  var dataRangeProceso = respuestaEvaluadores.getRange(respuestaEvaluadores.getLastRow(), 1, 1, 52);
  var dataProceso = dataRangeProceso.getValues();
  //PARA TODO
  var fecha= dataProceso[0][0];
  var fechaCartaInicio = Utilities.formatDate(new Date(), "GMT", "dd - MM - yyyy");
  var emailEvaluador= dataProceso[0][1];
  var confidencialidad= dataProceso[0][2];
  var guiaEvaluacion= dataProceso[0][3];
  var documentShortURL= dataProceso[0][4];
  var url = UrlShortener.Url.get(documentShortURL);
  var URL = url.longUrl;
  Logger.log(URL);
  var programa = dataProceso[0][5];
  var nombreEstudiante = dataProceso[0][6];
  var nombreDirector = dataProceso[0][7];
  var tituloDocumento = dataProceso[0][8];
  var tipoDocumento = dataProceso[0][9];
  //PARA PROPUESTA
  var formulacionPropuesta = dataProceso[0][10];
  var justificacionFormulacionPropuesta = dataProceso[0][11];
  var objetivosPropuesta = dataProceso[0][12];
  var justificacionObjetivosPropuesta = dataProceso[0][13];
  var marcoConceptualPropuesta = dataProceso[0][14];
  var justificacionMarcoConceptualPropuesta = dataProceso[0][15];
  var metodologiaPropuesta = dataProceso[0][16];
  var justificacionMetodologiaPropuesta = dataProceso[0][17];
  var metodosProcesamientoPropuesta = dataProceso[0][18];
  var justificacionMetodosProcesamientoPropuesta = dataProceso[0][19];
  var impactoResultadosPropuesta = dataProceso[0][20];
  var justificacionImpactoResultadosPropuesta = dataProceso[0][21];
  var cronogramaPropuesta = dataProceso[0][22];
  var justificacionCronogramaPropuesta = dataProceso[0][23];
  var presupuestoPropuesta = dataProceso[0][24];
  var justificacionPresupuestoPropuesta = dataProceso[0][25];
  var recomendacionAceptacionPropuesta = dataProceso[0][26];
  var revisarDocumentoFinalPropuesta = dataProceso[0][27];
  var reformulacionPropuesta = dataProceso[0][28];
  var nombreEvaluadorPropuesta = dataProceso[0][29];
  var institucionPropuesta = dataProceso[0][30];
  //PARA FINAL
  var formulacionFinal = dataProceso[0][31];
  var justificacionFormulacionFinal = dataProceso[0][32];
  var objetivosFinal = dataProceso[0][33];
  var justificacionObjetivosFinal = dataProceso[0][34];
  var marcoConceptualFinal = dataProceso[0][35];
  var justificacionMarcoConceptualFinal = dataProceso[0][36];
  var metodologiaFinal = dataProceso[0][37];
  var justificacionMetodologiaFinal = dataProceso[0][38];
  var metodosProcesamientoFinal = dataProceso[0][39];
  var justificacionMetodosProcesamientoFinal = dataProceso[0][40];
  var alcancesInnovacionFinal = dataProceso[0][41];
  var calidadTrabajoFinal = dataProceso[0][42];
  var justificacionCalidadTrabajoFinal = dataProceso[0][43];
  var comentariosGeneralesFinal = dataProceso[0][44];
  var recomiendaSustentacionFinal = dataProceso[0][45];
  var aceptaSerPresidenteFinal = dataProceso[0][46];
  var nombreEvaluadorFinal = dataProceso[0][47];
  var institucionFinal = dataProceso[0][48];
  var comentariosFinal = dataProceso[0][50];
  if (recomendacionAceptacionPropuesta == "No/I reject"){
    var subjectProblema = "No recomiendan " + tipoDocumento + " del estudiante " + nombreEstudiante + " dirigido por " + nombreDirector + ".";
    var bodyProblema = ("El evaluador " + nombreEvaluadorPropuesta + " tiene las siguientes observaciones "+ reformulacionPropuesta + "." ); 
    MailApp.sendEmail("evaluacion.posgrado.pisa@correounivalle.edu.co", subjectProblema , bodyProblema);
  } else if (recomiendaSustentacionFinal == "NO") {
    var subjectProblema = "No recomiendan " + tipoDocumento + " del estudiante " + nombreEstudiante + " dirigido por " + nombreDirector + ".";
    var bodyProblema = ("El evaluador " + nombreEvaluadorFinal + " considera que no es viable la sustentación pública de la tesis."); 
    MailApp.sendEmail("evaluacion.posgrado.pisa@correounivalle.edu.co", subjectProblema , bodyProblema);
  } else {  
    var cell = respuestaEvaluadores.getRange(respuestaEvaluadores.getLastRow(),52);
    cell.setFormulaR1C1('=if(REGEXMATCH(R[0]C[-42],"Final"),"Final","Propuesta")');
    Logger.log(cell.getValue());
    var querySheet = ("'Propuestas y Tesis sometidas'" + "!C3:AS");
    respuestaEvaluadores.getRange(respuestaEvaluadores.getLastRow(), 53).setFormula('=QUERY(' + querySheet + ', ' + '"SELECT C, D, E, J, K, M, N, O, P WHERE AF CONTAINS ' + "'" + URL + "'" + '")');
    var cartaEstudiante = ("ESTUDIANTE: " + nombreEstudiante + "-" + tipoDocumento);
    if(cell.getValue() == "Final"){
      var idEstudiante = '1WgOsVcz3xG8YafcpsR-UjkvEh2csTSRtuUFdXJhUiBI';
      //var confirmacionIdEstudiante = DriveApp.getFileById('1WgOsVcz3xG8YafcpsR-UjkvEh2csTSRtuUFdXJhUiBI').makeCopy(cartaEstudiante).getId();
    } else {
      var idEstudiante = '1Q75_hgDgdNVAtgBUqjlhIqL41z3PrlrBqLVK5iWv8sc';
      //var confirmacionIdEstudiante = DriveApp.getFileById('1Q75_hgDgdNVAtgBUqjlhIqL41z3PrlrBqLVK5iWv8sc').makeCopy(cartaEstudiante).getId();
    }
    var confirmacionIdEstudiante = DriveApp.getFileById(idEstudiante).makeCopy(cartaEstudiante).getId();
    var confirmacionEstudiante = DocumentApp.openById(confirmacionIdEstudiante);
    var cuerpoConfirmacionEstudiante = confirmacionEstudiante.getActiveSection();
    //Para TODO
    cuerpoConfirmacionEstudiante.replaceText("%fechaCartaInicio%", fechaCartaInicio);
    cuerpoConfirmacionEstudiante.replaceText("%nombreEstudiante%", nombreEstudiante);
    cuerpoConfirmacionEstudiante.replaceText("%tituloDocumento%", tituloDocumento);
    cuerpoConfirmacionEstudiante.replaceText("%programa%", programa);
    cuerpoConfirmacionEstudiante.replaceText("%nombreEvaluadorPropuesta%", nombreEvaluadorPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%nombreEvaluadorFinal%", nombreEvaluadorFinal);
    cuerpoConfirmacionEstudiante.replaceText("%confidencialidad%", confidencialidad);
    //Para PROPUESTA
    cuerpoConfirmacionEstudiante.replaceText("%formulacionPropuesta%", formulacionPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionFormulacionPropuesta%", justificacionFormulacionPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%objetivosPropuesta%", objetivosPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionObjetivosPropuesta%", justificacionObjetivosPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%marcoConceptualPropuesta%", marcoConceptualPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMarcoConceptualPropuesta%", justificacionMarcoConceptualPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%metodologiaPropuesta%", metodologiaPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMetodologiaPropuesta%", justificacionMetodologiaPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%metodosProcesamientoPropuesta%", metodosProcesamientoPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMetodosProcesamientoPropuesta%", justificacionMetodosProcesamientoPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%impactoResultadosPropuesta%", impactoResultadosPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionImpactoResultadosPropuesta%", justificacionImpactoResultadosPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%cronogramaPropuesta%", cronogramaPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionCronogramaPropuesta%", justificacionCronogramaPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%presupuestoPropuesta%", presupuestoPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionPresupuestoPropuesta%", justificacionPresupuestoPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%recomendacionAceptacionPropuesta%", recomendacionAceptacionPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%revisarDocumentoFinalPropuesta%", revisarDocumentoFinalPropuesta);
    //para FINAL
    cuerpoConfirmacionEstudiante.replaceText("%formulacionFinal%", formulacionFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionFormulacionFinal%", justificacionFormulacionFinal);
    cuerpoConfirmacionEstudiante.replaceText("%objetivosFinal%", objetivosFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionObjetivosFinal%", justificacionObjetivosFinal);
    cuerpoConfirmacionEstudiante.replaceText("%marcoConceptualFinal%", marcoConceptualFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMarcoConceptualFinal%", justificacionMarcoConceptualFinal);
    cuerpoConfirmacionEstudiante.replaceText("%metodologiaFinal%", metodologiaFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMetodologiaFinal%", justificacionMetodologiaFinal);
    cuerpoConfirmacionEstudiante.replaceText("%metodosProcesamientoFinal%", metodosProcesamientoFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMetodosProcesamientoFinal%", justificacionMetodosProcesamientoFinal);
    cuerpoConfirmacionEstudiante.replaceText("%alcancesInnovacionFinal%", alcancesInnovacionFinal);
    cuerpoConfirmacionEstudiante.replaceText("%calidadTrabajoFinal%", calidadTrabajoFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionCalidadTrabajoFinal%", justificacionCalidadTrabajoFinal);
    cuerpoConfirmacionEstudiante.replaceText("%comentariosGeneralesFinal%", comentariosGeneralesFinal);
    cuerpoConfirmacionEstudiante.replaceText("%recomiendaSustentacionFinal%", recomiendaSustentacionFinal);
    cuerpoConfirmacionEstudiante.replaceText("%aceptaSerPresidenteFinal%", aceptaSerPresidenteFinal);
    confirmacionEstudiante.saveAndClose();
    var pdfConfirmacionEstudiante = DriveApp.createFile(confirmacionEstudiante.getAs("application/pdf"));
    var pdfConfirmacionId = pdfConfirmacionEstudiante.getId();
    var subjectEstudiante = "Respuesta del evaluador de " + tipoDocumento + ".";
    var bodyEstudiante = ("Respetad@s " + nombreEstudiante + " y " + nombreDirector + ".\n\nLa coordinación del posgrado hace envío de la evaluación de su " + tipoDocumento + ".\nFelicitaciones!"); 
    Logger.log(respuestaEvaluadores.getRange(respuestaEvaluadores.getLastRow(), 55).getValue());
    var correos = ("evaluacion.posgrado.pisa@correounivalle.edu.co," + respuestaEvaluadores.getRange(respuestaEvaluadores.getLastRow(), 55).getValue()+"," + respuestaEvaluadores.getRange(respuestaEvaluadores.getLastRow(), 57).getValue());
    MailApp.sendEmail(correos, subjectEstudiante , bodyEstudiante, {
      attachments: [pdfConfirmacionEstudiante.getAs(MimeType.PDF)],
    });
    Logger.log(correos);
    var folderEvaluatorsRequest = DriveApp.getFolderById("1DIr8mXJCwmc_2prV1gIZzWcqyJdK4G7_");
    var fileEstudiante = DriveApp.getFileById(pdfConfirmacionEstudiante.getId());
    folderEvaluatorsRequest.addFile(fileEstudiante);
    var root = DriveApp.getRootFolder();
    root.removeFile(pdfConfirmacionEstudiante);
    DriveApp.getFileById(confirmacionIdEstudiante).setTrashed(true);
  }
}

function letterToEvaluators() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var evaluadores = ss.getSheetByName('Asignación de Evaluadores');
  var hojaCorreos = ss.getSheetByName('Envio de correos automáticos eva');
  var hojaMensaje = ss.getSheetByName("Correo solicitud de evaluación");
  var dataRangeProceso = evaluadores.getRange(3, 11, 1, 60);
  var dataProceso = dataRangeProceso.getValues();
  var fechaSolicitud = dataProceso[0][0];
  var fechaCartaInicio = Utilities.formatDate(new Date(), "GMT", "dd-MM-yyyy");
  var programa = dataProceso[0][1];
  var nombreEstudiante = dataProceso[0][2];
  var apellidoEstudiante = dataProceso[0][3];
  var correoEstudiante = dataProceso[0][4];
  var director = dataProceso[0][9];
  var correoDirector = dataProceso[0][10];
  var titulo = dataProceso[0][20];
  var resumenTrabajo = dataProceso[0][21];
  var keyWords = dataProceso[0][24];
  var tipoTrabajo = dataProceso[0][25];
  var documento = dataProceso[0][31];
  var cuantosEvaluadores = dataProceso[0][40];
  var URL = dataProceso[0][44];
  var aceptacion = 'https://docs.google.com/forms/d/e/1FAIpQLSdsWkc6QkKRPfIXFA1vMXPgRCVf2SfZiToKVdA9F0lov92nIQ/viewform?usp=pp_url&entry.406285137=' + URL + '&entry.1833441442=' + programa +'&entry.109767644=' + nombreEstudiante + ' ' + apellidoEstudiante +'&entry.1431544173=' + director + '&entry.1200499305=' + titulo + '&entry.2117079111=' + tipoTrabajo + '&entry.1942458416&entry.626264777&entry.593390219&entry.1650230059&entry.2114359596&entry.169738932&entry.142326488&entry.224835711&entry.38887212&entry.1827453043&entry.1227381213&entry.216539192&entry.2018541163&entry.1297196829&entry.1290095238&entry.1105524184&entry.160383615&entry.494482189&entry.154747419&entry.764672352&entry.1697461536&entry.974536642&entry.879997161&entry.206682334';
  var acepto1 = UrlShortener.Url.insert({ longUrl: aceptacion});
  var acepto = acepto1.id;
  var noAceptacion = 'https://docs.google.com/a/correounivalle.edu.co/forms/d/e/1FAIpQLSefF5FDBq7nv3vfAcs6J5dHhmSYSGOwQ6sgPvxMIegn3UIiKQ/viewform';
  var noAcepto1 = UrlShortener.Url.insert({ longUrl: noAceptacion});
  var noAcepto = noAcepto1.id;
  var cartaEvaluador = (nombreEstudiante + " " + apellidoEstudiante + " " + tipoTrabajo + " " + programa);
  var confirmacionId = DriveApp.getFileById('1MwsqqmCFKL_x4XANsIjhLHDA-bvrJwWaR82ON65F88k').makeCopy(cartaEvaluador).getId();
  var confirmacion = DocumentApp.openById(confirmacionId);
  var filaEvaluadores = evaluadores.getActiveCell().getRow();
  Logger.log(filaEvaluadores);
  var dataRangeProcesoCorreoEvaluador = evaluadores.getRange(filaEvaluadores, 4, 1, 4);
  var dataProcesoEvaluador = dataRangeProcesoCorreoEvaluador.getValues();
  var nombreEvaluador = (dataProcesoEvaluador[0][1] + " " + dataProcesoEvaluador[0][2] + " " + dataProcesoEvaluador[0][3]);
  var evaluatorsEmails = dataProcesoEvaluador[0][0];
  var cuerpoConfirmacion = confirmacion.getActiveSection();
  cuerpoConfirmacion.replaceText("%fechaCartaInicio%", fechaCartaInicio);
  cuerpoConfirmacion.replaceText("%nombreEvaluador%", nombreEvaluador);
  cuerpoConfirmacion.replaceText("%titulo%", titulo);
  cuerpoConfirmacion.replaceText("%programa%", programa);
  cuerpoConfirmacion.replaceText("%tipoTrabajo%", tipoTrabajo);
  cuerpoConfirmacion.replaceText("%resumenTrabajo%", resumenTrabajo);
  cuerpoConfirmacion.replaceText("%keyWords%", keyWords);
  cuerpoConfirmacion.replaceText("%studentName%", nombreEstudiante + " " + apellidoEstudiante);
  cuerpoConfirmacion.replaceText("%acepto%", acepto);
  cuerpoConfirmacion.replaceText("%noAcepto%", noAcepto);
  confirmacion.saveAndClose();
  var pdfConfirmacion = DriveApp.createFile(confirmacion.getAs("application/pdf"));
  var pdfConfirmacionId = pdfConfirmacion.getId();
  pdfConfirmacion.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var payment;
  var subject = hojaMensaje.getRange("B2").getValue();
  var body1 = hojaMensaje.getRange("B3").getValue();
  var body2 = hojaMensaje.getRange("B4").getValue();
  var body3 = hojaMensaje.getRange("B5").getValue();
  var payment;
  if (cuantosEvaluadores == "3 evaluadores")
  {
    payment = hojaMensaje.getRange("E9").getValue()
  } else if (cuantosEvaluadores == "2 evaluadores")
  {
    payment = hojaMensaje.getRange("E10").getValue()
  } else {
    payment = hojaMensaje.getRange("E11").getValue()
  }
  var body4 = hojaMensaje.getRange("B11").getValue() + payment;
  var body5 = hojaMensaje.getRange("B8").getValue()
  var firma = hojaMensaje.getRange("B9").getValue();
  var body = body1 + "\n" + body2 + "\n" + body3 + "\n" + body4 + "\n" + body5 + "\n"+ firma;
  MailApp.sendEmail(evaluatorsEmails, subject, body, {
    attachments: [pdfConfirmacion.getAs(MimeType.PDF)]
  });
  var folderEvaluatorsRequest = DriveApp.getFolderById("0B08tcqjuUlVbRGhZYzlXZXRMMnM");
  var file = DriveApp.getFileById(pdfConfirmacion.getId());
  folderEvaluatorsRequest.addFile(file);
  var root = DriveApp.getRootFolder();
  root.removeFile(pdfConfirmacion);
  DriveApp.getFileById(confirmacionId).setTrashed(true);
}

//ESTO ES LO QUE DEBE FUNCIONAR
function manualDocumentEvaluation(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var respuestas = ss.getSheetByName('Propuestas y Tesis sometidas');
  var respuestaEvaluadores = ss.getSheetByName('Respuestas de Evaluación');
  var filaEvaluadores = respuestaEvaluadores.getActiveCell().getRow();
  Logger.log (filaEvaluadores);
  var dataRangeProceso = respuestaEvaluadores.getRange(filaEvaluadores, 1, 1, 52);
  var dataProceso = dataRangeProceso.getValues();
  //PARA TODO
  var fecha= dataProceso[0][0];
  var fechaCartaInicio = Utilities.formatDate(new Date(), "GMT", "dd - MM - yyyy");
  var emailEvaluador= dataProceso[0][1];
  var confidencialidad= dataProceso[0][2];
  var guiaEvaluacion= dataProceso[0][3];
  var documentShortURL= dataProceso[0][4];
  var url = UrlShortener.Url.get(documentShortURL);
  var URL = url.longUrl;
  //Logger.log(URL);
  var programa = dataProceso[0][5];
  var nombreEstudiante = dataProceso[0][6];
  var nombreDirector = dataProceso[0][7];
  var tituloDocumento = dataProceso[0][8];
  var tipoDocumento = dataProceso[0][9];
  //PARA PROPUESTA
  var formulacionPropuesta = dataProceso[0][10];
  var justificacionFormulacionPropuesta = dataProceso[0][11];
  var objetivosPropuesta = dataProceso[0][12];
  var justificacionObjetivosPropuesta = dataProceso[0][13];
  var marcoConceptualPropuesta = dataProceso[0][14];
  var justificacionMarcoConceptualPropuesta = dataProceso[0][15];
  var metodologiaPropuesta = dataProceso[0][16];
  var justificacionMetodologiaPropuesta = dataProceso[0][17];
  var metodosProcesamientoPropuesta = dataProceso[0][18];
  var justificacionMetodosProcesamientoPropuesta = dataProceso[0][19];
  var impactoResultadosPropuesta = dataProceso[0][20];
  var justificacionImpactoResultadosPropuesta = dataProceso[0][21];
  var cronogramaPropuesta = dataProceso[0][22];
  var justificacionCronogramaPropuesta = dataProceso[0][23];
  var presupuestoPropuesta = dataProceso[0][24];
  var justificacionPresupuestoPropuesta = dataProceso[0][25];
  var recomendacionAceptacionPropuesta = dataProceso[0][26];
  var revisarDocumentoFinalPropuesta = dataProceso[0][27];
  var reformulacionPropuesta = dataProceso[0][28];
  var nombreEvaluadorPropuesta = dataProceso[0][29];
  var institucionPropuesta = dataProceso[0][30];
  //PARA FINAL
  var formulacionFinal = dataProceso[0][31];
  var justificacionFormulacionFinal = dataProceso[0][32];
  var objetivosFinal = dataProceso[0][33];
  var justificacionObjetivosFinal = dataProceso[0][34];
  var marcoConceptualFinal = dataProceso[0][35];
  var justificacionMarcoConceptualFinal = dataProceso[0][36];
  var metodologiaFinal = dataProceso[0][37];
  var justificacionMetodologiaFinal = dataProceso[0][38];
  var metodosProcesamientoFinal = dataProceso[0][39];
  var justificacionMetodosProcesamientoFinal = dataProceso[0][40];
  var alcancesInnovacionFinal = dataProceso[0][41];
  var calidadTrabajoFinal = dataProceso[0][42];
  var justificacionCalidadTrabajoFinal = dataProceso[0][43];
  var comentariosGeneralesFinal = dataProceso[0][44];
  var recomiendaSustentacionFinal = dataProceso[0][45];
  var aceptaSerPresidenteFinal = dataProceso[0][46];
  var nombreEvaluadorFinal = dataProceso[0][47];
  var institucionFinal = dataProceso[0][48];
  var comentariosFinal = dataProceso[0][50];
  if (recomendacionAceptacionPropuesta == "No/I reject"){
    var subjectProblema = "No recomiendan " + tipoDocumento + " del estudiante " + nombreEstudiante + " dirigido por " + nombreDirector + ".";
    var bodyProblema = ("El evaluador " + nombreEvaluadorPropuesta + " tiene las siguientes observaciones "+ reformulacionPropuesta + "." ); 
    MailApp.sendEmail("evaluacion.posgrado.pisa@correounivalle.edu.co", subjectProblema , bodyProblema);
  } else if (recomiendaSustentacionFinal == "NO") {
    var subjectProblema = "No recomiendan " + tipoDocumento + " del estudiante " + nombreEstudiante + " dirigido por " + nombreDirector + ".";
    var bodyProblema = ("El evaluador " + nombreEvaluadorFinal + " considera que no es viable la sustentación pública de la tesis."); 
    MailApp.sendEmail("evaluacion.posgrado.pisa@correounivalle.edu.co", subjectProblema , bodyProblema);
  } else {  
    var cell = respuestaEvaluadores.getRange(filaEvaluadores,52);
    cell.setFormulaR1C1('=if(REGEXMATCH(R[0]C[-42],"Final"),"Final","Propuesta")');
    Logger.log(cell.getValue());
    var querySheet = ("'Propuestas y Tesis sometidas'" + "!C3:AS");
    respuestaEvaluadores.getRange(filaEvaluadores, 53).setFormula('=QUERY(' + querySheet + ', ' + '"SELECT C, D, E, J, K, M, N, O, P WHERE AF CONTAINS ' + "'" + URL + "'" + '")');
    var cartaEstudiante = ("ESTUDIANTE: " + nombreEstudiante + "-" + tipoDocumento);
    if(cell.getValue() == "Final"){
      var idEstudiante = '1WgOsVcz3xG8YafcpsR-UjkvEh2csTSRtuUFdXJhUiBI';
      //var confirmacionIdEstudiante = DriveApp.getFileById('1WgOsVcz3xG8YafcpsR-UjkvEh2csTSRtuUFdXJhUiBI').makeCopy(cartaEstudiante).getId();
    } else {
      var idEstudiante = '1Q75_hgDgdNVAtgBUqjlhIqL41z3PrlrBqLVK5iWv8sc';
      //var confirmacionIdEstudiante = DriveApp.getFileById('1Q75_hgDgdNVAtgBUqjlhIqL41z3PrlrBqLVK5iWv8sc').makeCopy(cartaEstudiante).getId();
    }
    Logger.log(respuestaEvaluadores.getRange(filaEvaluadores, 53).setFormula('=QUERY(' + querySheet + ', ' + '"SELECT C, D, E, J, K, M, N, O, P WHERE AF CONTAINS ' + "'" + URL + "'" + '")').getValues());
    var confirmacionIdEstudiante = DriveApp.getFileById(idEstudiante).makeCopy(cartaEstudiante).getId();
    var confirmacionEstudiante = DocumentApp.openById(confirmacionIdEstudiante);
    var cuerpoConfirmacionEstudiante = confirmacionEstudiante.getActiveSection();
    //Para TODO
    cuerpoConfirmacionEstudiante.replaceText("%fechaCartaInicio%", fechaCartaInicio);
    cuerpoConfirmacionEstudiante.replaceText("%nombreEstudiante%", nombreEstudiante);
    cuerpoConfirmacionEstudiante.replaceText("%tituloDocumento%", tituloDocumento);
    cuerpoConfirmacionEstudiante.replaceText("%programa%", programa);
    cuerpoConfirmacionEstudiante.replaceText("%nombreEvaluadorPropuesta%", nombreEvaluadorPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%nombreEvaluadorFinal%", nombreEvaluadorFinal);
    cuerpoConfirmacionEstudiante.replaceText("%confidencialidad%", confidencialidad);
    //Para PROPUESTA
    cuerpoConfirmacionEstudiante.replaceText("%formulacionPropuesta%", formulacionPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionFormulacionPropuesta%", justificacionFormulacionPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%objetivosPropuesta%", objetivosPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionObjetivosPropuesta%", justificacionObjetivosPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%marcoConceptualPropuesta%", marcoConceptualPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMarcoConceptualPropuesta%", justificacionMarcoConceptualPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%metodologiaPropuesta%", metodologiaPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMetodologiaPropuesta%", justificacionMetodologiaPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%metodosProcesamientoPropuesta%", metodosProcesamientoPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMetodosProcesamientoPropuesta%", justificacionMetodosProcesamientoPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%impactoResultadosPropuesta%", impactoResultadosPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionImpactoResultadosPropuesta%", justificacionImpactoResultadosPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%cronogramaPropuesta%", cronogramaPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionCronogramaPropuesta%", justificacionCronogramaPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%presupuestoPropuesta%", presupuestoPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionPresupuestoPropuesta%", justificacionPresupuestoPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%recomendacionAceptacionPropuesta%", recomendacionAceptacionPropuesta);
    cuerpoConfirmacionEstudiante.replaceText("%revisarDocumentoFinalPropuesta%", revisarDocumentoFinalPropuesta);
    //para FINAL
    cuerpoConfirmacionEstudiante.replaceText("%formulacionFinal%", formulacionFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionFormulacionFinal%", justificacionFormulacionFinal);
    cuerpoConfirmacionEstudiante.replaceText("%objetivosFinal%", objetivosFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionObjetivosFinal%", justificacionObjetivosFinal);
    cuerpoConfirmacionEstudiante.replaceText("%marcoConceptualFinal%", marcoConceptualFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMarcoConceptualFinal%", justificacionMarcoConceptualFinal);
    cuerpoConfirmacionEstudiante.replaceText("%metodologiaFinal%", metodologiaFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMetodologiaFinal%", justificacionMetodologiaFinal);
    cuerpoConfirmacionEstudiante.replaceText("%metodosProcesamientoFinal%", metodosProcesamientoFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionMetodosProcesamientoFinal%", justificacionMetodosProcesamientoFinal);
    cuerpoConfirmacionEstudiante.replaceText("%alcancesInnovacionFinal%", alcancesInnovacionFinal);
    cuerpoConfirmacionEstudiante.replaceText("%calidadTrabajoFinal%", calidadTrabajoFinal);
    cuerpoConfirmacionEstudiante.replaceText("%justificacionCalidadTrabajoFinal%", justificacionCalidadTrabajoFinal);
    cuerpoConfirmacionEstudiante.replaceText("%comentariosGeneralesFinal%", comentariosGeneralesFinal);
    cuerpoConfirmacionEstudiante.replaceText("%recomiendaSustentacionFinal%", recomiendaSustentacionFinal);
    cuerpoConfirmacionEstudiante.replaceText("%aceptaSerPresidenteFinal%", aceptaSerPresidenteFinal);
    confirmacionEstudiante.saveAndClose();
    var pdfConfirmacionEstudiante = DriveApp.createFile(confirmacionEstudiante.getAs("application/pdf"));
    var pdfConfirmacionId = pdfConfirmacionEstudiante.getId();
    var subjectEstudiante = "Respuesta del evaluador de " + tipoDocumento + ".";
    var bodyEstudiante = ("Respetad@s " + nombreEstudiante + " y " + nombreDirector + ".\n\nLa coordinación del posgrado hace envío de la evaluación de su " + tipoDocumento + ".\nFelicitaciones!"); 
    Logger.log(respuestaEvaluadores.getRange(filaEvaluadores, 55).getValue());
    var correos = ("evaluacion.posgrado.pisa@correounivalle.edu.co," + respuestaEvaluadores.getRange(filaEvaluadores, 55).getValue()+"," + respuestaEvaluadores.getRange(filaEvaluadores, 57).getValue());
    Logger.log(correos);
    MailApp.sendEmail(correos, subjectEstudiante , bodyEstudiante, {
      attachments: [pdfConfirmacionEstudiante.getAs(MimeType.PDF)],
    });
    var folderEvaluatorsRequest = DriveApp.getFolderById("1DIr8mXJCwmc_2prV1gIZzWcqyJdK4G7_");
    var fileEstudiante = DriveApp.getFileById(pdfConfirmacionEstudiante.getId());
    folderEvaluatorsRequest.addFile(fileEstudiante);
    var root = DriveApp.getRootFolder();
    root.removeFile(pdfConfirmacionEstudiante);
    DriveApp.getFileById(confirmacionIdEstudiante).setTrashed(true);
  }
}
