function onSubmit(e) {
  var nombre = e.namedValues['Correo electrónico/Email student'];
  
  //1 crear una hoja y poner nombre
  //var newSheet = createSpreadsheet(nombre);
  var newSheet = createSheet(nombre);

  //2 get palabras claves
  var palabras = e.namedValues['Palabra(s) claves /key words '];
  var listaPalabras = palabras.toString().split(',');

  //3 get evaluadores list
  var listaColumnasBusqueda = ['O','P','Q'];
  var rowsMatch = [];
  for(var i=0; i<listaColumnasBusqueda.length; i++){
    var listaPalabrasEvaluadores = getEvaluadoresListaPalabras(listaColumnasBusqueda[i]);
    rowsMatch = getRowsMatch(rowsMatch, listaPalabras,listaPalabrasEvaluadores);
  }
  
  //4 copy evaluadores to new sheet
  for(var i=0; i<rowsMatch.length; i++){
    var evaluadorData = getEvaluadorRowData(rowsMatch[i]);
    newSheet.appendRow(evaluadorData);
  }
  
  //5 calcular porcentaje
  //Problemas con el cálculo.
  
  //6 crear menu email
  
  //7 crear email template
  //se va a usar el template? O se formateará el HTML en el cuerpo? O se enviará un adjunto con el template lleno?
}

//Se cambia el parámetro name por nombre, se copia el campo de nombre de columnas y el valor correspondiente al newSubmit en la hoja nueva.
function createSheet(nombre){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var origin = ss.getSheetByName('Respuestas Propuestas y Tesis');
  var newsheet = ss.insertSheet().setName(nombre);
  //var data1 = origin.getRange(1, 1, 1, origin.getLastColumn()).getValues();
  //newsheet.appendRow(data1[0]);

  return newsheet;
}

//Se cambia la hoja "evaluadores" por "Aplica para doctorado y maestría".
function getEvaluadoresListaPalabras(columna){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var evaluadoresSheet = ss.getSheetByName('Aplica para doctorado y maestría');
  var evaluadoresLastRow = evaluadoresSheet.getLastRow();
  var columnaOA1Notation = columna + '2:'+ columna + evaluadoresLastRow.toString();
  var evaluadoresListaPalabras = evaluadoresSheet.getRange(columnaOA1Notation);
  return evaluadoresListaPalabras.getValues();
}

function getRowsMatch(rowsMatch, palabras, palabrasEvaluador){
  for(var lp=0; lp<palabras.length; lp++){
    for(var i=0; i<palabrasEvaluador.length; i++){
      if(palabrasEvaluador[i][0].indexOf(palabras[lp]) != -1){
        if(rowsMatch.indexOf(i+2) == -1){
         rowsMatch.push(i+2);
        }
      }
    }
  }
  return rowsMatch;
}

function getEvaluadorRowData(row){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var evaluadoresSheet = ss.getSheetByName('Aplica para doctorado y maestría');
  var evaluadorRow = evaluadoresSheet.getRange(row, 1, 1, evaluadoresSheet.getLastColumn()-1);
  return evaluadorRow.getValues()[0];
}

//función que genera un menú y su correspondiente ítem, para el envío de correos a los posibles evaluadores

function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu('Enviar Correos')
  .addItem('Enviar correos a Evaluadores seleccionados', 'mailToEvaluators')
  .addToUi();
}

function mailToEvaluators(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var evaluadoresSheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeSelected = evaluadoresSheet.getActiveRange();
  var ui = SpreadsheetApp.getUi();
  var pregunta = 'Estas seguro de enviar el email a los '+rangeSelected.getHeight()+' evaluadores seleccionados?';
  var response = ui.alert(pregunta , ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    var emails =[];
    for(var i = 0; i<rangeSelected.getHeight(); i++) {
      var row = rangeSelected.getRow() + i;
      var email = evaluadoresSheet.getRange('S'+row).getValues()[0][0];
      if(email != '' && emails.indexOf(email) == -1){
        emails.push(email);
        MailApp.sendEmail(email, "Solicitud de Evaluacion", prepareEmailBody(evaluadoresSheet.getName()), {
          htmlBody:getHtmlBody(evaluadoresSheet.getName())
        });
      }
    }
  }
}

function prepareEmailBody(estudianteEmail){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var respuestas = ss.getSheetByName('Respuestas Propuestas y Tesis');
  var columnaOA1Notation = 'E2:E'+respuestas.getLastRow().toString();
  var respuestasData = respuestas.getRange(columnaOA1Notation).getValues();
  var estudentIndex = 0;
  for(var i=0; i<respuestasData.length; i++){
    if(respuestasData[i].indexOf(estudianteEmail) != -1){
      estudentIndex = i +1;
      break;
    }
  }
  var body = '';
  if(estudentIndex){
    var estudianteRow = respuestas.getRange(estudentIndex+1, 1, 1, respuestas.getLastColumn()).getValues()[0];
    body = "Apreciados investigadores,  conocedores de su experticia profesional, y su disposición para apoyar en los procesos de evaluación del programa de posgrados de Ingeniería Sanitaria  y Ambiental, la dirección del programa le agradecería su apoyo con la evaluación del documento: Por favor no reenvíe este correo.  Entré al link correspondiente a su respuesta.  Para la evaluación si desea aceptar contará con tres semanas a partir de la fecha de recibo de este mensaje. Si no acepta y conoce a otro investigador que nos pueda recomendar por favor envíenos la información para contactarlo junto con las razones por la cuales no puede aceptar.";
    body += '\nPrograma/Program: ' + estudianteRow[1];
    body += '\nNombre(s)/First name: ' + estudianteRow[2];
    body += '\nApellido(s)/Last name: ' + estudianteRow[3];
    body += '\nCorreo electrónico/Email student: ' + estudianteRow[4];
    body += '\nTítulo/Title: ' + estudianteRow[20];
    body += '\nResumen/Summary: ' + estudianteRow[21];
    body += '\nTipo de Trabajo/Type of work: ' + estudianteRow[25];
    body += '\n\n Si su respuesta es positiva por favor ingrese a este link Acepto';
    body += '\n Si su respuesta es negativa por favor ingrese a este link No acepto';
    body += '\n\n Muchas Gracias¡';
    body += '\n\n Cordialmente,';
    body += '\nJaneth Sanabria Gómez';
    body += '\nPosgrado en Ingeniería Sanitaria y Ambiental';
    body += '\nEscuela de Recursos Naturales y del Ambiente';
    body += '\nEdificio 341 Sede Meléndez ';
    body += '\nCalle 13 # 100-00';
    body += '\nTel 3302002';
  }
  return body;
}

function getHtmlBody(estudianteEmail){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var respuestas = ss.getSheetByName('Respuestas Propuestas y Tesis');
  var columnaOA1Notation = 'E2:E'+respuestas.getLastRow().toString();
  var respuestasData = respuestas.getRange(columnaOA1Notation).getValues();
  var estudentIndex = 0;
  for(var i=0; i<respuestasData.length; i++){
    if(respuestasData[i].indexOf(estudianteEmail) != -1){
      estudentIndex = i +1;
      break;
    }
  }
  var body = '';
  if(estudentIndex){
    var estudianteRow = respuestas.getRange(estudentIndex+1, 1, 1, respuestas.getLastColumn()).getValues()[0];
    body = "Apreciados investigadores,  conocedores de su experticia profesional, y su disposición para apoyar en los procesos de <b>evaluación del programa de posgrados de Ingeniería Sanitaria  y Ambiental</b>, la dirección del programa le agradecería su apoyo con la evaluación del documento: Por favor no reenvíe este correo.  Entré al link correspondiente a su respuesta.  Para la evaluación si desea aceptar contará con tres semanas a partir de la fecha de recibo de este mensaje. Si no acepta y conoce a otro investigador que nos pueda recomendar por favor envíenos la información para contactarlo junto con las razones por la cuales no puede aceptar.";
    body += '<table>';
    body += '<tr><td>Programa/Program</td><td>' + estudianteRow[1]+'</td></tr>';
    
    body += '<tr><td>Nombre(s)/First name</td><td>' + estudianteRow[2]+'</td></tr>';
    body += '<tr><td>Apellido(s)/Last name</td><td>' + estudianteRow[3]+'</td></tr>';
    body += '<tr><td>Correo electrónico/Email student</td><td>' + estudianteRow[4]+'</td></tr>';
    body += '<tr><td>Título/Title</td><td>' + estudianteRow[20]+'</td></tr>';
    body += '<tr><td>Resumen/Summary</td><td>' + estudianteRow[21]+'</td></tr>';
    body += '<tr><td>Tipo de Trabajo/Type of work</td><td>' + estudianteRow[25]+'</td></tr></table>';
    body += '<tr><td>Documento/Document</td><td>' + estudianteRow[31]+'</td></tr></table>';
    body += '<br><br> Si su respuesta es positiva por favor ingrese a este link <a href="https://docs.google.com/a/correounivalle.edu.co/forms/d/e/1FAIpQLSdsWkc6QkKRPfIXFA1vMXPgRCVf2SfZiToKVdA9F0lov92nIQ/viewform" target="_blank">Acepto</a>';
    body += '<br> Si su respuesta es negativa por favor ingrese a este link <a href="https://docs.google.com/a/correounivalle.edu.co/forms/d/e/1FAIpQLSefF5FDBq7nv3vfAcs6J5dHhmSYSGOwQ6sgPvxMIegn3UIiKQ/viewform" target="_blank">No acepto</a>';
    body += '<br><br> Muchas Gracias¡';
    body += '<br><br>Cordialmente,';
    body += '<br><br>Janeth Sanabria Gómez';
    body += '<br>Posgrado en Ingeniería Sanitaria y Ambiental';
    body += '<br>Escuela de Recursos Naturales y del Ambiente';
    body += '<br>Edificio 341 Sede Meléndez ';
    body += '<br>Calle 13 # 100-00';
    body += '<br>Tel 3302002';
  }
  return body;
} 

