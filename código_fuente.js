  var idTemporal= "14zieMzL0hdfeccNAdvSAEplQsYi8ki3J";
  var idDocPlantilla= "1Ayk0Hxrc09YvX5jArdEdt95OkFa2ANReuF4c1rj0-MU";

  var colCodigo= 2;
  var colNombres= 3;
  var colApellidos= 4;
  var colNota= 5;
  var colObservacion= 6;
  var colEmail= 7;
  var colPDF= 8;

function generarPDfs(){
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos")
  var alumnos= hoja.getRange(2,1,hoja.getLastRow()-1,hoja.getLastColumn()).getValues();

  alumnos.forEach(function(alumno,i){
  var estudiante={}
  estudiante.codigo = alumno[colCodigo-1];
  estudiante.nombres = alumno[colNombres-1];
  estudiante.apellidos = alumno[colApellidos-1];
  estudiante.nota = alumno[colNota-1];
  estudiante.observaciones = alumno[colObservacion-1];
  estudiante.email = alumno[colEmail-1];
  if(!alumno[colPDF-1]){
  var urlPDF = generarPDF(estudiante);
  hoja.getRange(i+2,colPDF).setValue(urlPDF);
  estudiante.pdf=urlPDF;
  enviarMail(estudiante);
  }
  })
}
function generarPDF(estudiante) {
  var idTemporal= "14zieMzL0hdfeccNAdvSAEplQsYi8ki3J";
  var idDocPlantilla= "1Ayk0Hxrc09YvX5jArdEdt95OkFa2ANReuF4c1rj0-MU";

  var doc=DocumentApp.openById(idDocPlantilla);
  var archivoPlantilla= DriveApp.getFileById(idDocPlantilla);
  var carpeta=DriveApp.getFolderById(idTemporal);

  var hoja=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos");



  //copia
  var nombredoc="Certificado de: "+estudiante.nombres+" "+estudiante.apellidos;
  var copiaArchivoPlantilla= archivoPlantilla.makeCopy(carpeta)
  copiaArchivoPlantilla.setName(nombredoc);
  


  var copiaId= copiaArchivoPlantilla.getId();
  var doc= DocumentApp.openById(copiaId);
  doc.setName(nombredoc);
  
  doc.getBody().replaceText("%Código%",estudiante.codigo);
  doc.getBody().replaceText("%Nombres%",estudiante.nombres);
  doc.getBody().replaceText("%Apellidos%",estudiante.apellidos);
  doc.getBody().replaceText("%Calificación%",estudiante.nota);
  doc.getBody().replaceText("%Observación%",estudiante.observaciones);

  doc.saveAndClose();
  const pdfBlod= copiaArchivoPlantilla.getAs(MimeType.PDF);
  var pdfcreado=carpeta.createFile(pdfBlod);
  pdfcreado.addViewer(estudiante.email);
  var urlPDF=pdfcreado.getUrl();
  return urlPDF;

}

function enviarMail(estudiante){
  var mensaje= "Joven "+estudiante.nombres+" "+estudiante.apellidos+
               ", "+"este es el reporte de la Nota Final "+ estudiante.pdf;
  MailApp.sendEmail(estudiante.email, "Reporte de Notas", mensaje);
}

function onOpen(){
  var ui= SpreadsheetApp.getUi();
  var menu= ui.createMenu("Resportes");
  menu.addItem("Generar Calificaciones","generarPDfs").addToUi();

}
