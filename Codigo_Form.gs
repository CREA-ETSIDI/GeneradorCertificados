//Autor: Leonel Aguilera
/*Codigo escrito el 17 de octubre del 2020*/

/*Se que no es un código limpio
  Se que hacerle mantenimiento va a ser jodido
  Pero me da I G U A L
  ¡FURULA!
  (o eso creo)
  Y eso es lo que importa
  Jodeos junta del futuro xD*/


function Sort(){
  let Datos = SpreadsheetApp.openById("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0");
  let Respuestas = SpreadsheetApp.openById("1QSRUln18CLkK_MUrFVinLWM57xyOKwMEptMnvugfZeM").getSheets()[0];
  
  //let last = Respuestas.getLastRow();
  let last = 13;
  Logger.log("Last: "+last);
  let fecha = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
  let email = Respuestas.getRange(last, 2).getValue();
  let nombre = Respuestas.getRange(last, 3).getValue();
  let matricula = Respuestas.getRange(last, 4).getValue();
  let DNI = Respuestas.getRange(last, 5).getValue();
  let rol = Respuestas.getRange(last, 6).getValue();
  let curso = Respuestas.getRange(last, 7).getValue();
  let periodo = Respuestas.getRange(last, 8).getValue();
  Logger.log("Email: "+email);
  
  if(rol == "Alumnx, recibí el curso") {
    let URLs = Datos.getSheetByName("URLs");
    
    let fila = AnalizarPorColumna(2,URLs,1,curso);
    let columna = AnalizarPorFila(2,URLs,1,periodo);
    
    Logger.log("fila: "+fila+"; columna: "+columna);
    let URL = URLs.getRange(fila, columna).getValue();
    Logger.log("URL: "+URL);
    if (URL == "NOPE")
    {
      EnviarError(nombre,email,curso,1);
    }
    else if (URL == "INEXISTENTE")
    {
      EnviarError(nombre,email,curso,6);
    }
    else if (URL == "")
    {
      EnviarError(nombre,email,curso,2);
    }
    else if (Datos.getSheetByName("Horas").getRange(fila, columna).getValue() == "" || Datos.getSheetByName("Plantillas").getRange(fila, 2).getValue() == "")
    {
      EnviarError(nombre,email,curso,3);
    }
    else {
      Logger.log("Llegó al else")
      var Participantes = SpreadsheetApp.openByUrl(URL);
      
      if(ComprobarRealizacion(email,Participantes.getSheets()[0],2)) {
        Logger.log("Ejecutando: Generar documento("+nombre+","+email+","+DNI+","+fecha+","+fila+","+columna+",2)");
        GenerarDocumento(nombre,email,DNI,fecha,fila,columna,2);
      }
      else {
        EnviarError(nombre,email,curso,4);
      }
    }
  }
  else {
    let URLs = Datos.getSheetByName("URLs");
    //Es profo, need more info
    let fila = AnalizarPorColumna(2,URLs,1,curso);
    let columna = AnalizarPorFila(2,URLs,1,periodo);
    if (Datos.getSheetByName("Horas").getRange(fila, columna).getValue() == "" || Datos.getSheetByName("Plantillas").getRange(fila, 2).getValue() == "")
    {
      EnviarError(nombre,periodo,curso,3);
    }
    else {
      if(Datos.getSheetByName("ListaEmailsProfes").getRange(fila, columna).getValue().indexOf(email) > -1)
      {
        Logger.log("Ejecutando: Generar documento("+nombre+","+email+","+DNI+","+fecha+","+fila+","+columna+",3)");
        GenerarDocumento(nombre,email,DNI,fecha,fila,columna,3);
      }
      else
      {
        EnviarError(nombre,email,curso,5);
      }
    }
  }
}

//Si no se encuentra a la persona que rellenó el formulario, le envía un correo
function EnviarError(Nombre,Email,NombreCurso,Error) {
  Logger.log("Enviando error #"+Error);
  switch(Error) {
    case 1:
      //NOPE
      GmailApp.sendEmail(Email, "Certificado CREA", "Hola "+Nombre+'\n'+"Rellenaste un formulario para pedir un certificado del "+NombreCurso+", sin embargo, no estamos generando certificados para el periodo seleccionado, si crees que ha cometido un error, escríbenos.");
      break;
    case 2:
      GmailApp.sendEmail(Email, "Certificado CREA", "Hola "+Nombre+'\n'+"Rellenaste un formulario para pedir un certificado del "+NombreCurso+", sin embargo, aún no hemos empezado a generar certificados para este en el periodo seleccionado, por favor, vuelva a intentarlo en un par de días.");
      GmailApp.sendEmail("roboticaeuiti@gmail.com","Faltan listas", "Heeey... Sep, soy yo, el robot de los certificados y vengo a comentaros que alguien acaba de pedir un certificado, sin embargo cuando rellenasteis la hoja de cálculo para el "+NombreCurso+" faltaron datos (o bien la hoja de cálculo de respuestas del formulario o la URL de esta), así que, primero que nada, arreglad eso (o programadme para que lo haga yo) y luego rellenad manualmente el certificado (aunque también podeis llamarme manualmente a mi para que me encargue de ello ejecutando la función Sort())");
      break;
    case 3 :
      GmailApp.sendEmail("roboticaeuiti@gmail.com","Faltan plantillas... Vagos", "Heeey... Sep, soy yo, el robot de los certificados y vengo a comentaros que alguien acaba de pedir un certificado, sin embargo cuando rellenasteis la hoja de cálculo para el "+NombreCurso+" faltaron datos (o bien las hojas certificadas o la ID de la plantilla), así que, primero que nada, arreglad eso (o programadme para que lo haga yo) y luego rellenad manualmente el certificado (aunque también podeis llamarme manualmente a mi para que me encargue de ello ejecutando la función Sort())");
      GmailApp.sendEmail(Email, "No-Certificado CREA", "Hola "+Nombre+'\n'+"Rellenaste un formulario para pedir un certificado del "+NombreCurso+", sin embargo, aún no hemos empezado a generar certificados para este en el periodo seleccionado, por favor, espere a que solucionemos esto y le enviaremos su certificado en cuanto sea posible.")
    case 4:
      GmailApp.sendEmail(Email, "Certificado CREA", "Hola "+Nombre+'\n'+"Rellenaste un formulario para pedir un certificado del "+NombreCurso+", sin embargo, nuestro robot no te ha encontrado en la lista de inscripción a ese curso para el periodo seleccionado, si crees que ha cometido un error, escríbenos");
      break;
    case 5:
      GmailApp.sendEmail(Email, "Certificado CREA", "Hola "+Nombre+'\n'+"Rellenaste un formulario para pedir un certificado del "+NombreCurso+", sin embargo, nuestro robot no te ha encontrado en la lista de profes de ese curso para el periodo seleccionado, si crees que ha cometido un error, escríbenos");
      break;
    case 6:
      //INEXISTENTE
      GmailApp.sendEmail(Email, "Certificado CREA", "Hola "+Nombre+'\n'+"Rellenaste un formulario para pedir un certificado del "+NombreCurso+", sin embargo, este curso no se realizó el periodo seleccionado, si crees que ha cometido un error, escríbenos");
      break;
    default:
      GmailApp.sendEmail("roboticaeuiti@gmail.com", "Fallo en el script de certificados CREA", "El script de alguna forma se ha roto ¿Lo habeis retocado?, da igual, arregladlo... ¡RÁPIDO!\n\n\n\nAh, sí, por cierto, pequeño detalle importante... Para que se enviase este mensaje el script tiene que haberse ejecutado, así que... Una de dos... O bien lo estais probando (en cuyo caso, enhorabuena, la habeis cagado) o bien el activador lo ha ejecutado, así que mirad las respuestas al formulario no vaya a ser que alguien se quede sin certificado");
  }      
}

function GenerarDocumento(Nombre,Email,DNI,fecha,fila,columna,rol) {
  //Do stuff
  Logger.log("Generando documento...")
  let CertNom = "Certificado_"+Nombre;
  let datos = SpreadsheetApp.openById("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0");
  let respuestas = SpreadsheetApp.openById("1QSRUln18CLkK_MUrFVinLWM57xyOKwMEptMnvugfZeM").getSheets()[0];
  Logger.log("Datos abiertos")
  let ID = DriveApp.getFileById(datos.getSheetByName("Plantillas").getRange(fila, rol).getValue()).makeCopy(CertNom, DriveApp.getFolderById("1ikZ6a9S2RDWRPdY1igsdEnGq1oLYkGvv")).getId();
  Logger.log("Copia hecha")
  let Slide = SlidesApp.openById(ID);
  
  let edicion = datos.getSheetByName("URLs").getRange(1,columna).getValue();
  let horas = datos.getSheetByName("Horas").getRange(fila,columna).getValue();
  
  Logger.log("Empezando a reemplazar los placeholders")
  ID = ReplaceShit(Slide,Nombre,DNI,edicion,fecha,horas);  
  Logger.log("Placeholders reemplazados");
  
  //Utilities.sleep(2000); //Esto de editar un documento y hacerle trastadas al instante puede salir mal, es mejor tomárselo con calma y esperar
  //De verdad que está dando mucho por saco... La presentación se insta edita dpm, pero el pdf sale con los malditos placeholders y ya no sé qué hacer
  //...
  //...
  //...
  //Pues al final el delay no hacía falta xD
  //Por lo visto eso era un "bug" producido porque no se guardan los cambios hechos por el script hasat que este termine su ejecución, para eso está el comando saveAndClose() añadido en ReplaceShit, que fuerza a que se guarden los cambios.
  
  Logger.log("Siesta echada");
  Slide = DriveApp.getFileById(ID);
  let Blob = Slide.getBlob().getAs('application/pdf');
  Logger.log("Convertido a pdf");
  let PDFf = DriveApp.getFolderById("1ikZ6a9S2RDWRPdY1igsdEnGq1oLYkGvv").createFile(Blob).setName(CertNom + ".pdf");
  Logger.log("PDF generado");
  Slide.setTrashed(true);
  GmailApp.sendEmail(Email, 'Tu Certificado', 'Hola '+Nombre+'\n\nAquí tienes el certificado de tu curso CREA', {
    attachments: [PDFf.getAs(MimeType.PDF)],
    name: 'Certificados CREA'
  });
  Logger.log("email enviado");
  Slide.setTrashed(true);
  Logger.log("Editable eliminado");
  PDFf.makeCopy(DriveApp.getFolderById("10cBHSXkjFxQ65osf3GJ2mdyp6WFM2nUm"));
  Logger.log("Copia guardada en certificados expedidos");
  Logger.log("Fin del programa");
}

function AnalizarPorColumna(i0,Lista,columna,match) {
  let fila = false;
  for (let i = i0;i<=Lista.getLastRow();i++) {
    if(match==Lista.getRange(i, columna).getValue()) {
      fila = i;
    }
  }
  return fila;
}

function AnalizarPorFila(i0,Lista,fila,match) {
  let columna = false;
  for (let i = i0;i<=Lista.getLastColumn();i++) {
    if(match==Lista.getRange(fila,i).getValue()) {
      columna = i;
    }
  }
  return columna;
}

function ComprobarRealizacion(email,sheetobject,columna) {
  let numero = sheetobject.getLastRow();
  
  for(let i = 1; i <= numero; i++)
  {
    if(sheetobject.getRange(i, columna).getValue()==email)
    {
      return true;
      break;
    }
  }
  return false;
}

function GetItemsID() {
  var form = FormApp.openById("1P5iAwi8FKvmAae_eoVrPQF-kKqtvZAqFetJwdxT57-o");
  var items = form.getItems();
  for (var i in items)
  {
    Logger.log(i+' '+items[i].getTitle()+': '+items[i].getId());
  }
}

function ColumnaExtra() {
  let a = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
  let b=a[3]+a[4];
  Logger.log(b);
  if(b=="08") {
    let spreadsheet = SpreadsheetApp.openById("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0");
    let sheet1 = spreadsheet.getSheetByName("URLs");
    let sheet2 = spreadsheet.getSheetByName("Horas");
    let sheet3 = spreadsheet.getSheetByName("ListaEmailsProfes");
    let numero = sheet1.getLastColumn();
    sheet1.insertColumnAfter(numero);
    sheet2.insertColumnAfter(numero);
    sheet3.insertColumnAfter(numero);
    let yearA = numero+2016;
    //NOTA: esta parte del código no está hecha para funcionar eternamente y tendrá que ser actualizada a partir del año 9999 D.C. o con cualquier cambio de calendario
    let yearB = yearA+1-1000*Math.floor(yearA/1000);
    yearB=yearB-100*Math.floor(yearB/100);
    if(yearB<10) {
      yearB="0"+yearB;
    }
    sheet1.getRange(1, numero+1).setValue("Curso "+yearA+"/"+yearB);
    sheet2.getRange(1, numero+1).setValue("Curso "+yearA+"/"+yearB);
    sheet3.getRange(1, numero+1).setValue("Curso "+yearA+"/"+yearB);
    
    FormApp.openById('1P5iAwi8FKvmAae_eoVrPQF-kKqtvZAqFetJwdxT57-o').getItemById(1321130332).asMultipleChoiceItem().setChoiceValues(sheet1.getRange(1, 2, 1, numero).getValues()[0]);
  }
  
  
  /*
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  var form = FormApp.openById('1P5iAwi8FKvmAae_eoVrPQF-kKqtvZAqFetJwdxT57-o');
  ScriptApp.newTrigger("Sort").forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger("ColumnaExtra").timeBased().atDate(yearA+1, 8, 31).create();
  */
}


function getShapeID() {
  var slide = SlidesApp.openById("1ZT2JssnGFElzSejLgqh3T77gzEH-6hhSENdmW9GDpTw").getSlides()[0];
  var shapes = slide.getShapes();
  for (i=0; i<shapes.length; i++) {
    Logger.log("Texto = "+shapes[i].getText().asString()+"  ID: "+shapes[i].getObjectId());
  }
  /*const objectId = "ga40a1a51c9_2_12";
  const DNI = "Placeholder";
  var objeto = getShapeByIDvCREA(slide,objectId);
  objeto.getText().setText("Placeholder");
  objeto = getShapeByIDvCREA(slide,"ga40b73e421_0_1");
  objeto.getText().setText("Con DNI "+DNI+", la participación y finalización del");*/
}

function getShapeByIDvCREA(slide,ID) {
  var shapes = slide.getShapes();
  for (i=0; i<shapes.length; i++) {
    if(shapes[i].getObjectId()==ID)
    {
      return shapes[i];
      break;
    }
  }
}

function asdf() {
  /*
  var string = "1/10/2020 11:29:41"
  Logger.log(NumToFecha(string));
  
  var str = "asdf algo imaginacion task patata palabra random"
  var newstr = "";
  if (str.indexOf("task") > -1) {
    Logger.log(str.indexOf("task"));
    for(let i=0;i<str.length ;i++) {
      if(i==str.indexOf("task")) {
        newstr += "nueva cosa";
        i += 3;
      }
      else {
        newstr += str[i];
      }
    }
    Logger.log(newstr);
  }
  */
  let a = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
  let b=a[3]+a[4];
  Logger.log(a);
  Logger.log(b);
}

function NumToFecha(fecha) {
  Logger.log(fecha);
  let string = fecha[0];
  let i = 1;
  if (fecha[1] != '/') {
    string = string+fecha[1];
    i++;
  }
  while(fecha[i] != " ") {
    string = string + fecha[i];
    i++;
    Logger.log("fecha[i]: "+fecha[i]);
    Logger.log("string: "+string);
  }
  return string;
}

function ReplaceShit(Presentacion, nombre,DNI,edicion,fecha,horas) {
  let marcas = ["%NOMBRE%","%DNI%","%EDICION%","%FECHA%","%HORAS%"];
  let replace = [nombre, DNI, edicion, fecha, horas];
  let offset = [7,4,8,6,6];
  let Slide = Presentacion.getSlides()[0];
  
  let shapes = Slide.getShapes()
  for (let i=0; i<shapes.length; i++) {
    let texto = shapes[i].getText().asString();
    for (let j=0; j<marcas.length;j++) {
      if (texto.indexOf(marcas[j]) > -1) {
        let nuevotexto = "";
        for(let k=0;k<texto.length ;k++) {
          if(k==texto.indexOf(marcas[j])) {
            nuevotexto += replace[j];
            k += offset[j];
          }
          else {
            nuevotexto += texto[k];
          }
        }
        texto = nuevotexto;
      }
    }
    shapes[i].getText().setText(texto);
  }
  Presentacion.saveAndClose();
  return Presentacion.getId();
}


/*
[20-10-19 20:44:35:127 CEST] Texto = Certificado
  ID: ga40a1a51c9_2_3
  
[20-10-19 20:44:35:149 CEST] Texto = Club de Robótica Electrónica y Automática
  ID: ga40a1a51c9_2_6
  
[20-10-19 20:44:35:154 CEST] Texto = Certifica a
  ID: ga40a1a51c9_2_9
  
[20-10-19 20:44:35:160 CEST] Texto = Placeholder nombre
  ID: ga40a1a51c9_2_12
  
[20-10-19 20:44:35:166 CEST] Texto = Firmado,
Junta Directiva del Club de Robótica, Electrónica y Automática (CREA)
  ID: ga40a1a51c9_2_19
  
[20-10-19 20:44:35:171 CEST] Texto = Fecha
  ID: ga40a1a51c9_2_20
  
[20-10-19 20:44:35:177 CEST] Texto = 
  ID: ga40a1a51c9_0_3
  
[20-10-19 20:44:35:181 CEST] Texto = 
  ID: ga40a1a51c9_0_5
  
[20-10-19 20:44:35:186 CEST] Texto = 
  ID: ga40a1a51c9_0_2
  
[20-10-19 20:44:35:191 CEST] Texto = Con DNI Placeholder, la participación y finalización del
  ID: ga40b73e421_0_1
  
[20-10-19 20:44:35:195 CEST] Texto = Curso Arduino Zero edición Placeholder constando de Placeholder horas de formación
  ID: ga40b73e421_0_3
  
[20-10-19 20:44:35:201 CEST] Texto = Placeholder fecha
  ID: ga40c1880c8_0_1
*/




/*
[20-10-17 04:59:20:148 PDT] 0 Nombre y apellidos: 1633920210
[20-10-17 04:59:20:225 PDT] 1 Número de matrícula: 790080973
[20-10-17 04:59:20:287 PDT] 2 DNI: 380188083
[20-10-17 04:59:20:376 PDT] 3 Datos del curso asistido/impartido: 306130935
[20-10-17 04:59:20:439 PDT] 4 ¿Cual fue tu rol en el curso?: 1902310025
[20-10-17 04:59:20:504 PDT] 5 ¿Qué tipo de curso fue?: 1317373257
[20-10-17 04:59:20:589 PDT] 6 ¿Cuándo fue el curso?: 1321130332
*/