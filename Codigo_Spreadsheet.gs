function Start()
{
  var menu = SpreadsheetApp.getUi().createMenu('Cursos');
  SpreadsheetApp.openById("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0").toast("No modificar nada de esta hoja de cálculo sin antes haber leido el manual de uso","⚠️¡ALERTA!⚠️",5);
  menu.addItem('Actualizar formulario', 'Actualizador').addToUi();
  menu.addItem('Añadir fila', 'Fila').addToUi();
  menu.addItem('Quitar fila', 'FilaM').addToUi();
  Actualizador();
  
}
function Actualizador() {
  /*
  let sheet1 = SpreadsheetApp.openById("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0").getSheetByName("URLs");
  let sheet2 = SpreadsheetApp.openById("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0").getSheetByName("Horas");
  let sheet3 = SpreadsheetApp.openById("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0").getSheetByName("Plantillas");
  */
  let sheets = SpreadsheetApp.openById("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0").getSheets();
  
  let formulario = FormApp.openById("1P5iAwi8FKvmAae_eoVrPQF-kKqtvZAqFetJwdxT57-o");
  
  let sht = 0
  for(let i=0;i<sheets.length;i++)
  {
    if(sheets[i].getLastRow() > sheets[sht].getLastRow())
    {
      sht=i;
    }
  }
  Logger.log("SHT="+sht);
  for(let i=0;i<sheets.length;i++)
  {
    Logger.log("i="+i);
    Logger.log("LastRow="+sheets[i].getLastRow());
    Logger.log("PaternRow="+sheets[i].getLastRow());
    if(sheets[i].getLastRow() < sheets[sht].getLastRow())
    {
      sheets[i].insertRowAfter(sheets[i].getLastRow());
      sheets[i].getRange(sheets[i].getLastRow()+1,1).setValue(sheets[sht].getRange(sheets[i].getLastRow()+1,1).getValue());
      Logger.log(sheets[sht].getRange(sheets[i].getLastRow(),1).getValue())
      i--;
    }
  }
  let QID = 1317373257;
  //Y guardamos en la variable item el objeto de la pregunta del formulario
  let item = formulario.getItemById(QID);
  
  let choices = ["Error 404: Contacte con la junta"];
  Logger.log(sheets[sht].getLastRow())
  for(var i = 2; i <= sheets[sht].getLastRow(); i++)
  {
    Logger.log(i)
    choices[i-2]=sheets[0].getRange(i, 1).getValue();
  }
  for (let j = 0; j < choices.lenght; j++)
  {
    Logger.log("Choice: "+choices[j]);
  }
  item.asMultipleChoiceItem().setChoiceValues(choices);
  
  QID = 1321130332;
  //Y guardamos en la variable item el objeto de la pregunta del formulario
  item = formulario.getItemById(QID);
  
  choices = ["Error 404: Contacte con la junta"];
  for(var i = 2; i <= sheets[0].getLastColumn(); i++)
  {
    Logger.log(i)
    choices[i-2]=sheets[0].getRange(1,i).getValue();
  }
  for (let j = 0; j < choices.lenght; j++)
  {
    Logger.log("Choice: "+choices[j]);
  }
  item.asMultipleChoiceItem().setChoiceValues(choices);
  SpreadsheetApp.openById("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0").toast("Ya se ha actualizado el formulario", "Proceso completado", 5);
  faltanDatos("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0");
}





function Fila() {
  let spreadsheet = SpreadsheetApp.openById("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0");
  let sheets = spreadsheet.getSheets();
  for(let i=0;i<sheets.length;i++)
  {
    sheets[i].insertRowAfter(sheets[i].getLastRow());
  }
  spreadsheet.toast("Fila añadida", "Tarará tararaaaaaa");
}

function FilaM() {
  let spreadsheet = SpreadsheetApp.openById("1std86474xM3Rh70fGTqu9In7OrUDzvAaYRM8v5tqOk0");
  let sheets = spreadsheet.getSheets();
  for(let i=0;i<sheets.length;i++)
  {
    sheets[i].deleteRow(sheets[i].getMaxRows());
  }
  spreadsheet.toast("Fila eliminada", "Tarará tararaaaaaa");
}

function faltanDatos(ID) {
  let SPsheet = SpreadsheetApp.openById(ID);
  let sheets = SPsheet.getSheets();
  
  let fechas = sheets[0].getLastColumn();
  let cursos = sheets[0].getLastRow();
  loop:
  for(let i=2;i<=fechas;i++) {
    for(let j=2;j<=cursos;j++) {
      let texto = sheets[0].getRange(j, i).getValue()
      if(texto != "" && texto != "NOPE" && texto != "INEXISTENTE") {
        if(texto.indexOf('69') > -1 || texto.indexOf('420') > -1 ) {
          SPsheet.toast("Nice");
        }
        for(let k=1; k<sheets.length; k++) {
          if(k != 2) {
            Logger.log("i="+i+"; j="+j+"; k="+k+"; "+sheets[k].getRange(j, i).getValue());
            if(sheets[k].getRange(j, i).getValue() == "") {
              SPsheet.toast("Faltan datos en "+sheets[k].getName(), "¡ALERTA!");
              k=sheets.length;
              j=cursos+1;
              i=cursos+1;
            }
          }
          else {
            if(sheets[2].getRange(j,2).getValue() == "" || sheets[2].getRange(j,3).getValue() == "") {
              SPsheet.toast("Faltan datos en "+sheets[k].getName(), "¡ALERTA!");
              k=sheets.length;
              j=cursos+1;
              i=cursos+1;
            }
          }
        }
      }
    }
  }
}