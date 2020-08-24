// Code by Taiel Martinez 2020
// El uso de este script para fines comerciales esta prohibido.
// Registro: https://github.com/TaielMartinez/googleApiSendMail/
function sendEmails() {
  // conf //
  const develop = true;
  const production = false;
  const conf_letra_valores = "B";
  
  const letra_numero = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
  var date = Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yyyy")
  
  console.log(" --------- start --------- ");
  var conf_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
  
  // Declaracion de variables
  var hoja_respuestas_name = config(17);
  var mail_adress = config(18);
  var replace_motivo = config(19);
  var replace_asunto = config(20);
  var replace_organizacion = config(21);
  var replace_responsable = config(22);
  var replace_tipo = config(23);
  var veredicto = config(24);
  var date_enviado = config(25);
  var mail_enviado = config(27);
  var texto_enviar_mail = config(28);
  var message_aprobado_A = config(29);
  var message_aprobado_B = config(30);
  var message_aprobado_C = config(31);
  var message_aprobado = config(32);
  var message_rechazado = config(33);
  var message_falta_info = config(34);
  var subject_aprobado = config(35);
  var subject_rechazado = config(36);
  var subject_falta_info = config(37);
  // fin variables
  console.log("variables declaradas")
  
  
  
  var dataRange = getSheet(hoja_respuestas_name, "A", 2, "AG");
  var hojaFormulario = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja_respuestas_name);
  var data = dataRange.getValues();
  var index = 2;
  for (var i in data) {
    var row = data[i];
    if(row[mail_enviado] == texto_enviar_mail){
      
      var message;
      var subject; 
      var emailAddress = "taielmartinezgort@gmail.com";   //develop
      if(develop == false && production == true){
        mailAddress = row[mail_adress];
      }
      
      switch(row[veredicto]) {
        case "Aprovado A":
          message = remplazar_mensaje(message_aprobado, row).replace("<texto-clase>", message_aprobado_A);
          subject = remplazar_mensaje(subject_aprobado, row).replace("<texto-clase>", message_aprobado_A);
          break;
        case "Aprovado B":
          message = remplazar_mensaje(message_aprobado, row).replace("<texto-clase>", message_aprobado_B);
          subject = remplazar_mensaje(subject_aprobado, row).replace("<texto-clase>", message_aprobado_B);
          break;
        case "Aprovado C":
          message = remplazar_mensaje(message_aprobado, row).replace("<texto-clase>", message_aprobado_C);
          subject = remplazar_mensaje(subject_aprobado, row).replace("<texto-clase>", message_aprobado_C);
          break;
        case "Rechazado":
          message = remplazar_mensaje(message_rechazado, row);
          subject = remplazar_mensaje(subject_rechazado, row);
          break;
        case "Faltan datos":
          message = remplazar_mensaje(subject_falta_info, row);
          subject = remplazar_mensaje(subject_falta_info, row);
          break;
        case "Caso especial":
          message = false;
          subject = false;
          break;
        default:
          console.log("ERROR: veredicto no reconocido: " + row[veredicto] + " en: " + letra_numero[mail_enviado] + index);
      }
      
      
      //console.log(message);
      if(emailAddress && message && subject){
        console.log("------------------ Send Mail ------------------");
        console.log("email: " + emailAddress);
        console.log("asunto: " + subject);
        console.log("mensaje: " + message);
        MailApp.sendEmail(emailAddress, subject, message);
        console.log(letra_numero[mail_enviado]+index);
        hojaFormulario.getRange(letra_numero[mail_enviado]+index).setValue("ENVIADO");
        hojaFormulario.getRange(letra_numero[date_enviado]+index).setValue(date);
        console.log("=============================================================");
      }
    }
    index++;
  }
  
  
  
  
  
  
  function getSheet(hoja, inicio_letra, incio_numero, final_letra, final_numero){
    console.log("getSheet");
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja);
    var _sheet = "";
    if(!final_numero){
      final_numero = sheet.getLastRow();
    }
    if(hoja){
      _sheet = "'"+hoja+"'!";
    }
    console.log("hoja: "+hoja+" inicio_letra: "+inicio_letra+" incio_numero: "+incio_numero+" final_letra: "+final_letra+" final_numero: "+final_numero);
    return (sheet.getRange(_sheet+inicio_letra+incio_numero+":"+final_letra+final_numero));
  }
  
  function config(linea){
    return((conf_sheet.getRange("'CONFIG'!"+conf_letra_valores+linea)).getValues());
  }
  
  function remplazar_mensaje(texto, row){
    return texto.toString().replace("<motivo>", row[replace_motivo]).replace("<asunto>", row[replace_asunto]).replace("<organizacion>", row[replace_organizacion]).replace("<responsable>", row[replace_responsable]).replace("<tipo>", row[replace_tipo])
  }
  
}


