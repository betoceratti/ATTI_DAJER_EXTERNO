

//FUNCION PARA CONECTAR HTML MEDIANTE LA CREACION DE UN MODAL PARA EL LOGIN


function formLogin(){
 
  var fileweb = "ATTI_LOGIN"  
  var html = HtmlService.createTemplateFromFile(fileweb);
  var pagina = html.evaluate();
  
  var modal = SpreadsheetApp.getUi();
    
  modal.showModalDialog(pagina, 'A T T I SYSTEM ');
  // modal.showSidebar(html);
  
};



//FUNCION PARA ACCESAR MEDIANTE USUARIO Y PASSWORD

function login(usuario,password){
  
   var usuario = usuario;
   var password = password;  
   var message = "Pago de sistema PENDIENTE!, vencimiento 15/mes.";
    // var usuario = Browser.inputBox('Captura tu usuario');
    //var password =Browser.inputBox('Captura tu contraseña');

    if ( usuario == 'CERATTI' & password == 2702) {
    let notificacion ='Hola' + " " + usuario + " " + "Bienvenido ya tienes acceso al menu! ";
    notify(notificacion);
    crearMenu();
    //borrarLista();

    }  else {
    Browser.msgBox('Datos incorrectos,si olvidaste tus claves contacta al admin del sistema o verifica tu captura (sensible a mayusculas y minusculas)');
    var libro =SpreadsheetApp.getActive();
    var hoja =libro.setActiveSheet(libro.getSheetByName('Inicio'),true);

    hoja.getRange('B19').activate();
    }

 
};  //termina funcion
  

/*else if(usuario == 'Adbr' & password == 1306){
let notificacion ='HOLA' + " " + "DARIO BARRIOS" + " " + "YA TIENES ACCESO AL MENU ";
notify(notificacion);

var libro =SpreadsheetApp.getActive();
var hoja = libro.setActiveSheet(libro.getSheetByName('DASHBOARD DIGITAL'),true);
hoja.getRange('k1').setValue(usuario);
hoja.getRange('k1').activate().setFontFamily('Comfortaa').setFontSize(14).setFontColor('white');
menuDario();
/*Browser.msgBox("CERATTI TEC" + " " + " Le recordamos que esta pendiente el pago por el servicio del sistema, vencimiento 15/cada mes ");
MAIL = GmailApp.sendEmail('robertoceratti@gmail.com', "Alerta seguriy CERATTI-PYTHON ","EL USUARIO : " + usuario + " HA ACCESADO AL SISTEMA ");
}*/ 



//MODAL DE GASTOS BOOTSTRAP Y JAVA SCRIPT

function formGastos(){
    var hojaGastos = sicaDajer.getSheetByName('GC');
    var listado = hojaGastos.getRange('b152:b179').getValues();
    var gadmon = hojaGastos.getRange('B59:B73').getValues();
    var gventas = hojaGastos.getRange('f59:f70').getValues(); 
    var ahorro = hojaGastos.getRange('j59:j70').getValues();  
    //Browser.msgBox("Pago De Sistema SICA DAJER ++PENDIENTE++ ");
        
    var fileweb = "GASTOS"  
    var html = HtmlService.createTemplateFromFile('gastosForm'); 
      //html.listado = listado;
      
  
   const pagina = html.evaluate();  
         pagina.setHeight(350).setWidth(400);

      var modal = SpreadsheetApp.getUi();  
      modal.showModalDialog(pagina, 'Control Gastos ');
    //modal.showSidebar(pagina);
    // modal.showSidebar(html);
  
};


//funcion para obtener la data que llenaria los selec (opciones desplegables)

function getListas(){
  var hojaListas = attiSystem.getSheetByName('LOGIC');

   var listas = hojaListas.getRange('p1:r17').getValues();
   listas.shift();

   //console.log(listas);
   return listas;
};




//FUNCION PARA PROCESAR Y GAURADR OS GASTOS 


function gastosContable(data){

   var gastos = sicaDajer.getSheetByName('GC');
  var lastRow = gastos.getRange('q2').getValue(); 
  
  var gasto =[[data.fecha,data.importe,data.categoria,data.concepto,data.nota]];
  //Logger.log(categoria);
  //Logger.log(concepto);
  //var rd = gastos.getRange(lastRow + 1,18, 1,5);
  //var rd = gastos.getRange('B29:f29');
  //var fila = gastos.getRange(29, 2, 1, 5);
  
   var fila= gastos.getRange(lastRow + 1,18, 1,5);
    fila.setValues(gasto);
    //Browser.msgBox('Gasto registrado con exito'); 
    
     

};


//Funcion para alta de clientes Dajer

function altaCredito(){
  
 var formulario = attiSystem.getSheetByName('CFORM'); 
 
 var BaseData =  sicaDajer.getSheetByName('KYC');
 var ultimafila = BaseData.getRange('AL2').getValue(); 

  //hoja de plantilla
  
  var hoja = sicaDajer.getActiveSheet();
  var sheetPlantilla = sicaDajer.getSheetByName('PLANTILLA_EMAIL');  
  var lastRow = sheetPlantilla.getRange('f1').getValue(); 

  
  var rangoOrigen = formulario.getRange('A41:X41').getValues();
  var rangoDestino = BaseData.getRange(ultimafila +1, 4,1,24);
  var nombre = formulario.getRange('d4').getValue();
  var monto = formulario.getRange('k10').getValue();
  var plazo = formulario.getRange('k14').getValue();
  
  
  //Logger.log(rangoDestino);
  //Logger.log(ultimafila);
  //Para insertar una fila completa tambien podemos guaradr todas las variebles en un arreglo de arreglos
   var cliente =[[nombre,monto,hoy]]
  
  if(nombre != "" && monto>0){
    rangoDestino.setValues(rangoOrigen);
    var limpiar1 = formulario.getRange('D4:D24').clearContent();
    var limpiar2 = formulario.getRange('K4:K17').clearContent();
    var limpiar3 = formulario.getRange('K19:K20').clearContent();
    var limpiar4 = formulario.getRange('K23:K29').clearContent();
    //sheetPlantilla.getRange(lastRow,2,1,3).setValues(cliente);
    //envioEmail();
   // var pegarDatos = formulario.getRange(1,4,1,3).setValues(cliente);
  
    Browser.msgBox('Alta exitosa !'+ ' ' + ' Cliente :'+ ' '+ nombre);
    var  libro = SpreadsheetApp.getUi();
    var respuesta =libro.alert('Deseas agregar a otro cliente ?', libro.ButtonSet.YES_NO);
    
      if(respuesta == 'YES'){
    
      formulario.getRange('D4').activate();
      
      } else if (respuesta == 'NO'){
       formulario.hideSheet();
       
       var redireccionar =attiSystem.setActiveSheet(attiSystem.getSheetByName('Inicio'),true);
      
      }           
    
     }else{
       
       Browser.msgBox('Revisa tu formulario, al parecer los datos estan incompletos');
       formulario.getRange('D4').activate();
    
    } 

};



//FUNCION PARA LLAMAR A LA CEDULA DE EL CASO NUEVO PARA CREAR CEDULA NUEVA O RENOVACION
function crearCedula(){


 const hojaCliente = sicaDajer.getSheetByName('KYC');
 var ultimaFila = hojaCliente.getRange('al2').getValue();
 var cliente = hojaCliente.getRange("d"+ ultimaFila).getValue();
 var credito = hojaCliente.getRange("AB"+ ultimaFila).getValue();
 var nombre = cliente +"_"+ credito;
 var  hojaBuscada = cliente;

 let hojas = sicaDajer.getSheets();
 hojas.forEach((hoja,index) => {
   if(hoja.getName() == hojaBuscada){
     hoja.showSheet();
     var hojaactiva =libro.setActiveSheet(libro.getSheetByName(hojaBuscada),true);
    //console.log('hoja encontrada' + hoja.getName() + "N HOJA :"+ index);
   }
 }); 
 
 

 //hojaactiva
 /*var hojaActiva = libro.getSheetByName(cliente);
 var lastRow = hojaActiva.getLastRow();
 var filasAcopiar = hojaActiva.getRange("m" + lastRow).getValue();
 var inicio = lastRow - filasAcopiar + 1;
 


 var rangoCopiar = hojaActiva.getRange(inicio,1,filasAcopiar,13);
 var rangoDestino = hojaActiva.getRange(lastRow+5,1,filasAcopiar,13);*/

 //rangoCopiar.copyTo(rangoDestino);

 /*console.log(hojaActiva.getName());
 console.log(cliente);
 console.log(filasAcopiar);
 console.log(lastRow); 
 console.log(inicio);*/


};




//funcion para insertar los valores pro considerando la anotacion osea un rango de celdas de otra hoja

function conectarCedula(){



  //KYC HOJA DATOS
  const hojaCliente = sicaDajer.getSheetByName('KYC');
  var ultimaFila = hojaCliente.getRange('al2').getValue();
  var clienteAnterior =  hojaCliente.getRange("d"+ (ultimaFila-1)).getValue();
  var cliente = hojaCliente.getRange("d"+ ultimaFila).getValue(); 

  //HOJA ACTIVA , QUE DEBE SER LA CEDULA DE EL CLIENTE ABIERTA
  var hojaActiva = sicaDajer.getActiveSheet();
  var nombreHoja = hojaActiva.getName();


  if(cliente == nombreHoja | clienteAnterior == nombreHoja){

      const lastRow = hojaActiva.getLastRow();
   
      var nCredito = hojaActiva.getRange('k'+ (lastRow -1)).getValue();  
      
      var nombre = nombreHoja +"_"+nCredito;
    

      var lista1 = hojaActiva.getRange(lastRow-1,1,1,1 ).getA1Notation();
      var lista2 = hojaActiva.getRange(lastRow-1,2,1,1 ).getA1Notation();
      var lista3 = hojaActiva.getRange(lastRow-1,3,1,1 ).getA1Notation();
      var lista4 = hojaActiva.getRange(lastRow-1,4,1,1 ).getA1Notation();
      var lista5 = hojaActiva.getRange(lastRow-1,5,1,1 ).getA1Notation();
      var lista6 = hojaActiva.getRange(lastRow-1,6,1,1 ).getA1Notation();
      var lista7 = hojaActiva.getRange(lastRow-1,7,1,1 ).getA1Notation();
      
      var lista8 = hojaActiva.getRange(lastRow-1,8,1,1 ).getA1Notation();
      var lista9 = hojaActiva.getRange(lastRow-1,9,1,1 ).getA1Notation();
      var lista10= hojaActiva.getRange(lastRow-1,10,1,1 ).getA1Notation();
      var lista11= hojaActiva.getRange(lastRow-1,11,1,1 ).getA1Notation();

      var lista12 = hojaActiva.getRange(lastRow,1,1,1 ).getA1Notation();
      var lista13 = hojaActiva.getRange(lastRow,2,1,1 ).getA1Notation();
      var lista14 = hojaActiva.getRange(lastRow,3,1,1 ).getA1Notation();
      var lista15 =  hojaActiva.getRange(lastRow,4,1,1 ).getA1Notation();
      var lista16 = hojaActiva.getRange(lastRow,5,1,1 ).getA1Notation();

      var dato1=  hojaActiva.getRange(lastRow-4,1,1,1 ).getA1Notation();
      var dato2=  hojaActiva.getRange(lastRow-3,1,1,1 ).getA1Notation();
      var dato3=  hojaActiva.getRange(lastRow-2,1,1,1 ).getA1Notation();  
      

      //SICDJR HOJA DE DATOS CARTERA SE REGISTRAN LOS PAGOS DE CREDITOS
      const sicdjr = sicaDajer.getSheetByName('SICDJR');
      const ultimaFila = sicdjr.getRange('A1').getValue();
      
      sicdjr.getRange(ultimaFila+1,10,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + lista1);
      sicdjr.getRange(ultimaFila+1,11,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista2);
      sicdjr.getRange(ultimaFila+1,12,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista3);
      sicdjr.getRange(ultimaFila+1,13,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + lista4);
      sicdjr.getRange(ultimaFila+1,14,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista5);
      sicdjr.getRange(ultimaFila+1,15,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista6);
      sicdjr.getRange(ultimaFila+1,16,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + lista7);

      sicdjr.getRange(ultimaFila+1,19,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista8);
      sicdjr.getRange(ultimaFila+1,20,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista9);
      sicdjr.getRange(ultimaFila+1,21,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + lista10);
      sicdjr.getRange(ultimaFila+1,22,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista11);

      sicdjr.getRange(ultimaFila+1,25,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista12);
      sicdjr.getRange(ultimaFila+1,26,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + lista13);
      sicdjr.getRange(ultimaFila+1,27,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista14);
      sicdjr.getRange(ultimaFila+1,28,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista15);
      sicdjr.getRange(ultimaFila+1,29,1,1).setValue("="+"'"+ nombreHoja+"'"+"!" + lista16);
    
      

      
      
      //HOJA DE CONTROL DE PAGOS DEL MES EN CURSO , SE CONTABILIZA A BALANCE
      var control = sicaDajer.getSheetByName(mesActual);
      var lastColumn = control.getRange('a1').getValue() + 3;
      var multiplo = control.getRange(39,lastColumn+1).getA1Notation();

      
      control.getRange(1,lastColumn + 1,1,1).setValue(nombre);
      control.getRange(40,lastColumn+1,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + dato1 +"*"+multiplo);
      control.getRange(41,lastColumn+1,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + dato2 +"*"+multiplo);
      control.getRange(42,lastColumn+1,1,1).setValue("="+"'"+ nombreHoja+"'" +"!" + dato3 +"*"+multiplo);

      
      //console.log(nombre + "/"+ lista1 + "/" + lista2 + "/" + lista3);
      //console.log(nombre + "/"+ dato1 + "/" + dato2 + "/" + dato3);
      //console.log("sicdjr ultima fila " + ultimaFila + "control ultima columna " + lastColumn);
    
      notify("¡La conexion del credito a nombre de : " + nombreHoja + " con n de Credito " + nCredito + " fue Exitosa!");
   

     }else{

       notify("¡La cedula activa , no corresponde a el caso a dar de alta, verifique el proceso!");

     }

    
   

};





//FUNCION PARA ENVIO DE CORREOS DESDE SHEETS
function envioEmail(){
  
  var sheet = sicaDajer.getSheetByName('PLANTILLA_EMAIL');  
  var lastRow = sheet.getRange('f1').getValue();
  
  var destinatario = sheet.getRange(lastRow,1).getValue();
  var nombre = sheet.getRange(lastRow,2).getValue();
  var monto = sheet.getRange(lastRow, 3).getValue();
  var plazo = sheet.getRange(lastRow,4).getValue();
  var alta = sheet.getRange(lastRow,5).getValue();  
  var plantilla = sheet.getRange('f5').getValue();
  var asunto = "Alta de caso a nombre de " + nombre;

  //en caso de envio de muchos obtener todos os contactos
  var contactos = sheet.getRange(372, 1, 2, 5).getValues();
  
  
  //PARA ENVIO DE MAILS A VARIOS DESTINATARIOS AL MISMO TIEMPO USAMOS CLICLO FOREACH
  //Logger.log(contactos);
  /*contactos.forEach(function(fila){
    Logger.log(fila[0]);
    GmailApp.sendEmail(fila[0], 'prueba','body')
  })*/
  
  
  //Con este manera el cuerpo del correo se va sin formato , todo pegado asi que haremos una plantilla, lo mismo podriamos hacer coin asunto 
  //var body = "Alta de credito a nombre de : " + nombre + "con fecha " + alta +"por un monto de :  "+ monto + " comentarios : " + comentarios;
  //plantilla con replace
  var body = plantilla.replace('{nombre}', nombre).replace('{fecha}', alta).replace('{monto}', monto).replace('{plazo}',plazo);
 
  //Logger.log(body);
  if(nombre != "" && monto != ""){
     var email = GmailApp.sendEmail(destinatario, asunto,body)
     Browser.msgBox('Email enviado con exito"');
  
  
  }else{
   Browser.msgBox('El mail no se ha enviado,hubo un error.');
  
  } 
 
};







//FUNCION PARA EDITAR EL CAMPO AHORRO

//MODAL PARA AHORRO
function ahorroForm(){
  var html =HtmlService.createHtmlOutputFromFile('AhorroUpdate');
  var modal = SpreadsheetApp.getUi();
      modal.showModalDialog(html, 'A T T I'); 


};

function actualizar(importe,nota){
  var hoja = sicaDajer.getSheetByName('CARTERA VIGENTE');
  var kyc = sicaDajer.getSheetByName('KYC');
  var sicdjr = sicaDajer.getSheetByName('SICDJR');
  var importe = importe;
  var nota = nota;
  var cliente = hoja.getRange('f62').getValue();
  var credito = hoja.getRange('d62').getValue();
  var filakyc = hoja.getRange('E62').getValue();
  var filasicdjr = hoja.getRange('E63').getValue();
  var columnakyc = hoja.getRange('H62').getValue();
  var columnasicdjr =hoja.getRange('H63').getValue();
  //Logger.log(sicdjr);
  
  //RANGOS DESTINO
  var kycdestino = kyc.getRange(filakyc, columnakyc);
  var sicdjrdestino = sicdjr.getRange(filasicdjr,columnasicdjr);
  //insertamos los valores
    kycdestino.setValue(importe);
    kycdestino.setNote(nota);
    sicdjrdestino.setValue(importe);
    sicdjrdestino.setNote(nota);

    rentabilidad(cliente)

};




