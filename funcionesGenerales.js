
// @ts-nocheck
//VARIABLES GLOBALES
//ATAJO PARA BUSCAR LINEAS CTRL + G EN ARCHIVOS HTML O GS

const attiSystem = SpreadsheetApp.getActiveSpreadsheet();
//const sicaDajer = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1DqDnGp41cHBUz7J9gjiHR93xxtCsl03m-wHcnmZ51uI/edit#gid=1564341209');
const sicaDajer = SpreadsheetApp.openById('1DqDnGp41cHBUz7J9gjiHR93xxtCsl03m-wHcnmZ51uI');
var hojaActiva = sicaDajer.getActiveSheet();
var hojaActivaatti = attiSystem.getActiveSheet();
var mesActual = "CONTROL PAGOS JUNIO";
const month = "JUNIO"
var hoy = fechaHoy();
var hora = getHora();
var sheetLogica = attiSystem.getSheetByName('LOGIC');
/*https://script.google.com/home/projects/1IrXyLB7ejox7f96CzNZ3kCzu6gBS0wVR67m0K0wXTZsjkd8uHep_owVt/edit*/
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++





//FUNCION PARA PODER OBTENER LA FORMULA DE UN ACELDA Y PODER EDIRALA DESDE ESTE SISTEMA
function getFormula(){
  let columna = Browser.inputBox('Captura la columna ');
  let fila = parseInt( Browser.inputBox('Captura la fila de la semana'));
  let nPagos=  parseInt( Browser.inputBox('Captura la cantidad de pagos a aplicar'));
 
  let control = sicaDajer.getSheetByName(mesActual);
  let nombre = control.getRange(columna + 1).getValue();
  let formula = control.getRange(columna + 39).getFormulas();

  let semana1 = control.getRange(columna + 44).getFormulas();
  let semana2 = control.getRange(columna + 45).getFormulas();
  let semana3 = control.getRange(columna + 46).getFormulas();
  let semana4 = control.getRange(columna + 47).getFormulas();
  Browser.msgBox(nombre +"/ la formula es :" + formula);

  var pregunta = SpreadsheetApp.getUi();
  var respuesta = pregunta.alert("Deseas aplicar pagos adelantados?",pregunta.ButtonSet.YES_NO);

  if(respuesta == 'YES'){
    let cellDestino = control.getRange(columna + 39).setValue(formula + "+" + nPagos);
    if(fila == 44){
       control.getRange(columna + fila).setValue(semana1  + "+" + nPagos);
       notify("!Proceso exitoso¡");

    }else if(fila == 45){
      control.getRange(columna + fila).setValue(semana2  + "+" + nPagos);
      notify("!Proceso exitoso¡");
    }else if(fila == 46){
         control.getRange(columna + fila).setValue(semana3  + "+" + nPagos);
         notify("!Proceso exitoso¡");
    }else if(fila == 47){
        control.getRange(columna + fila).setValue(semana4  + "+" + nPagos);
         notify("!Proceso exitoso¡");
    }

    
  }else{
    notify('Proceso cancelado..');

  }
  
 
};


//FUNCION PARA TENER LA FECHA DE EL DIA CON UN FORMATO PERSONALIZADO
function fechaHoy(){
    let dia = new Date().getDate();
    let mes = new Date().getMonth()+1;
    let year = new Date().getFullYear();
    let fecha = dia + "/" + mes + "/" + year
  //console.log(fecha);   
  return fecha

};



//GETHORA

function getHora(){
  let fecha = new Date();
  let hora = fecha.getHours();
  let minutos = fecha.getMinutes();


  var momento = hora + ":" + minutos;

  return momento;



};

//FUNCION PARA OBTENER EL USURIO ACTIVO

function getUser(e){

    let mail = Session.getActiveUser().getEmail();

    let usuario = Session.getUser().getUsername();
    console.log(usuario);




};

//FUNCION PRA OBTENER TODAS LAS HOJAS DE DAJER 

function dajerSistem(){

  var hojas = sicaDajer.getSheets();
  var id = sicaDajer.getId();

  //console.log(sicaDajer.getName(),sicaDajer.getNumSheets());

  var sheets = [];

  var infoDajer = attiSystem.getSheetByName('LOGIC');
  //console.log(id);

 hojas.forEach((hoja,index) => {
  //console.log(hoja.getName(),index,hoja.getSheetId());

  sheets = [hoja.getName(),hoja.getSheetId(),hoja.getIndex()];
 // console.log(sheets);
  
  infoDajer.appendRow(sheets);
  
//infoDajer.getRange(100,1,59,3).setValues(hoja.getName(),hoja.getSheetId(),hoja.getIndex());

});

};


//FUNCION PARA CERRAR HOJAS EN SICA DAJER
function cerrarhojasSica(){
  var hojas = sicaDajer.getSheets();
  var dashboard = sicaDajer.getSheetByName('DASHBOARD DIGITAL').hideSheet();

  /*  hojas.forEach(sheet => {
      if(sheet.getIndex() != 4 && sheet.getIndex() != 25){
      //console.log(sheet.getName()+ sheet.getIndex());
       sheet.hideSheet();
  }});*/

    notify("Se han cerrado las hojas de SICADAJER");

};


function cerrarHoja(){
  let hojaClose = attiSystem.getActiveSheet().hideSheet();

};


function llamarLogic(){
  attiSystem.setActiveSheet(attiSystem.getSheetByName('LOGIC'),true);

};

function crearMenu(){
    var menu = SpreadsheetApp.getUi().createMenu('Menu Dajer');
    var menu1 = SpreadsheetApp.getUi().createMenu('Reportes Contables');
    var menu2 = SpreadsheetApp.getUi().createMenu('Funciones Admin');
    var menu3 = SpreadsheetApp.getUi().createMenu('Cartera Credito');

    menu1.addItem('Detalle_Recuperacion', 'contables')
    .addItem('Detalle_Gastos', 'reporteGastos')
    .addItem('Detalle_Ahorro', 'clientesAhorro')
    .addItem('Kpis', 'kipisModal')
    .addItem('Detalle_flujo_caja', 'cashflow');

    menu2.addItem('Nuestro Equipo','team')
    .addItem('Prestamos Colaboradores','colabModal')
    .addItem('Aplicar Pagos en Cedula Cliente','buscarCliente')
    .addItem(mesActual,'pagosModal')
    .addItem("Registro de Gastos","formGastos")
    .addItem("Limpiar Lista",'borrarLista')
    .addItem("cerrar SICA",'cerrarhojasSica')
    .addItem("Pagos Adelantados",'getFormula')
    .addItem('Cotizador','cotizadorForm')
    .addItem('Regitro Usuarios','registroModal')
    .addItem('Modificar Contraseña','updatePassword');


    menu3.addItem('Operaciones de hoy! ','exigibles')
    .addItem('C_Vigente ', 'vigentes')
    .addItem('C_Vencida ', 'vencidos')
    .addItem("Casos Recientes","casosRecientes")
    .addItem("Tabla Vencimientos","tablaVencimientos")
    .addItem('Rentabilidad x Cliente','rentModal');



    menu.addItem('Finanzas.', 'indicadores')
    menu.addSeparator();
    menu.addSubMenu(menu2);
    menu.addSeparator();
    menu.addSubMenu(menu3);
    menu.addSeparator();
    menu.addSubMenu(menu1);
    menu.addToUi();



};


//FUNCION PARA QUITAR O PONER PROTECCION AL ARCHIVO
function agregarEditor(){
  let hojaProtejida = sicaDajer.getSheetByName('LISTAS DINAMICAS');

  hojaProtejida.protect().addEditor('ceratti.web@gmail.com');

};



//FUNCION PARA EMITIR NOTIDICACIONES TOAST 
function notify(mensaje) {
   
  SpreadsheetApp.getActive ().toast(mensaje,"A T T I SYSTEM");
  
};



//ELIMINA LA LISTA DE CLEDAS EDITADAS HAY Q PONERLE UN TRIGGER PARA LO HAGA DIARIO 
function borrarLista(){

  const hojaLista = sicaDajer.getSheetByName('LISTAS DINAMICAS');
  const fila = hojaLista.getRange('af1').getValue();
  const rango = hojaLista.getRange('ae2:ae20').clearContent();
  notify('Se eliminaron los datos de el dia de ayer..');

};


//FUNCION PAR CAPTURAR CELDAS EDITADAS

function capturarCeldas(){


  const hojaLista = sicaDajer.getSheetByName('LISTAS DINAMICAS');
  const hojaActivaSica = sicaDajer.getActiveSheet();
  const fila = hojaLista.getRange('af1').getValue();


  let nombreHoja = hojaActivaSica.getName();
  let celdaActiva = hojaActivaSica.getActiveCell();
  let valor = hojaActiva.getActiveCell().getValue();
  let ubicacion =  hojaActivaSica.getActiveCell().getRow();
  let fecha = new Date().getDate() + "/" + (new Date().getMonth() +1);


  if(celdaActiva.getRow() > 1 & celdaActiva.getColumn() > 1 & valor == true & hojaActivaSica.getIndex() > 20){

  hojaLista.getRange(fila+1,31,1,1).setValue("Se clickeo un pago  de : " + nombreHoja + " en la fila  :" +  ubicacion + " con fecha :" + fecha + "a las : " + getHora());

  //notify("Se edito la hoja con nombre  : " + nombreHoja);

  }

};



//FUNCION PARA SABER LA UBICACION EN EL ARCHIVO DE UN USUARIO
function buscarUbicacion(){
    let nameSheet = hojaActivaatti.getName();  
    let user = sicaDajer.getEditors()[1].getEmail();

    let celActive = hojaActiva.getActiveCell();
    let row = celActive.getRow();
    let column = celActive.getColumn();
    

    const ubicacion = "El usuario esta en la hoja " + nameSheet + " -en la fila " + row + "- y columna  " + column;
    notify(ubicacion);
    //console.log(ubicacion);

};


//FUNCION PARA DAR UN AVISO CUANDO SE REALIZE UN PRIMER PAGO
function aviso(){

  var hojaAviso = attiSystem.getSheetByName('Inicio');
  var contador = hojaAviso.getRange('a1').getValue();

  if(contador > 1){
    
    notify("Se hizo un pago !!");
  }
  //console.log(contador);

};


//FUNCION PARA BUSCAR DASTOS DE ACUERDO A UN VALOR

function buscarV(){
    const hojaBusqueda = attiSystem.getSheetByName('LOGIC');
    const filas = hojaBusqueda.getRange('f219').getValue();
    rangoBusqueda = hojaBusqueda.getRange(220,2,filas,4).getValues();

    var nombres = rangoBusqueda.map(nombre => nombre[0]);



    const  hojaForm = attiSystem.getSheetByName('CFORM');
    const valorBuscado = hojaForm.getRange('D4').getValue();
    const celActive = hojaForm.getActiveCell();
    const nombre = hojaForm.getRange('d4').getValue();
    const v1 = hojaForm.getRange('d6');
    const v2 = hojaForm.getRange('d8');
    const v3 = hojaForm.getRange('d10');
    const indice = nombres.indexOf(valorBuscado);
    //console.log(indice);


    if(indice !== -1 & valorBuscado !== "" & hojaActivaatti.getName() == hojaForm.getName() & celActive.getRow()>3 & celActive.getColumn()>3){
      var cliente = rangoBusqueda[indice][0];
      var rfc = rangoBusqueda[indice][1];
      var curp = rangoBusqueda[indice][2];
      var tel = rangoBusqueda[indice][3];

      v1.setValue(rfc);
      v2.setValue(curp);
      v3.setValue(tel);
      //notify("Busqueda exitosa");
    }else{
      v1.clearContent();
      v2.clearContent();
      v3.clearContent();
      notify("Cliente nuevo; capture todos los campos");
    }


};


//funcion para insertar los valores pro considerando la anotacion osea un rango de celdas de otra hoja

function conectarCedula(){
 
  const hojaDestino = attiSystem.getSheetByName('COLUMNAS');

  const lastRow = hojaActivaatti.getLastRow();

  var lista1 = hojaActivaatti.getRange(lastRow,1,1,5 ).getA1Notation();
  var lista2 =  hojaActivaatti.getRange(lastRow,7,1,2 ).getA1Notation();
  var dato1=  hojaActivaatti.getRange(lastRow-3,1,1,1 ).getA1Notation();
   var dato2=  hojaActivaatti.getRange(lastRow-2,1,1,1 ).getA1Notation();
    var dato3=  hojaActivaatti.getRange(lastRow-1,1,1,1 ).getA1Notation();
  var nombreHoja = hojaActivaatti.getName();
  var multiplo = hojaDestino.getRange(100-4,1).getA1Notation();
  const celdaDestino = hojaDestino.getRange(100,1,1,5).setValue("="+ nombreHoja +"!" + lista1);
  const celdaDestino2 = hojaDestino.getRange(100,7,1,2).setValue("="+ nombreHoja +"!" + lista2);
  const celdaDestino3 = hojaDestino.getRange(100-3,1,1,1).setValue("="+ nombreHoja +"!" + dato1 +"*"+multiplo);
  notify("PROCESO EXITOSO");
  //console.log(lista3);

};







