
//FUNCION PARA DESPLEGAR UN HTML CON LLENADO EXIGIBLES
// <i class="fas fa-print"></i> <i class="fas fa-file-pdf"></i>
  


function exigibles(){

  var hojaData = sicaDajer.getSheetByName('LISTAS DINAMICAS');
  var hojaGastos = sicaDajer.getSheetByName('GC');
  var hojaColab = sicaDajer.getSheetByName('PRESTAMOS COLABORADORES');
  var hojaExigibles = attiSystem.getSheetByName('SICDJR');
  var hojaBalance = attiSystem.getSheetByName('CHART');


  //var rangoData = hojaData.getRange('A6:C20').getValues();  
  var diferencia = hojaBalance.getRange('c1').getDisplayValue();
  var filaInicio = hojaData.getRange('x1').getDisplayValue();
  var nFilas= hojaExigibles.getRange('AL6').getDisplayValue();
  var filasAltas = hojaExigibles.getRange('ap6').getValue();
  var filas = hojaData.getRange('af1').getValue();
  var rangoData =hojaExigibles.getRange(7, 35 ,nFilas, 4).getDisplayValues();
  var rangoAltas = hojaExigibles.getRange(7,39,filasAltas,4).getDisplayValues();
  let suma = rangoData.reduce((suma,monto)=> suma + monto[1],0);  
  //console.log(suma);

  
  //ALGORITMO PARA PAGOS HOY
  var hojaOrigen = sicaDajer.getSheetByName(mesActual);
  var fI = hojaOrigen.getRange('c199').getDisplayValue();
  var nF = hojaOrigen.getRange('b199').getDisplayValue();
  var  pagosHoy =hojaOrigen.getRange(200,2,nF,2).getValues();
  //Logger.log(pagosHoy);
  let = nPagos = pagosHoy.reduce((contar,elemento)=>  contar + 1,0);
  //console.log(rangoData);

  //SECCION ÁRA VISUALIZARA EDICIONES DE CELDAS CLICKEO

  var clickeos = hojaData.getRange(2,31,filas,1).getValues();
  //console.log(clickeos);

  //SECCION PARA CAPTURAR LOS GASTOS DE EL DIA
  var nRows = hojaGastos.getRange('ak1').getValue();
 var gastosHoy = hojaGastos.getRange(2,36,nRows,4).getDisplayValues();

  //SECCION PRA CAPTURAR OPERACION DEL DIA DE COLABORADORES
  var nRowscolab = hojaColab.getRange('aa1').getValue();
 var colabHoy = hojaColab.getRange(2,26,nRowscolab,4).getDisplayValues();


  var totales = sicaDajer.getSheetByName('TC');
  var totalExigibles = totales.getRange('E13').getDisplayValue();
  var numero = totales.getRange('H22').getDisplayValue();
   //var rangoData = [nombre,pago];
  //Logger.log(rangoData);
    var plantilla = HtmlService.createTemplateFromFile("Exigibles");
    //plantilla.nombres = nombres;
    //plantilla.ids = ids;
     plantilla.rangoData = rangoData;
     plantilla.totalExigibles = totalExigibles;
     
     plantilla.numero = numero;
     plantilla.pagosHoy = pagosHoy;
     plantilla.clickeos = clickeos;
     plantilla.gastosHoy = gastosHoy;
     plantilla.colabHoy = colabHoy;
     plantilla.rangoAltas = rangoAltas;
     plantilla.diferencia = diferencia;
  
    const pagina = plantilla.evaluate();
    pagina.setWidth(580).setHeight(400);
  
   const ui = SpreadsheetApp.getUi();
   ui.showModalDialog(pagina, "A T T I")

   /* var alerta = SpreadsheetApp.getUi();
      var respuesta =  alerta.alert("Deseas descargar en PDF ?",alerta.ButtonSet.YES_NO);

       if(respuesta == 'YES'){       
         
        descargarPdf(pagina); 
     
      
      } else{        

        var plantilla = HtmlService.createTemplateFromFile("Exigibles");
   
            plantilla.rangoData = rangoData;
            plantilla.totalExigibles = totalExigibles;
             plantilla.pagosHoy = pagosHoy;
            plantilla.numero = numero;
          
        const pagina = plantilla.evaluate();
        pagina.setWidth(500).setHeight(400);
        
        const ui = SpreadsheetApp.getUi();
        ui.showModalDialog(pagina, "A T T I")
                  
      
      }     */           
  

};  
  

//FUNCION PARA OBTENER LA DATA QUE DESPLEGAREMOS EN EL MODAL HTML CASOS RECIEN OTORGADOS

function casosRecientes(){
  

  let fechaMax = new Date ("November 30, 2021");
  let dateMin = new Date("November 01, 2021");
  let porFecha = rango => rango[16] > dateMin && rango[16] < fechaMax;//filter
  let porYear = rango => rango[16] > dateMin;//filter

    var hojaData = sicaDajer.getSheetByName('KYC');
  
  var lastRow = hojaData.getRange('AL2').getDisplayValue();
 
  var rangoData = hojaData.getRange(lastRow -9,4, 10,26).getDisplayValues();

  var filtroxNombre = rangoData.filter(filtrado => filtrado[0] == "SALVADOR LOPEZ SANCHEZ");
  let rangoxFecha = rangoData.filter(mes => mes[16] == "18/10/2021");
  
  //let nCasos = rangoData.reduce((total,casos)=> total + 1,0);
  //Logger.log(rangoData[0][16] );
  //console.log(dateMin);  
  
  let mes = new Date().getMonth() +1;

  var plantilla = HtmlService.createTemplateFromFile("Recientes");
  plantilla.rangoData = rangoData;
  plantilla.ultimaFila = lastRow;
  plantilla.filtroxNombre = filtroxNombre;
  plantilla.rangoxFecha = rangoxFecha;
  plantilla.mes = mes;

  const pagina = plantilla.evaluate();
  pagina.setWidth(950).setHeight(500);
  
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
  
 
  
};


//FUNCION PARA OBTENER LA DATA QUE DESPLEGAREMOS EN EL MODAL HTML TABLA DE VENCIMIENTOS

function tablaVencimientos(){
  
  var hojaData = sicaDajer.getSheetByName('VENCIMIENTOS');
  var total = hojaData.getRange('C13').getDisplayValue();
  var lastRow = hojaData.getRange('AK1').getDisplayValue();
 
  var rangoData = hojaData.getRange(2,38, lastRow,9).getValues();
  var calendario = hojaData.getRange(3,1,13,6).getDisplayValues();
  var mes = hojaData.getRange('A17').getValue();
  
  //let nCasos = rangoData.reduce((total,casos)=> total + 1,0);
  Logger.log(calendario);
  
  var dia = new Date().getDate();  
  //for (var i = 0; i <rangoData.length; i++)
   // Logger.log(rangoData[i]);

  var plantilla = HtmlService.createTemplateFromFile("Vencimientos");
  plantilla.rangoData = rangoData;
  plantilla.calendario = calendario;
  plantilla.ultimaFila = lastRow;
  plantilla.total = total;
  plantilla.mes = mes;
  plantilla.dia = dia;
  
  
  const pagina = plantilla.evaluate();
  pagina.setWidth(950).setHeight(500);
  
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
  
 
  
};

//FUNCION PARA DESPLEGAR LA PALNTILLA DE INDICADORES
//DASHBOARD
function indicadores(){

  var hoja =sicaDajer.getSheetByName('TC');
 
  var login = sicaDajer.getSheetByName('LOGIN');
  
  var porcentajevencido= hoja.getRange(13,10).getValue()*100;
  var montovencido = hoja.getRange('h13').getDisplayValue();
  var caja= hoja.getRange(13,11).getDisplayValue();
   var dif=login.getRange(1,11).getDisplayValue();
  var cashflow = login.getRange('K4').getDisplayValue();
  var cv=hoja.getRange(9,5).getDisplayValue();
  var vhoy=hoja.getRange(13,5).getDisplayValue();
   var clvi=hoja.getRange(17, 14).getDisplayValue();
  /*var clve=RP.getRange(2, 8).getDisplayValue();*/
  var ahorro = hoja.getRange('K9').getDisplayValue();
  var bancos = hoja.getRange('k17').getDisplayValue();
  
 
  var nclientes=login.getRange(1,13).getDisplayValue();
  
  //Logger.log(rangoData);
  let  page =  "Indicadores";
    
    
  var html =HtmlService.createTemplateFromFile(page);
    html.caja = caja;
    html.porcentajevencido = porcentajevencido.toFixed(0);
    html.cv = cv;
    html.clvi = clvi;
    html.dif = dif;
    html.cashflow = cashflow;
    html.montovencido = montovencido;
    html.ahorro = ahorro;
    html.bancos = bancos;
   // html.prueba = prueba;
  const pagina = html.evaluate();
        pagina.setHeight(600).setWidth(700);
    
    
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')  
    

};


//FUNCION PARA SELECCIONAR EN HOJA VIGENTES EL N DE CREDITO PARA ACTUAÑIZARA AHORO
function selectCredito(numero){
  const hojaCliente = sicaDajer.getSheetByName('CARTERA VIGENTE');
  const campo = hojaCliente.getRange('d62');

  var nCredito = numero;

  campo.setValue(nCredito);

 ahorroForm()

};




//MODAL PARA RENTABILIDAD X CLIENTE
function rentModal(){

  const hojaLista = sicaDajer.getSheetByName('KYC');
  const ultimaFila = hojaLista.getRange('AQ1').getValue();

  const listado = hojaLista.getRange('AN' + 2 + ':AN'+ ultimaFila).getDisplayValues();

  //console.log(listado); 

  var html =HtmlService.createTemplateFromFile('buscadorRentabilidad');
      html.listado = listado; 
      const pagina =html.evaluate();
          pagina.setHeight(200).setWidth(370);
          var modal = SpreadsheetApp.getUi();
              modal.showModalDialog(pagina, 'A T T I');

};





//HTML RENTABILIDAD
function rentabilidad(cliente){

  const hojaCliente = sicaDajer.getSheetByName('CARTERA VIGENTE');
  var nombreCliente = hojaCliente.getRange('B41'); 
  //var cliente = hojaCliente.getRange('A3' ).getValue();
  var cliente = cliente;

  nombreCliente.setValue(cliente);

  var lastRow = hojaCliente.getRange('C41').getValue();
  var  rangoData = hojaCliente.getRange(45,2,lastRow,11).getDisplayValues();
  const evaluacion = hojaCliente.getRange('F39').getDisplayValue();
  const totales = hojaCliente.getRange('b43:L43').getDisplayValues();
  const mensaje = hojaCliente.getRange('h40').getDisplayValue();
  const totalRecup = hojaCliente.getRange('h42').getDisplayValue();


  var html =HtmlService.createTemplateFromFile('Rentabilidad');
        
      html.rangoData = rangoData; 
      html.cliente = cliente;
      html.evaluacion = evaluacion;
      html.totales = totales;
      html.mensaje = mensaje;
      html.totalRecup = totalRecup;
      
    const pagina =html.evaluate();
            pagina.setHeight(500).setWidth(850);
      
    var modal = SpreadsheetApp.getUi();
        modal.showModalDialog(pagina, 'A T T I');
    
     /* var alerta = SpreadsheetApp.getUi();
        var respuesta =  alerta.alert("Deseas descargar en PDF ?",alerta.ButtonSet.YES_NO);

        if(respuesta == 'YES'){       
          
          descargarPdf(pagina); 
      
        
        } else{        
          
          var html =HtmlService.createTemplateFromFile('Rentabilidad');
        
            html.rangoData = rangoData; 
            html.cliente = cliente;
            html.evaluacion = evaluacion;
            html.totales = totales;
            html.mensaje = mensaje;
            
            const pagina =html.evaluate();
                  pagina.setHeight(500).setWidth(850);
            var modal = SpreadsheetApp.getUi();
                modal.showModalDialog(pagina, 'A T T I');      
                      
      }     */           
  
};




//MODAL PARA PRESTAMO COLABORADORES
function colabModal(){

        
        const hojaLista = sicaDajer.getSheetByName('PRESTAMOS COLABORADORES');
        const sheet = sicaDajer.getSheetByName('LOGIN');
        const ultimaFila = hojaLista.getRange('A1').getValue();

        const listado = hojaLista.getRange(ultimaFila-9,1,10,7).getDisplayValues();
        const cabeceras = hojaLista.getRange(1,1,1,6).getDisplayValues();
        var saldoAaron = hojaLista.getRange('e1').getDisplayValue();
        var saldoRob = hojaLista.getRange('f1').getDisplayValue();
        var saldoJes = hojaLista.getRange('g1').getDisplayValue();

        //console.log(saldoAaron); 

        var html =HtmlService.createTemplateFromFile('Colaboradores');

              html.listado = listado; 
              html.cabeceras = cabeceras;
              html.saldoAaron = saldoAaron;
              html.saldoRob = saldoRob;
              html.saldoJes = saldoJes;

              var pagina =html.evaluate();
                  pagina.setHeight(500).setWidth(850);
                  var modal = SpreadsheetApp.getUi();
                      modal.showModalDialog(pagina, 'A T T I');
   

    /*  var alerta = SpreadsheetApp.getUi();
      var respuesta =  alerta.alert("Deseas descargar en PDF ?",alerta.ButtonSet.YES_NO);

       if(respuesta == 'YES'){       
         
        descargarPdf(pagina); 
     
      
      } else{        

              var html =HtmlService.createTemplateFromFile('Colaboradores');

              html.listado = listado; 
              html.cabeceras = cabeceras;
              html.saldoAaron = saldoAaron;
              html.saldoRob = saldoRob;

              var pagina =html.evaluate();
                  pagina.setHeight(500).setWidth(810);
              var modal = SpreadsheetApp.getUi();
                  modal.showModalDialog(pagina, 'A T T I');
      
      }        */        

};


//MODAL PARA BISQUEDA DE CEDULA
function buscarCliente(){

  const hojaLista = sicaDajer.getSheetByName('LISTAS DINAMICAS');
  const ultimaFila = hojaLista.getRange('L54').getValue();

  //const listado = hojaLista.getRange('B' + 6 + ':B'+ ultimaFila).getDisplayValues();
  const listado = hojaLista.getRange( 54,13,ultimaFila,1).getDisplayValues();

  //console.log(listado); 

  var html =HtmlService.createTemplateFromFile('buscarCedula');
      html.listado = listado; 
      const pagina =html.evaluate();
          pagina.setHeight(200).setWidth(370);
          var modal = SpreadsheetApp.getUi();
              modal.showModalDialog(pagina, 'A T T I');



};

//ESTADOS DE CUENTA O CEDULA DE CLIENTE 
function edodeCuenta(cliente){

  var cliente = cliente;
  var  hojaBuscada = cliente;
  let hojas = sicaDajer.getSheets();

  hojas.forEach((hoja,index) => {
    
    if(hoja.getName() == hojaBuscada){
    var  lastRow = hoja.getLastRow();
    var nFilas = hoja.getRange('l'+ (lastRow-1)).getValue();
    var filas = hoja.getRange('n'+lastRow).getValue()-1;
    var calendario = hoja.getRange(lastRow-filas,4,nFilas,4).getDisplayValues();
    var vigente = hoja.getRange('B'+ (lastRow-1)).getDisplayValue();
    var vencido = hoja.getRange('C'+ (lastRow-1)).getDisplayValue();
    var pagosVencidos = hoja.getRange('A'+ (lastRow-1)).getDisplayValue();
    var nCredito = hoja.getRange('k'+ (lastRow -1)).getValue();
    var montoInicial = hoja.getRange(lastRow-(filas-1),3,1,1).getDisplayValue();
    var fechaInicial = hoja.getRange(lastRow-(filas-2),3,1,1).getDisplayValue();
    var pago = hoja.getRange(lastRow-(filas-7),3,1,1).getDisplayValue();
    var hoy = fechaHoy();

    console.log(montoInicial + "" + fechaInicial);

    var html =HtmlService.createTemplateFromFile('cedulaCliente');
       
        html.calendario = calendario;  
        html.cliente = cliente;
        html.vigente = vigente;
        html.vencido = vencido;
        html.pagosVencidos = pagosVencidos;
        html.nCredito = nCredito;
        html.montoInicial = montoInicial;
        html.fechaInicial = fechaInicial;
        html.pago = pago;
        html.hoy = hoy;


  const pagina =html.evaluate();
          pagina.setHeight(500).setWidth(550);
    
  var modal = SpreadsheetApp.getUi();
      modal.showModalDialog(pagina, 'A T T I');
   
   
   /*  */ 
    
   }  
  
 }); 



};



//FUNCION PARA AP0LICAR PAGOS SIN ABRIR CEDULAS GENIAL

function pagarCedula(cliente,importe){

    const hojaLista = sicaDajer.getSheetByName('LISTAS DINAMICAS');
    const fila = hojaLista.getRange('af1').getValue();

    var cliente = cliente;
    var importe = importe;
    var  hojaBuscada = cliente;
    let hojas = sicaDajer.getSheets();

   hojas.forEach((hoja,index) => {
  
    if(hoja.getName() == hojaBuscada){
    var  lastRow = hoja.getLastRow();
    var nFilas = hoja.getRange('l'+ (lastRow-1)).getValue();
    var filas = hoja.getRange('n'+lastRow).getValue()-1;
    var calendario = hoja.getRange(lastRow-filas,4,nFilas,4).getDisplayValues();
    var vigente = hoja.getRange('B'+ (lastRow-1)).getDisplayValue();
    var vencido = hoja.getRange('C'+ (lastRow-1)).getDisplayValue();
    var pagosVencidos = hoja.getRange('A'+ (lastRow-1)).getDisplayValue();
    var nCredito = hoja.getRange('k'+ (lastRow -1)).getValue();
    var montoInicial = hoja.getRange(lastRow-(filas-1),3,1,1).getDisplayValue();
    var fechaInicial = hoja.getRange(lastRow-(filas-2),3,1,1).getDisplayValue();
    var pagados = hoja.getRange('m'+ (lastRow -1)).getValue();
    var proximoPago = hoja.getRange(lastRow-(filas-pagados),4,1,1).setValue(true);
    var moratorios = hoja.getRange(lastRow-(filas-pagados),11,1,1).setValue(importe); 
    let ubicacion =  proximoPago.getRow();
    
  
    
   }  
  
  hojaLista.getRange(fila+1,31,1,1).setValue("Se edito la celda : " + " Caso " + cliente + "con fecha : " + hoy);

 }); 


};


function pagosModalCedula(cliente){


  const hojaLista = sicaDajer.getSheetByName(mesActual);

  /*const ultimaFila = hojaLista.getRange('B138').getValue();
  const filaIinicio = hojaLista.getRange('A5').getValue();
  const listado = hojaLista.getRange(139,3,ultimaFila,1).getValues();
  const listaBusqueda = hojaLista.getRange(1,4,1,hojaLista.getLastColumn()).getValues();*/


  //console.log(listaBusqueda); 

      var html =HtmlService.createTemplateFromFile('pagosMesCedula');
      html.cliente = cliente; 
      html.mesActual = mesActual;
      const pagina =html.evaluate();
          pagina.setHeight(300).setWidth(380);
          var modal = SpreadsheetApp.getUi();
              modal.showModalDialog(pagina, 'A T T I');

};






//MODAL PARA APLICAR PAGOS EN CNTROL DE PAGOS DEL MES
function pagosModal(){


  const hojaLista = sicaDajer.getSheetByName(mesActual);

  const ultimaFila = hojaLista.getRange('B138').getValue();
  const filaIinicio = hojaLista.getRange('A5').getValue();
  const listado = hojaLista.getRange(139,3,ultimaFila,1).getValues();
  const listaBusqueda = hojaLista.getRange(1,4,1,hojaLista.getLastColumn()).getValues();


  //console.log(listaBusqueda); 

      var html =HtmlService.createTemplateFromFile('pagosMes');
      html.listado = listado; 
      html.mesActual = mesActual;
      const pagina =html.evaluate();
          pagina.setHeight(300).setWidth(380);
          var modal = SpreadsheetApp.getUi();
              modal.showModalDialog(pagina, 'A T T I');

};



//FUNCION PARA APLICAR PAGOS EN EL CONTROL DE MES MEDIANTE FORMULARIO 

function pagos(nombre,importe,nota){

  let hojaDatos = sicaDajer.getSheetByName(mesActual);
  let hojaActiva = sicaDajer.getActiveSheet();
  var celdaActiva = hojaDatos.getActiveCell();
  var filaActiva = celdaActiva.getRow();
  var columActiva = celdaActiva.getColumn();
  //let valorBuscaddo = hojaDatos.getRange('b125').getValue();
  let valorBuscaddo = nombre;
  //let importe = parseFloat(Browser.inputBox('Captura el monto a pagar:')) ;
  let ultimaFila = hojaDatos.getRange('c122').getValue() + 1;
  let listaBusqueda = hojaDatos.getRange(1,4,1,hojaDatos.getLastColumn()).getValues();

  //console.log(listaBusqueda[0]);
  //console.log(valorBuscaddo);


  listaBusqueda.forEach(cliente => {

  if(mesActual == hojaDatos.getName() && valorBuscaddo != ""){
  var indice = cliente.indexOf(valorBuscaddo) +4;
  //console.log(indice);
  //celdaActiva = hojaDatos.getRange(ultimaFila, indice).activate();
  celdaActiva = hojaDatos.getRange(ultimaFila, indice).setValue(importe);
  var anotacion = hojaDatos.getRange(ultimaFila, indice).setNote(nota);
  var mensaje = "El pago se ha procesado correctamente a nombre de : " + valorBuscaddo;
  notify(mensaje);


  }


  }
);


};

//FUNCION PARA DAR REVERSO A PAGOS EN CONTROL DE PAGOS APLICADOS POR ERROR
function reversoPago(nombre,nota){

    let hojaDatos = sicaDajer.getSheetByName(mesActual);
    let hojaActiva = sicaDajer.getActiveSheet();
    var celdaActiva = hojaDatos.getActiveCell();
    var filaActiva = celdaActiva.getRow();
    var columActiva = celdaActiva.getColumn();
    //let valorBuscaddo = hojaDatos.getRange('b125').getValue();
    let valorBuscaddo = nombre;

    let ultimaFila = hojaDatos.getRange('c122').getValue() + 1;
    let listaBusqueda = hojaDatos.getRange(1,4,1,hojaDatos.getLastColumn()).getValues();

    //console.log(listaBusqueda[0]);
    //console.log(valorBuscaddo);
    listaBusqueda.forEach(cliente => {

    if(mesActual == hojaDatos.getName() && valorBuscaddo != ""){
    var indice = cliente.indexOf(valorBuscaddo) +4;
    //console.log(indice);
    //celdaActiva = hojaDatos.getRange(ultimaFila, indice).activate();
    celdaActiva = hojaDatos.getRange(ultimaFila, indice).clearContent();
    var anotacion = hojaDatos.getRange(ultimaFila, indice).setNote(nota);
    var mensaje = "El pago se ha eliminado correctamente a nombre de : " + valorBuscaddo;
    notify(mensaje);


    }

    })


};






//FUNCION PARA DESPLEGAR INIDICADORES CONTABLES POR AÑO
//INGRESOS Y GASTOS
function contables(){
  var hojatablero = sicaDajer.getSheetByName('TC');
  var hojaMes = sicaDajer.getSheetByName('VENCIMIENTOS');
  
  var datos = hojatablero.getRange('C'+2+':I'+5).getDisplayValues();
  var esperado =hojatablero.getRange('m10:p12').getDisplayValues();
  
  var ingresos = hojatablero.getRange('M8').getDisplayValue();
  var gastos = hojatablero.getRange('N8').getDisplayValue();
  var utilidad = hojatablero.getRange('O8').getValue();
  var hoy = hojatablero.getRange('e22').getDisplayValue();
  var mes = hojaMes.getRange('A17').getValue();
  // html.bancos = new Intl.NumberFormat().format(bancos);
  //Logger.log(datos);
  var year = datos.map(function(años){ return(años[6])});


  var html =HtmlService.createTemplateFromFile('tablaRecuperacion');
      html.datos = datos;
      html.esperado = esperado;
      html.ingresos = ingresos;
      html.gastos = gastos;
      html.utilidad = utilidad;
      html.hoy = hoy;
      html.mes = mes;
    
    //   html.utilidad = new Intl.NumberFormat().format(utilidad);
  
  const pagina =html.evaluate();
        pagina.setHeight(500).setWidth(550);
  
  var modal = SpreadsheetApp.getUi();
      modal.showModalDialog(pagina, 'A T T I');

};


//funcion para guaradr la rentabilidad real

function sR(){
  var hojaRr = attiSystem.getSheetByName('RR');
  var hojatablero = sicaDajer.getSheetByName('TC');

  var ingresos = hojatablero.getRange('M8').getDisplayValue();
  var gastos = hojatablero.getRange('N8').getDisplayValue();
  var utilidad = hojatablero.getRange('O8').getDisplayValue();
  var vigente = hojatablero.getRange('e9').getDisplayValue();
  var vencido = hojatablero.getRange('h13').getDisplayValue();
  var colocacion = hojatablero.getRange('i20').getDisplayValue();
  var ingresosContables = hojatablero.getRange('O12').getDisplayValue();

  var row = [ingresosContables,ingresos,gastos,utilidad,vigente,vencido,colocacion,new Date(),month];

  hojaRr.appendRow(row);
  //console.log(colocacion);

};


function getSr(){
    var hojaRr = attiSystem.getSheetByName('RR');

    const data = hojaRr.getDataRange().getValues();
    data.shift();

    // console.log(data);
    return data;



};


function modalSr(){
  const plantilla = HtmlService.createTemplateFromFile('htmlSr');

    plantilla.rangoData = getSr()
 

   const pagina = plantilla.evaluate();
    pagina.setWidth(750).setHeight(450);
  
   const ui = SpreadsheetApp.getUi();
   ui.showModalDialog(pagina, "A T T I")



};




//FUNCION PARA OBTENER LA DATA QUE DESPLEGAREMOS EN EL MODAL HTML CARTERA VIGENTE
//DE LOS GASTOS
function reporteGastos(){
  
  var hojaData = sicaDajer.getSheetByName('GC');
  var lastRow = hojaData.getRange('Q1').getDisplayValue();
  var rangoData = hojaData.getRange(2,27, lastRow-33,9).getValues();
  rangoData.sort(function(a, b){return b-a}); 
  //Logger.log(rangoData);
  
  //variables para filter
  let category = "GASTOS_ADMON";//GASTOS_DE_VENTA,SALIDA_AHORRO GASTOS_ADMON
  let descripcion = "SISTEMA MAN";//APOYO SOCIOS, SISTEMA MAN EJEC COBRANZA
  let fechaMax = new Date ("January 31, 2021");
  let dateMin = new Date("January 01, 2021");



  //Predicados
  let paraMap = monto => monto[0];//map
  let paraReduce = (acumulado, monto) => acumulado + monto[1];//reduce
  let paraCategory = categoria => categoria[2] == category;//filter
  let paraConcepto = concepto => concepto[3] == descripcion;//filter
  let porFecha = rango => rango[0] > dateMin & rango[0] < fechaMax;//filter
  let porYear = rango => rango[0] > dateMin;//filter
  
  let datosFila = rangoData.map(paraMap);

  let categorias = rangoData.filter(paraCategory);  
  let conceptos = rangoData.filter(paraConcepto);
  let xconceptoyFecha = conceptos.filter(porFecha);
  let xcategoriayFecha = categorias.filter(porFecha);
  let rangosFecha = rangoData.filter(porFecha);

  let tXcategoria = categorias.reduce(paraReduce,0);
  let tXconcepto = conceptos.reduce(paraReduce,0);
  let txRango = rangosFecha.reduce(paraReduce,0);
  let txrangoyconcepto = xconceptoyFecha.reduce(paraReduce,0);
  let txrangoycategoria = xcategoriayFecha.reduce(paraReduce,0);
  //let tGastos = rangoData.reduce((suma,monto )=> suma + monto[1],0);
  let tGastos = rangoData.reduce(paraReduce,0);
  
  
  let socios = 50000;
  let sistema = 24000;

  let maximo = [socios,sistema]

  
    console.log(tGastos);
  //console.log(Math.max.apply(null,maximo));



  var plantilla = HtmlService.createTemplateFromFile("reporteGastos");
  plantilla.rangoData = rangoData;
  plantilla.tGastos = tGastos;
  const pagina = plantilla.evaluate();
  pagina.setWidth(900).setHeight(500);
 
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I_SYSTEM ')
    
};



//FUNCION PARA DESPLEGAR EL DETALLE DE AHORRO DE CLIENTES 
function clientesAhorro(){
  
  var hojaData = sicaDajer.getSheetByName('LISTAS DINAMICAS');
   var saldo = sicaDajer.getSheetByName('TC');
  var filaInicio = hojaData.getRange('AB1').getDisplayValue();
  var nFilas = hojaData.getRange('AD1').getDisplayValue();
  var rangoData =hojaData.getRange(1, 28 ,nFilas, 2).getValues();
  var ahorro = saldo.getRange('k9').getDisplayValue();
  
  
  //for (var i = 0; i <rangoData.length; i++)
   // Logger.log(rangoData[i]);

  var plantilla = HtmlService.createTemplateFromFile("ahorroClientes");
  plantilla.rangoData = rangoData;
  plantilla.ahorro = ahorro;
  const pagina = plantilla.evaluate();
  pagina.setWidth(500).setHeight(500);
  
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
  //return pagina
    
};


//FUNCION PARA DESPLEGAR FLUJO DE EFECTIVO
function cashflow(){
  
  var hojaData = attiSystem.getSheetByName('LOGIC');
  

  var rangoData = hojaData.getRange('Q104:T112').getDisplayValues();
  var ingresos = hojaData.getRange('r112').getValue();
  var salidas = hojaData.getRange('t112').getValue();
  
  var hojaBancos = attiSystem.getSheetByName('EF');
  var nFilas = hojaBancos.getRange('e167').getValue();

  var movimientosBancos = hojaBancos.getRange('b173:e188').getDisplayValues();
  
  //for (var i = 0; i <rangoData.length; i++)
   // Logger.log(rangoData[i]);

  var plantilla = HtmlService.createTemplateFromFile("cashFlow");
  plantilla.rangoData = rangoData;  
  plantilla.ingresos = ingresos;
  plantilla.salidas = salidas;
  plantilla.movimientosBancos = movimientosBancos;

  const pagina = plantilla.evaluate();
  pagina.setWidth(600).setHeight(450);
  
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
  
  //return pagina;
    
};



//MODAL PARA INDICADORES
function kipisModal(){

    const hojaLista = sicaDajer.getSheetByName('PYTHON');
    const ultimaFila = hojaLista.getRange('E1').getValue();
    const hojaLogica = attiSystem.getSheetByName('SICDJR');

    const listado = hojaLista.getRange(2,1,ultimaFila,2).getDisplayValues();
    const kpis = hojaLogica.getRange('ay2:az17').getDisplayValues();

    //console.log(listado); 

   var html =HtmlService.createTemplateFromFile('Kpis');
      html.listado = listado; 
      html.kpis = kpis;
      const pagina =html.evaluate();
          pagina.setHeight(500).setWidth(610);
          var modal = SpreadsheetApp.getUi();
              modal.showModalDialog(pagina, 'A T T I');

      /*var alerta = SpreadsheetApp.getUi();
      var respuesta =  alerta.alert("Deseas descargar en PDF ?",alerta.ButtonSet.YES_NO);

       if(respuesta == 'YES'){       
         
        descargarPdf(pagina); 
     
      
      } else{        
          var html =HtmlService.createTemplateFromFile('Kpis');
              html.listado = listado; 
              html.kpis = kpis;
          const pagina =html.evaluate();
                pagina.setHeight(500).setWidth(610);
          var modal = SpreadsheetApp.getUi();
              modal.showModalDialog(pagina, 'A T T I');

              
      }     */           


};




//FUNCION PARA OBTENER LA DATA QUE DESPLEGAREMOS EN EL MODAL HTML CARTERA VIGENTE

function vigentes(){
  
  var hojaData = sicaDajer.getSheetByName('CARTERA VIGENTE');
  //var saldo = Math.trunc(hojaData.getRange('k4').getDisplayValue());
  var saldo = hojaData.getRange('k4').getDisplayValue();
  var lastRow = hojaData.getRange('C1').getDisplayValue();
  var rangoData = hojaData.getRange(6,2, lastRow,13).getDisplayValues();
  
  let nCasos = rangoData.reduce((total,casos)=> total + 1,0);
  Logger.log(nCasos);
  
  var titulo = "FUNCIONA"
  
  //for (var i = 0; i <rangoData.length; i++)
   // Logger.log(rangoData[i]);

  var plantilla = HtmlService.createTemplateFromFile("Vigentes");
  plantilla.rangoData = rangoData;
  plantilla.nCasos = nCasos;
  plantilla.saldo = saldo;
  const pagina = plantilla.evaluate();
  pagina.setWidth(950).setHeight(500);
  
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
  
 
  
};


function vencidos(){
  
  var hojaData = sicaDajer.getSheetByName('CARTERA VIGENTE');
  var lastRow = hojaData.getRange('Q1').getDisplayValue();
  var rangoData = hojaData.getRange(6,16, lastRow,29).getValues();
  //Logger.log(rangoData);
  

  let fechaMax = new Date ("Jun 30, 2021");
  let dateMin = new Date("January 01, 2021");
  let dateMin22 = new Date("January 01, 2022");


  //index 10 monto vencido
  //index 9 vigente
  let resultado = rangoData.reduce((suma,monto)=> suma + monto[10],0);//total vencido

  let filtradoxvigente = rangoData.filter(filtro => filtro[9] > 0);//total vencido filtrando vigentes

  let filtradoxfecha21 = rangoData.filter(filtro => filtro[2] > dateMin);//filtrado por rango de fechas
  let filtradoxfecha22 = rangoData.filter(filtro => filtro[2] > dateMin22);//filtrado por rango de fechas
  let nombresFiltradosxfecha = filtradoxfecha21.map(fecha => fecha[0]);//impresion de los nombres de el filtrado 
  let total21 = filtradoxfecha21.reduce((suma,monto)=> suma + monto[10],0);//suma de vencido poe el filtro de fecha 
   let total22 = filtradoxfecha22.reduce((suma,monto)=> suma + monto[10],0);//suma de vencido poe el filtro de fecha   
  let totalClientes = filtradoxfecha21.reduce((suma,monto)=> suma + 1,0);//conteo de clientes vencidos por el filtro
  
  //for (var i = 0; i <rangoData.length; i++)
   Logger.log(totalClientes);

  var plantilla = HtmlService.createTemplateFromFile("Vencidos");
  plantilla.rangoData = rangoData;
  plantilla.total21 = total21;
  plantilla.total22 = total22;
  plantilla.resultado = resultado;
  const pagina = plantilla.evaluate();
  pagina.setWidth(950).setHeight(500);
  
  
  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(pagina, 'A T T I  SYSTEM ')
  
    
};




//FUNCION PARA DESPLEGAR MODAL DE NUESTRO EQUIPO (separando en archivos css y js )
function include(filename){

  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}



function team(){
  let plantilla = HtmlService.createTemplateFromFile('teams');
  //var  url = crearPdf();
  plantilla.getDatos = getDatos();
  let web = plantilla.evaluate().setWidth(900).setHeight(400);
  let window = SpreadsheetApp.getUi();

  window.showModalDialog(web, "A T T I");


}

function getDatos(){
 
  const hoja = sicaDajer.getSheetByName('LISTAS DINAMICAS');
  //var data = hoja.getDataRange().getValues();
  var data = hoja.getRange('ar2:ax4').getValues();
  //data.shift();
  //console.log(data);

  return data;
};


//COTIZADOR EN LINEA 

function cotizadorForm(){
  var html = HtmlService.createTemplateFromFile('Cotizador');
  const pagina =html.evaluate();
          pagina.setHeight(500).setWidth(550);
    
  var modal = SpreadsheetApp.getUi();
      modal.showModalDialog(pagina, 'COTIZADOR DAJER');



};



//MODAL PARA GENERARA CONTRASEÑAS
function passwordModal(){
  var html = HtmlService.createTemplateFromFile('PasswordGenertor');
  const pagina =html.evaluate();
          pagina.setHeight(300).setWidth(350);
    
  var modal = SpreadsheetApp.getUi();
      modal.showModalDialog(pagina, 'A T T I ');

};


//FUNCION PARA OBTENER UN HTML SU CONTENIDO Y MANDARLO A PDF
function getHtml(){
  var hojaData = sicaDajer.getSheetByName('LISTAS DINAMICAS');
  var hojaGastos = sicaDajer.getSheetByName('GC');
  var hojaColab = sicaDajer.getSheetByName('PRESTAMOS COLABORADORES');
   
  var hojaBalance = attiSystem.getSheetByName('CHART');
  var hojaExigibles = attiSystem.getSheetByName('SICDJR');

  //var rangoData = hojaData.getRange('A6:C20').getValues();  
  var filaInicio = hojaData.getRange('x1').getDisplayValue();
  var nFilas= hojaExigibles.getRange('AL6').getDisplayValue();
  var filasAltas = hojaExigibles.getRange('ap6').getValue();
  var filas = hojaData.getRange('af1').getValue();
  var rangoData =hojaExigibles.getRange(7, 35 ,nFilas, 4).getValues();
  var rangoAltas = hojaExigibles.getRange(7,39,filasAltas,4).getValues();
  let suma = rangoData.reduce((suma,monto)=> suma + monto[1],0);  
  var diferencia = hojaBalance.getRange('c1').getDisplayValue();
  //console.log(suma);

  
  //ALGORITMO PARA PAGOS HOY
  var hojaOrigen = sicaDajer.getSheetByName(mesActual);
  var fI = hojaOrigen.getRange('c199').getDisplayValue();
  var nF = hojaOrigen.getRange('b199').getDisplayValue();
  var  pagosHoy =hojaOrigen.getRange(200,2,nF,2).getValues();
  //Logger.log(pagosHoy);
  let = nPagos = pagosHoy.reduce((contar,elemento)=>  contar + 1,0);
  //console.log(rangoData);

  //SECCION ÁRA VISUALIZARA EDICIONES DE CELDAS CLICKEO

  var clickeos = hojaData.getRange(2,31,filas,1).getValues();
  //console.log(clickeos);

  //SECCION PARA CAPTURAR LOS GASTOS DE EL DIA
  var nRows = hojaGastos.getRange('ak1').getValue();
  var gastosHoy = hojaGastos.getRange(2,36,nRows,4).getDisplayValues();

  //SECCION PRA CAPTURAR OPERACION DEL DIA DE COLABORADORES
  var nRowscolab = hojaColab.getRange('aa1').getValue();
  var colabHoy = hojaColab.getRange(2,26,nRowscolab,4).getDisplayValues();


  var totales = sicaDajer.getSheetByName('TC');
  var totalExigibles = totales.getRange('E13').getDisplayValue();
  var numero = totales.getRange('H22').getDisplayValue();
   //var rangoData = [nombre,pago];
  //Logger.log(rangoData);
  var plantilla = HtmlService.createTemplateFromFile("Exigibles");
    //plantilla.nombres = nombres;
    //plantilla.ids = ids;
     plantilla.rangoData = rangoData;
     plantilla.totalExigibles = totalExigibles;
     
     plantilla.numero = numero;
     plantilla.pagosHoy = pagosHoy;
     plantilla.clickeos = clickeos;
     plantilla.gastosHoy = gastosHoy;
     plantilla.colabHoy = colabHoy;
     plantilla.rangoAltas = rangoAltas;
     plantilla.diferencia = diferencia;

  const pagina = plantilla.evaluate();

  pagina.setWidth(580).setHeight(400);
  

 
  
  return pagina;

};




function descargarPdf(){
 
  
    
   const sheet = attiSystem.getSheetByName('Inicio');
 
   var pagina = getHtml();
   var file = pagina.getAs('application/pdf');


      var idCarpeta = "1RonXJs6sBDeKqODGIHH2ZjQ1RBLAFsF7";
      //llamamos nuetra carpeta y ahi guardamos el pdf 
      var carpetaMaestra = DriveApp.getFolderById(idCarpeta);
      var dia = new Date().getDate();
      var mes = new Date().getMonth()+1;
      let pdf =carpetaMaestra.createFile(file).setName("Reporte Operaciones de el dia " + dia+"/"+mes );
      var link = pdf.getUrl();

      let titulo = "Reporte Operaciones de :" + "-" + hoy;

      var formula = 'HYPERLINK("'+ link+'";"'+titulo+'")'
      sheet.getRange('E1').setFormula(formula);
      sheet.getRange('E1').activate();

};


//FUNCION PARA SUBIR ARCHIVOS PARA DRIVE IMG/PDF ETC
function imgModal(){
    var html = HtmlService.createTemplateFromFile('upFile');
    var pagina = html.evaluate().setWidth(300).setHeight(300);

    var modal = SpreadsheetApp.getUi();
    modal.showModalDialog(pagina,"A T T I");


};


//👇para esta funcion

function subir(form){
   const idCarpeta = "1RonXJs6sBDeKqODGIHH2ZjQ1RBLAFsF7";//FOLDER PDF
     const idImg = "1X62BElikPFuJS6xOulHW4i6_6gIVgq3H";//FOLDER IMG
      //llamamos nuetra carpeta y ahi guardamos el pdf 
      const carpetaMaestra = DriveApp.getFolderById(idCarpeta);  

     
     
      var archivo = carpetaMaestra.createFile( form.file );
     
      var link = archivo.getUrl();
      return link;
      

      console.log( form.text + "/" + form.file);
};



//FUNCION PARA SUBIR ARCHIVOS PARA DRIVE IMG/PDF ETC FILE-READER
function upfileModal(){
  var html = HtmlService.createTemplateFromFile('subirImg');
  var pagina = html.evaluate().setWidth(500).setHeight(500);

  var modal = SpreadsheetApp.getUi();
  modal.showModalDialog(pagina,"A T T I");


};

//👇para esta funcion


//FUNCION PARA SUBIR IMGS CON VISTA PREVIA USANDO LA LIBRERIA FILE-READER
function upload(obj){

     const idCarpeta = "1RonXJs6sBDeKqODGIHH2ZjQ1RBLAFsF7";//FOLDER PDF
     const idImg = "1X62BElikPFuJS6xOulHW4i6_6gIVgq3H";//FOLDER IMG
      //llamamos nuetra carpeta y ahi guardamos el pdf 
      const carpetaMaestra = DriveApp.getFolderById(idImg);  

      var file = Utilities.newBlob(obj.bytes, obj.mimeType, obj.filename);
     
      var archivo = carpetaMaestra.createFile( file );
     
      var link = archivo.getUrl();
      return link;

};






