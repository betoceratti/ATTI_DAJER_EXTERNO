//modal regitro usuarios con clase

function registroModal(){

  var html = HtmlService.createTemplateFromFile('RegistroUsers');
  var pagina = html.evaluate();
  pagina.setWidth(900).setHeight(350);

  var modal = SpreadsheetApp.getUi();
  modal.showModalDialog(pagina,"A T T I");

};


//FUNCION PARA CREAR OBJETOS CLASES 

function crearUsuario(form){

    class User{
      constructor(nombre,usuario,email,password){
        this.nombre = nombre;
        this.usuario = usuario;
        this.email = email;
        this.password = password;
      }
    
    }

    /*let name = Browser.inputBox('Captura un nombre');
    let user = Browser.inputBox('Captura tu usuario ');
    let mail = Browser.inputBox('Captura un correo ');
    let key1 = Browser.inputBox('Captura una contraseña');
    let key2 = Browser.inputBox('Confirma tu contraseña');*/

    let name = form.name;
    let user = form.user;
    let mail = form.email;
    let key1 = form.password1;
    let key2 = form.password2;

    if(key1 == key2){
      colaborador = new User(name,user,key1,mail);
    guardar();
     Browser.msgBox('¡Usuario :' + user + " " + 'registrado exitosamente!')
     formLogin(); 
    }else{
      Browser.msgBox('¡Las contraseñas no coinciden, verifica tu captura!')
    }

    
    //console.log(cliente);

};//termina la funcion


function guardar(){

    const clientesDb = [[colaborador.nombre,colaborador.usuario,colaborador.email,colaborador.password]];

    
    let sheet = attiSystem.getSheetByName('Users');
    let lastRow = sheet.getLastRow();
    //sheet.appendRow(clientesDb);
    let rangoDestino = sheet.getRange(lastRow +1,1,1,4);
    rangoDestino.setValues(clientesDb);
   // notify("Operacion exitosa!!!!")

};


function updatePassword(){
  let usuario = Browser.inputBox('Captura tu usuario');  
  let password = Browser.inputBox('Captura tu password actual');

  let hojaUsuarios = attiSystem.getSheetByName('Users');
  let lastRow = hojaUsuarios.getRange('e1').getValue();
  let data = hojaUsuarios.getRange(2,1,lastRow,3).getValues();
  //console.log(data);
  let nombres = data.map(nombre =>  nombre[1]);
  //console.log(nombres);
  let indice = nombres.indexOf(usuario);
  //let id = data[indice][0];

  for(i=1; i<= lastRow; i++){
      if(hojaUsuarios.getRange(i,2).getValue() == usuario & hojaUsuarios.getRange(i,3).getValue() ==     password){
      Browser.msgBox("Hola" + " " + usuario + "Ahora ´puedes cambiar tu contraseña")
      let passwordNuevo = Browser.inputBox('Captura nueva contraseña');
      hojaUsuarios.getRange(i,3).setValue(passwordNuevo);
      Browser.msgBox("Tu password se ha actualizado con exito")
      }else{
         Browser.msgBox("Error en las claves, verifica tu captura..") 
      }

        }

};


//LOGIN PARA BUSCAR EN LA BASE DE USUARIOS Y DESPLEGAR ALGUNA ACCION
function loginSheet(){
      let usuario = Browser.inputBox('Captura tu usuario');
      let password = Browser.inputBox('Captura tu password');
      //var usuario = usuario;
      //var password = password;

      var users = attiSystem.getSheetByName('Users');
      let data = users.getDataRange().getValues();
      let lastRow = users.getLastRow();

  
      for(i=1; i<= lastRow; i++){
        if(users.getRange(i,2).getValue() == usuario & users.getRange(i,3).getValue() ==  password){
         var indice = i;        
        Browser.msgBox("Hola" + " " + usuario + " - " + indice);


        }
      

      }//for 
      
     
};



//FUNCION PARA PROTEGER HOJAS , CON RANGOS EXCEPTOS 

function protejerHoja(hojaProtegida){
    var sheet = attiSystem.getSheetByName(hojaProtegida.getName());  
    var rangosNoprotegidos = sheet.getRangeList(["b2:b10","d2:d10","f2:f10"]).getRanges();
    //console.log(rangosNoprotegidos.length)

    var proteccion = sheet.protect().setDescription('Solo Administrador').setUnprotectedRanges(rangosNoprotegidos);

    proteccion.removeEditors(proteccion.getEditors());
    if(proteccion.canDomainEdit()){
      proteccion.setDomainEdit(false);

    }

};

//muchas hojas
function protejerHojas(){
  const rangodeHojas =attiSystem.getSheets();

  rangodeHojas.forEach(hojaProtegida =>{
    if(hojaProtegida.getName() == "HH" || hojaProtegida.getName() == "Hoja 14") {
      protejerHoja(hojaProtegida);

    }
  
    
  })



};


//FUNCION PARA OBTENER EDITORES

function getEditores(){
  hojas = attiSystem.getSheets();
  editores = attiSystem.getEditors();
  propietario = attiSystem.getOwner().getEmail();

  editores.forEach(editor =>{

   console.log(editor.getEmail())
 })
  
  

};


//CICLOS O LOOPS

function mientras(){
  var contador = 0;

  while(contador < 9){
    console.log("hola")
    contador ++;
  }

};


//ciclo for normal

var objeto = {
    unArray: new Array(10000)
}

function badPerformance() {
    console.time("bad");

    for(let i=0; i< objeto.unArray.length; i++){
        objeto.unArray[i] = "hola";
    }

    console.timeEnd("bad");
}



//ciclo for optimizado

function goodPerformance() {
    console.time("good");

    let optimizado = objeto.unArray.length;

    for(let i=0; i< optimizado; i++){
        objeto.unArray[i] = "hola";
    }

    console.timeEnd("good");
};

//FUNCIONES ASINCRONAS UTILIZADOS COMUNMENTE EN USO DE APIS CON PROMESAS ASINK AWAIT THEN ECT

const api = [{

    nombre: "THE SKY",
    genero: "TERROR",
    año: 2022

},{

    nombre: "THE COVE",
    genero: "MISTERIO",
    año: 2021

},{

    nombre: "THE MOON",
    genero: "DOCUMENTAL",
    año: 2019

}];


//Un manera de usar las promesas pormise usamos timeout para imitar una peticion al servidor

const  getDtos = ()=>{

    return new Promise((resolve,reject)=>{

        setTimeout(() => {
            resolve(api);
        }, 1500);

    })

    
}



//impirmimos la respuetas de un promise de esta manera

//getDtos().then((api)=> console.log(api));

//con funcion asincrona 
async function esperar(params) {

    try{
        const respuesta = await getDtos();
        console.log(respuesta);

    }catch(error){

        console.log(error.message);

    }



    
};

// llamariamos ala funcion esperar();

//funciones matematicas 

function aleatorio(){
  var min = 10;
  var max = 21;

  for(i=0;i<10;i++){
    console.log(Math.round( Math.random() * 10 + min));
  }

  
};



