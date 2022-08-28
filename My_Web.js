function doGet(e){

   let page = HtmlService.createTemplateFromFile('My_web');
  
    //aqio insertamos le data que queremos vicualizara o mostrar en el html

   page.getData = getData();
   let html = page.evaluate();
 
    html.addMetaTag('viewport','width=device-width, initial-scale=1')

   return html



};

function include(file){

  return HtmlService.createHtmlOutputFromFile(file).getContent();

};


function getData(){

   const hojaData = sicaDajer.getSheetByName('PYTHON')
   const data = hojaData.getRange('A2:B17').getDisplayValues();


  return data;



};
