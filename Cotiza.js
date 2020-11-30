// Quote
// Planet Camisetas

/**
* createQuote
* Crea una nueva cotización a partir de los datos del formulario llenado y la envía por correo
* @param {void} - 
* @return {void} - Correo enviado al cliente con la cotización
**/
function createQuote() {
  
  // Abre la hoja de solicitud de Cotizaciones
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( 'Form Responses 1' );
  var record = sheet.getRange( sheet.getLastRow(), 1, 1, sheet.getLastColumn() ).getValues();
  
  // Crea un objeto a partir del registro obtenido
  var quotaData = record2Object( LABELS, record[0] );
  
  // Abre la hoja del template de cotización y registra los datos
  var quoteSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( 'tpl_quote' );
  var quoteNumber = getQuoteNumber();
  // Datos del cliente y Datos de la cotización específica ( número de camisetas, talla y precio )
  quoteSheet.getRange( 4, 3, 3, 1 ).setValues( [ [ quoteNumber ] , [ quotaData.fullname ] , [ quotaData.email ] ] );
  quoteSheet.getRange( 10, 3, 1, 3 ).setValues( [ [ quotaData.number, quotaData.size, getPrice( quotaData.number, quotaData.size ) ] ] );
  SpreadsheetApp.flush();
  
  // Genera un archivo PDF con la cotización  
  var url_base = URLBOOK.replace(/\/edit.*$/, '');
  var url_ext = '/export?exportFormat=pdf&format=pdf'   //export as pdf
  + '&gid=' + quoteSheet.getSheetId()
  + '&size=letter'      // paper size
  + '&portrait=true'    // orientation, false for landscape
  + '&fitw=true'        // fit to width, false for actual size
  + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
  + '&gridlines=false'  // hide gridlines
  + '&fzr=false';       // repeat row headers (frozen rows) on each page
  var options = {
    headers: {
      'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken(),
    }
  };
  var response = UrlFetchApp.fetch( url_base + url_ext, options );
  var pdfName = 'Quote' + '_' + quoteNumber  + '.pdf';
  var blob = response.getBlob().setName( pdfName );
  
  // Salva el pdf con la cotización en Drive
  var parents = DriveApp.getFileById( SpreadsheetApp.getActiveSpreadsheet().getId() ).getParents();
  var folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  folder.createFile( blob );
  // emailOptions.attachments = blob;
  
  // Se personaliza el mensaje del correo con el nombre del cliente
  var tpl_msg = HtmlService.createHtmlOutputFromFile( 'Planets/tpl_message.html' ).getContent();
  tpl_msg = tpl_msg.replace( '##NOMBRE##', quotaData.fullname );
  
  // Se envía el correo con el formato de Planet Camisetas y la cotización como attachment
  sendEmail2Client( {
    options: { attachments: blob },
    message: tpl_msg,
    email: quotaData.email,
    subject: 'Planet Camisetas :: Acá tenemos tu cotización' 
  } );
  
};

/**
* getPrice
* Obtiene el precio de la tabla de precios basado en el número de camisetas y la talla
* @param {number} Number - número de camisetas solicitadas 
* @param {string} Size - talla de las camisetas solicitadas
* @return {number} - El precio de la camiseta unitaria para la talla y número de camisetas silicitadas
**/
function getPrice( Number, Size ) {
  // Determina el indice de la fila según la talla
  var row = SIZES.indexOf( Size ) + 2;
  
  // Determina el indice de la columna según el número de camisetas
  if ( Number <= 100 ) {
    var column = 2; 
  } else if ( ( Number > 100 ) && ( Number <= 200 ) ) {
    column = 3; 
  } else if ( ( Number > 200 ) && ( Number <= 500 ) ) { 
    column = 4; 
  } else if ( Number > 500 ) {
    column = 5;
  };
  
  // Obtiene la hoja de precios para retornar el precio
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( 'Precios' );  
  return sheet.getRange( row, column ).getValue();  
};

/**
* getQuoteNumber
* Obtiene el número de la cotización
* @param {void} - 
* @return {void} - número único para la cotización
**/
function getQuoteNumber() {
  var size = 4;
  var key = Utilities.getUuid();
  key = key.substr( 0, size );
  for ( var index=0; index<size ; index++ ) key += String.fromCharCode( 97 + Math.random()*10 );
  return key;
};

/**
* record2Object
* Crea un objeto a partir de los labels y registro dado
* @param {array} Labels - nombres de las propiedades para el objeto
* @param {array} Record - Valores para el objeto
* @return {void} - Objeto con los datos
**/
function record2Object( Labels, Record ) {
  var resObj = {}; 
  // carga los valores en un objeto para facilitar la referenciación
  for ( var index=0; index<Record.length; index++ ) {
    resObj[ Labels[ index ] ] = Record[ index ];
  };
  return resObj;
};
