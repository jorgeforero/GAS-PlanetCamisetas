// Email
// Planet Camisetas

/**
* sendEmail2Client
* Envía el correo con la cotización al cliente
* @param {string} message - mensaje a incluir en el correo
* @param {string} email - correo electrónico del cliente
* @param {string} subject - asunto del correo
* @param {blob} file - archivo (cotización) a incluir en el correo
**/
function sendEmail2Client( Parameters ) {

  // Carga el template HTML para el correo
  var tpl_email = HtmlService.createHtmlOutputFromFile( 'Email/tpl_email.html' ).getContent();
  tpl_email = tpl_email.replace( '##EMAILCONTENT##', Parameters.message );
  Parameters.options.htmlBody = tpl_email;
  
  // Envío del correo
  GmailApp.sendEmail( Parameters.email, Parameters.subject, '', Parameters.options );
  
};
