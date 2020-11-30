// Planets Camisetas
// jorge.e.forero@gmail.com
// 2020

// Definiciones
// Tallas de camisetas que se están trabajando
const SIZES = [ 'XS', 'S', 'M', 'L', 'XL', 'XXL' ];

/**
* setSizes
* Crea el campo de tallas en la forma
**/
function setSizes() {

  // Obtiene la forma
  var form = FormApp.getActiveForm();
  // Adiciona la opcion de la talla
  var item = form.addListItem();
  item.setTitle( 'Seleccione la talla' );
  // Adiciona las opciones de talla
  var choices = [];
  for ( var index=0; index<SIZES.length; index++ ) {
    choices.push( item.createChoice( SIZES[ index ] ) )
  };
  item.setChoices( choices );
  
};
