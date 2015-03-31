var ss = SpreadsheetApp.getActiveSheet();

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu( 'Sketchfab' )
    .addItem( 'Upload', 'showUploadSidebar' )
    .addToUi();
}

function showUploadSidebar() {
  var style = HtmlService.createHtmlOutputFromFile( 'stylesheet' ),
      script = HtmlService.createHtmlOutputFromFile( 'javascript' ),
      html = HtmlService.createHtmlOutputFromFile( 'upload' )
    .setSandboxMode( HtmlService.SandboxMode.IFRAME )
    .setTitle( 'Upload' )
    .setWidth( 300 )
    .append( style.getContent() )
    .append( script.getContent() );

  SpreadsheetApp.getUi()
    .showSidebar( html );
}

function populateSheet( modelsArr, rowCount, columnCount ) {
  Logger.log( 'model count: ' + ( rowCount - 1 ) );
  
  // Ranges
  var fullRange = ss.getRange( 1, 1, rowCount, columnCount ),
    licRange = ss.getRange( 'G2:G' + rowCount ),
    catRange = ss.getRange( 'H2:H' + rowCount ),
    
    // Download
    licSlugs = [
      'CC Attribution',
      'CC Attribution-ShareAlike',
      'CC Attribution-NoDerivs',
      'CC Attribution-NonCommercial',
      'CC Attribution-NonCommercial-ShareAlike',
      'CC Attribution-NonCommercial-NoDerivs',
      'CC0 Public Domain'
    ],
    licRule,
    
    // Categories
    cats = getCategories(),
    catSlugs = [],
    catRule;

  // Populate sheet
  fullRange.setValues( modelsArr );

  // Build data validation rules
  for ( var row = 2; row < rowCount; row++ ) {
    var nameRange = ss.getRange( 'B' + row ),
      desRange = ss.getRange( 'C' + row ),
      pwRange = ss.getRange( 'F' + row ),

      nameRule = SpreadsheetApp.newDataValidation()
        .requireFormulaSatisfied( '=LTE(LEN(B' + row + '), 48)' )
        .setHelpText( 'Model name must be 48 characters or less.' )
        .build(),
      
      desRule = SpreadsheetApp.newDataValidation()
        .requireFormulaSatisfied( '=LTE(LEN(C' + row + '), 1024)' )
        .setHelpText( 'Model description must be 256 characters or less.' )
        .build(),

      pwRule = SpreadsheetApp.newDataValidation()
        .requireFormulaSatisfied( '=LTE(LEN(F' + row + '), 64)' )
        .setHelpText( 'Model password must be 64 characters or less.' )
        .build();

    nameRange.setDataValidation( nameRule );
    desRange.setDataValidation( desRule );
    pwRange.setDataValidation( pwRule );
  }

  for ( var cat = 0; cat < cats.length; cat++ ) {
    catSlugs.push( cats[ cat ].slug );
  }

  licRule = SpreadsheetApp.newDataValidation()
    .requireValueInList( licSlugs, true )
    .build();
  licRange.setDataValidation( licRule );

  catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList( catSlugs.sort(), true )
    .build();
  catRange.setDataValidation( catRule );

  // Pass categories back to client
  return cats;
}

function getMetaData() {
  var range = ss.getRange( 1, 1, ss.getMaxRows(), ss.getMaxColumns() ),
    values = range.getValues();
  Logger.log( 'getMetaData ran: ');
  Logger.log( values );
  return values;
}

function clearSheet() {
  var range = ss.getRange( 1, 1, ss.getMaxRows(), ss.getMaxColumns() );
  range.clear().clearDataValidations();
}

function getCategories() {
  var catUrl = 'https://api.sketchfab.com/v2/categories',
    catResponse = UrlFetchApp.fetch( catUrl );
  
  return JSON.parse( catResponse.getContentText() ).results;
}