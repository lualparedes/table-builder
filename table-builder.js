function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Create table', 'showSidebar')
    .addToUi();
}


function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Table builder');
  DocumentApp.getUi().showSidebar(ui);
}


function createTableArr(rows, cols) {
  
  var tableArr = [];

  for (var r = rows; r > 0; r--) {    
    var rowArr = [];
    for (var c = cols; c > 0; c--) {
      rowArr.push('');
    }
    tableArr.push(rowArr);
  }

  return tableArr;
}


function insertTable(rows, cols, type) {

  var rows = rows > 0 ? rows : 1;
  var cols = cols > 0 ? cols : 1;
  var body = DocumentApp.getActiveDocument().getBody();

  var tableArr = createTableArr(rows, cols);
  
  return body.appendTable(tableArr);
}