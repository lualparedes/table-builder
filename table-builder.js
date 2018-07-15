////////////////////////////////////////////////////////////////////////////////
// UI
////////////////////////////////////////////////////////////////////////////////

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


////////////////////////////////////////////////////////////////////////////////
// Table builder
////////////////////////////////////////////////////////////////////////////////

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

function buildTable(rows, cols) {
  
  var tableArr = createTableArr(rows, cols);

  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursorEl = doc.getCursor().getElement();
  var parent = cursorEl.getParent();

  return body.insertTable(parent.getChildIndex(cursorEl) + 1, tableArr);
}

function styleTableHeader(header, bgColor, fgColor, bold) {

  var attrHeader = {};
      attrHeader[DocumentApp.Attribute.BOLD] = bold;
      attrHeader[DocumentApp.Attribute.FOREGROUND_COLOR] = fgColor;

  var cell;
  var c = header.getNumChildren();

  for (var i = c-1; i >= 0; i--) {
    cell = header.getChild(i);
    cell.setBackgroundColor(bgColor);     
  }

  header.setAttributes(attrHeader);
}

function styleTable(table, type) {

  var headerBgColor, headerFgColor, headerBold;
  var borderColor, fontFamily, fgColor;
  var attrTable = {};
  
  if (type === 'standard') {
    headerBgColor = '#f3f3f3';
    headerFgColor = '#434343';
    headerBold = true;
    fontFamily = 'Roboto';
    borderColor = '#b7b7b7';
    fgColor = '#434343';
  }

  if (type === 'contents') {
    headerBgColor = '#4e9dff';
    headerFgColor = '#ffffff';
    headerBold = true;
    fontFamily = 'Roboto';
    borderColor = '#3d7bc8';
    fgColor = '#434343';
  }

  if (type === 'code') {
    headerBgColor = '#f3f3f3';
    headerFgColor = '#434343';
    headerBold = false;
    fontFamily = 'Roboto Mono';
    borderColor = '#f3f3f3';
    fgColor = '#434343';
  }

  if (type === 'code-dark') {
    headerBgColor = '#666666';
    headerFgColor = '#ffffff';
    headerBold = false;
    fontFamily = 'Roboto Mono';
    borderColor = '#666666';
    fgColor = '#ffffff';
  }

  attrTable[DocumentApp.Attribute.FONT_SIZE] = 12;
  attrTable[DocumentApp.Attribute.FONT_FAMILY] = fontFamily;
  attrTable[DocumentApp.Attribute.BORDER_COLOR] = borderColor;
  attrTable[DocumentApp.Attribute.FOREGROUND_COLOR] = fgColor;
  table.setAttributes(attrTable);

  styleTableHeader(table.getRow(0), headerBgColor, headerFgColor, headerBold);
}

function insertTable(rows, cols, type) {
  var rows = rows > 0 ? rows : 1;
  var cols = cols > 0 ? cols : 1;
  var table = buildTable(rows, cols);
  styleTable(table, type);
}