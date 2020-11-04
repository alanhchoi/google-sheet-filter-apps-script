function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('TW Helper')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('TW Helper')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheets() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function (sheet) {
    return { id: sheet.getSheetId(), name: sheet.getSheetName() };
  });
}

function activateSheet(id) {
  function idMatches(sheet) {
    return sheet.getSheetId() === id;
  }
  
  const targets = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(idMatches);
  
  if (targets.length > 0) {
    targets[0].activate();
  }
}

function copyToNewSpreadsheet(sheetIds) {
  function idMatches(sheet) {
    return sheetIds.includes(sheet.getSheetId());
  }
  
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheets = activeSpreadsheet.getSheets().filter(idMatches);
  
  if (targetSheets.length === 0) {
    return;
  }
  
  const ssNew = SpreadsheetApp.create("[View] " + activeSpreadsheet.getName());
  targetSheets.forEach(function (sheet) {
    sheet.copyTo(ssNew);
  });
  
  const url = ssNew.getUrl();
  const htmlOutput = HtmlService.createHtmlOutput('<a href="{url}" target="_blank">{url}</a>'.replace(/\{url\}/g, url));
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Click to view');
  
  return ssNew.getUrl();
}
