// Создание листа "log", если его нет.
function createTableLog(sheets){
  
  const sheetLog = sheets.insertSheet("log");
  
  sheetLog.getRange(1, 1).setValue("Ошибка");
  sheetLog.getRange(1, 2).setValue("Дата");
  sheetLog.deleteRows(2, sheetLog.getMaxRows()-1);
  sheetLog.deleteColumns(3, sheetLog.getMaxColumns()-2);
  
  return sheetLog;
}

// Вписывание ошибки на лист "log"
function writeError(errorStr){
  const sheets = SpreadsheetApp.openById(tableId);
  let sheetLog = sheets.getSheetByName("log");
  if (!sheetLog) sheetLog = createTableLog(sheets);
  
  const row = sheetLog.getMaxRows() + 1;
  const today = Utilities.formatDate(new Date(new Date() - oneDay), "GMT+3", "dd.MM.yyyy hh:mm:ss");
  sheetLog.getRange(row, 1).setValue(errorStr);
  sheetLog.getRange(row, 2).setValue(today);
}