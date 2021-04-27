// Удаление id запросов с листа "идентификаторы платежей" у которых дата больше today - amountDays.
function deleteOldPayments(bankName, sheetIdentifiers, column){ 
  const today = (new Date()).setHours(0, 0, 0, 0);
  
  const lastRow = sheetIdentifiers.getMaxRows();
  
  const Data = sheetIdentifiers.getRange(2, column, lastRow, 1).getValues();
  
  const rowsToDelete = [];
  
  for (let i = 2; i <= lastRow; i++){
    if (new Date(sheetIdentifiers.getRange(i, column).getNote()) < new Date(new Date(today - amountDaysForDeleteId*oneDay))) rowsToDelete.push(i);
  }
  if (rowsToDelete.length != 0) sheetIdentifiers.deleteRows(rowsToDelete[0], rowsToDelete.length);
}

// Получение последней строки на листе "идентификаторы платежей" для вписывания новых id.
function getLastRow(sheetIdentifiers, column){ 
  for(let i = 2; i <= sheetIdentifiers.getMaxRows()+1; i++){
    if (sheetIdentifiers.getRange(i, column).getValue() == "") return i;
  }
}

// Вписывание новых id на лист "идентификаторы платежей".
function setPerformedPayments(bankName, sheetIdentifiers, column, payments){ 
  const lastRow = getLastRow(sheetIdentifiers, column);
  
  const data = [];
  const notes = [];
  
  for (let i = 0; i < payments.length; i++){
    if (data[i] == undefined) {
      data.push([]);
      notes.push([]);
    }
    data[i][0] = payments[i].id;
    notes[i][0] = payments[i].executed
  }
  if (data.length > 0){
    sheetIdentifiers.getRange(lastRow, column, data.length, 1).setValues(data);
    sheetIdentifiers.getRange(lastRow, column, notes.length, 1).setNotes(notes);
  }
}

// Получение столбца (Определяем банк) на листе "идентификаторы платежей".
function getColumn(bankName, sheetIdentifiers){ 
  for (let i = 1; i <= sheetIdentifiers.getMaxColumns(); i++){
    if (sheetIdentifiers.getRange(1, i).getValue() == bankName) {
      return i;
    }
  }
  let column = sheetIdentifiers.getMaxColumns()+1;
  sheetIdentifiers.getRange(1, column).setValue(bankName);
  return column;
}

// Создание листа "идентификаторы платежей", если его нет.
function createTableIdentifiers(sheets){
  
  const sheetIdentifiers = sheets.insertSheet("идентификаторы платежей");
  
  sheetIdentifiers.getRange(1, 1).setValue("ModulBank");
  sheetIdentifiers.deleteRows(2, sheetIdentifiers.getMaxRows()-1);
  sheetIdentifiers.deleteColumns(2, sheetIdentifiers.getMaxColumns()-1);
  
  return sheetIdentifiers;
}

// Получение нужных выписок путём сравнивания id на листе "идентификаторы платежей" и id выписок полученных из запроса на сервер.
function getNotPerformedPayments(bankName, sheets, payments){ 
  let sheetIdentifiers = sheets.getSheetByName("идентификаторы платежей");
  if (!sheetIdentifiers) sheetIdentifiers = createTableIdentifiers(sheets);
  
  let column = getColumn(bankName, sheetIdentifiers);
  
  deleteOldPayments(bankName, sheetIdentifiers, column);
  let performedPayments = [];
  
  if (sheetIdentifiers.getMaxRows() > 1) performedPayments = sheetIdentifiers.getRange(2, column, sheetIdentifiers.getMaxRows()-1).getValues();
  
  let temp = [];
  
  for (let i = 0; i < payments.length; i++){
    let is_exist = false;
    for (let j = 0; j < performedPayments.length; j++){
      if (payments[i].id == performedPayments[j][0]) {
        is_exist = true;
        break;
      }
    }
    if (!is_exist) temp.push(payments[i]);
  }
  
  payments = temp;
  
  setPerformedPayments(bankName, sheetIdentifiers, column, payments);
  
  return payments;
}