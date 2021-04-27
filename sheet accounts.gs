// Возвращает счета с листа "счета"
function getAccountsOnSheet(sheetAccounts){
  const data = sheetAccounts.getRange(1, 2, 1, sheetAccounts.getMaxColumns()).getValues();
  const accounts = [];
  for (let i = 0; i < data[0].length; i++){
    if(data[0][i] !="") accounts.push(data[0][i]);
  }
  return accounts;
}

// Возвращает остаток на начало дня, путём вычитания суммы выписок за этот день из остатка на данный момент
function getLastValue(payments, account){
  let today = new Date();                          
  let lastValue = 0;
  
  let strToday = Utilities.formatDate(today, "GMT+3", "yyyy-MM-dd") + "T00:00:00";
  
  for (let i = payments.length-1; i >= 0; i--){
    if (payments[i].executed == strToday){
      if (payments[i].bankAccountNumber == account) {
        if (payments[i].category == "Debet")  lastValue += payments[i].amount;
        if (payments[i].category == "Credit")  lastValue -= payments[i].amount;
      }
    }
    else break;
  }
  return lastValue;
};

// Обновляет значения счета на листе "счета"
function updateAccount(sheetAccounts, payments, amount, row, column){
  sheetAccounts.getRange(row, column+1).setValue(Number(amount));
  const lastValue = getLastValue(payments, sheetAccounts.getRange(1, column).getValue());
  sheetAccounts.getRange(row, column).setValue(Number(amount) - Number(lastValue));
  sheetAccounts.getRange(row-1, column+1).setNote("");
  sheetAccounts.getRange(row-1, column+1).setBackgroundRGB(255, 255, 255);
  sheetAccounts.getRange(row, column).setNote("");
  sheetAccounts.getRange(row, column).setBackgroundRGB(255, 255, 255);
}

// Возвращает строку с нужной датой
function getRowAccounts(sheetAccounts){
  const today = Utilities.formatDate(new Date(), "GMT+3", "dd.MM.yyyy");
  
  if (sheetAccounts.getRange(sheetAccounts.getMaxRows(), 1).getValue() == today) return sheetAccounts.getMaxRows();
  
  sheetAccounts.getRange(sheetAccounts.getMaxRows() + 1, 1).setValue(today);
  return sheetAccounts.getMaxRows();
}

// Сравнивает значение остатка на начало дня с значением остатка конца прошлого дня
function checkValuesOfAccounts(sheetAccounts, row, column){
  const yesterday = Utilities.formatDate(new Date(new Date() - oneDay), "GMT+3", "dd.MM.yyyy");
  if (sheetAccounts.getRange(row-1, 1).getValue() == yesterday){
    const lastValue = sheetAccounts.getRange(row-1, column+1).getValue();
    const value = sheetAccounts.getRange(row, column).getValue();
    if (lastValue != value && lastValue !== ""){
      sheetAccounts.getRange(row-1, column+1).setNote("Не совпадает со следующим днём");
      sheetAccounts.getRange(row-1, column+1).setBackgroundRGB(255, 0, 0);
      sheetAccounts.getRange(row, column).setNote("Не совпадает с прошлым днём");
      sheetAccounts.getRange(row, column).setValue("");
      sheetAccounts.getRange(row, column).setBackgroundRGB(255, 0, 0);
      writeError('На листе "счета" не совпадают данные. Cчет: ' + sheetAccounts.getRange(1, column).getValue());
      Browser.msgBox("Ошибка!", 'На листе "счета" не совпадают данные. Cчет: ' + sheetAccounts.getRange(1, column).getValue(), Browser.Buttons.OK);
    }
  }
}

//Обновляет значения на листе "счета"
function updateAccountsValues(sheets, payments, companies){
  const sheetAccounts = sheets.getSheetByName("счета");
  
  const row = getRowAccounts(sheetAccounts);
  
  const accounts = getAccountsOnSheet(sheetAccounts);
  
  for (let i = 0; i < companies.length; i++){
    for (let j = 0; j < companies[i].bankAccounts.length; j++){
      let is_exist = false;
      let column;
      for (let k = 0; k < accounts.length; k++){
        if (companies[i].bankAccounts[j].number == accounts[k]){
          column = (k+1)*2;
          updateAccount(sheetAccounts, payments, companies[i].bankAccounts[j].balance, row, column);
          is_exist = true;
          break;
        }
      }
      if (!is_exist){
        column = sheetAccounts.getMaxColumns()+1;
        sheetAccounts.getRange(1, column, 1, 2).setValue(companies[i].bankAccounts[j].number).merge().setBorder(false, false, true, true, false, false);
        updateAccount(sheetAccounts, payments, companies[i].bankAccounts[j].balance, row, column);
      }
      checkValuesOfAccounts(sheetAccounts, row, column);
    }
  }
}