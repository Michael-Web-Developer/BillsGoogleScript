function main(){

  const sheetPayment = SpreadsheetApp.openById(tableId);
  
  const companiesModulBank = getCompaniesModulBank(); // Возвращает массив компаний из API или строку "Server is down"
  let paymentsModulBank = [];
  const companiesTochka = getCompaniesTochka();
  let dataTochka = {payments: [], companies: []};
  
  if (companiesModulBank != 503) paymentsModulBank = getObjectsModulBank(companiesModulBank); // Возвращает выписки из API
  if (companiesTochka != 503) dataTochka = getPaymentsTochka(companiesTochka.organizations);// view {payments: convertedPayments, companies: convertedCompanies}
  
  updateAccountsValues(sheetPayment, paymentsModulBank, companiesModulBank); // Обновляет остатки счетов на листе "счета"
  updateAccountsValues(sheetPayment, dataTochka.payments, dataTochka.companies);
  
  paymentsModulBank = getNotPerformedPayments("ModulBank", sheetPayment, paymentsModulBank); // Возвращает выборку выписок, которые ещё не были использованы.
  let paymentsTochka = getNotPerformedPayments("Точка", sheetPayment, dataTochka.payments);
  
  const sheet = sheetPayment.getSheetByName('выписки');
  sheet.getRange(1, 3, sheet.getMaxRows(), 1).setNumberFormat("0.00");
  sheet.getRange(1, 8, sheet.getMaxRows(), 2).setNumberFormat("@");
  
  setPayments(sheet, paymentsModulBank, companiesModulBank); // Вписывает выписки на лист "выписки"
  setPayments(sheet, paymentsTochka, dataTochka.companies);
}

function runUpdateBillsFromDrive(){
  let ss = SpreadsheetApp.openById(tableId)

  let driveObject = new Drive()
  let sheetAccount = new SheetAccount(ss)

  /**
   * @type {globalThis.DriveApp.File[]}
   */
  let files = driveObject.getFiles1C()

  /**
   * @type {Format1CData[]}
   */
  let arrayObjects1C = []

  files.forEach(value => {
    let object1C = driveObject.parse1CFile(value)
    arrayObjects1C.push(object1C)
  })

  arrayObjects1C.forEach(value => {
    sheetAccount.insertOrUpdate(value)
  })
}
