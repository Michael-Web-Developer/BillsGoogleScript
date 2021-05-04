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

  let driveObject = new DriveBills()
  let sheetAccount = new SheetAccount(ss)
  let dbObject = new SheetDb(ss)
  dbObject.setObjectContractors()
  dbObject.setDataByPurpose()
  dbObject.setObjectTax()
  dbObject.setDataByAccountForPaymentSystem()
  dbObject.setDataByFioForPaymentSystem()
  let objectOrderingV1 = new OrderingV1(ss)
  let objectOrderingV2 = new OrderingV2(ss)

  /**
   * @type {globalThis.DriveApp.File[]}
   */
  let files = driveObject.getFiles1C()
  let filesQiwi = driveObject.getFilesQiwi()
  let filesMandarin = driveObject.getFilesMandarin()
  let filesEkselent = driveObject.getFilesEkselent()
  let filesCypix = driveObject.getFilesCypix()

  /**
   * @type {Format1CData[]}
   */
  let arrayObjects1C = []

  /**
   * @type {FormatPaymentSystem[]}
   */
  let arrayPaymentSystemsData = []
  /**
   * @type {FormatOrderingVersion1[]}
   */
  let arrayObjectsForOrderingVer1 = []

  let arrayFilesForArchive = []

  files.forEach(value => {
    let object1C = driveObject.parse1CFile(value)
    arrayObjects1C.push(object1C)
    arrayFilesForArchive.push(value)
  })

  filesQiwi.forEach(value => {
    let formatPaymentSystem = driveObject.parseQiwiFile(value, objectOrderingV2)
    arrayPaymentSystemsData.push(formatPaymentSystem)
    arrayFilesForArchive.push(value)
  })

  filesMandarin.forEach(value => {
    let formatPaymentSystem = driveObject.parseMandarinFile(value, objectOrderingV2)
    arrayPaymentSystemsData.push(formatPaymentSystem)
    arrayFilesForArchive.push(value)
  })

  filesEkselent.forEach(value => {
    let formatPaymentSystem = driveObject.parseEleksnetFile(value, objectOrderingV2)
    arrayPaymentSystemsData.push(formatPaymentSystem)
    arrayFilesForArchive.push(value)
  })

  filesCypix.forEach(value => {
    let formatPaymentSystem = driveObject.parseCypixFile(value, objectOrderingV2)
    arrayPaymentSystemsData.push(formatPaymentSystem)
    arrayFilesForArchive.push(value)
  })

  arrayObjects1C.forEach(value => {
    sheetAccount.insertOrUpdate1CBills(value)
    let arrayPrepareOrderingVer1 = objectOrderingV1.transferFormat1CDataToOrderingVer1(value, dbObject)
    arrayObjectsForOrderingVer1 = arrayObjectsForOrderingVer1.concat(arrayPrepareOrderingVer1)
  })

  sheetAccount.insertOrUpdatePaymentSystemsData(arrayPaymentSystemsData)
  arrayPaymentSystemsData = objectOrderingV2.transferDataPaymentSystems(arrayPaymentSystemsData, dbObject)

  objectOrderingV1.insertDataInOrderingVer1(arrayObjectsForOrderingVer1)
  objectOrderingV2.insertDataInOrderingVer2(arrayPaymentSystemsData)

  //driveObject.moveToArchive(arrayFilesForArchive)
}
