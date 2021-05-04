class SheetAccount {
  /**
   * @param {globalThis.SpreadsheetApp.Spreadsheet} ss
   */
  constructor(ss){
    this.ss = ss
    this.sheet = ss.getSheetByName("счета")
  }

  /**
   * Получение массива существующих аккаунтов
   * @return {Array}
   */
  getAccountsOnSheet(){
    let accounts = this.sheet.getRange(1, 1, 1, this.sheet.getLastColumn()).getValues().flat()
    return accounts;
  }

  /**
   * Получение массива с датами
   * @return {Array}
   */
  getDates(){
    let dates = this.sheet.getRange(1,1,this.sheet.getLastRow())
      .getValues()
      .flat()
      .filter(value => value != '')
    return dates
  }

  /**
   * Добавление или создание информацию о счете
   * @param {Format1CData} format1CData
   * @return {void}
   */
  insertOrUpdate1CBills(format1CData){
    let accounts = this.getAccountsOnSheet()
    let dates = this.getDates()

    let currentAccount = format1CData["расчсчет"];
    let currentDate = format1CData["датаначала"]
    currentDate = Utilities.formatDate(currentDate, "GMT+3", "dd.MM.yyyy")
    let startAmount = format1CData["начальныйостаток"]
    let endAmount = format1CData["конечныйостаток"]

    let indexAccount = accounts.indexOf(currentAccount)

    if(indexAccount === -1) {
      this.addNewAccount(currentAccount)
      accounts = this.getAccountsOnSheet()
      indexAccount = accounts.indexOf(currentAccount)
    }

    let indexDate;

    dates.forEach((value,index) => {
      if(value instanceof Date){
        let dateStringRow = Utilities.formatDate(value, "GMT+3", "dd.MM.yyyy")
        if(dateStringRow === currentDate) indexDate = index
      } else if(value === currentDate){
        indexDate = index
      }
    })

    let row
    let column

    if(!indexDate){
      this.sheet.getRange(dates.length + 1, 1).setValue(currentDate)
      row = this.sheet.getLastRow()
      column = indexAccount+1
      this.sheet.getRange(row, column, 1, 2).setValues([[startAmount, endAmount]])
      this.checkValuesOfAccounts(format1CData, row, column)
    } else {
      row = +indexDate + 1
      column = indexAccount+1
      this.sheet.getRange(row, column, 1, 2).setValues([[startAmount, endAmount]])
      this.checkValuesOfAccounts(format1CData, row, column)
    }

    return
  }

  /**
   * Добавление или создание информацию о счете
   * @param {FormatPaymentSystem[]} formatPaymentSystems
   * @return {void}
   */
  insertOrUpdatePaymentSystemsData(formatPaymentSystems){
    let accounts = this.getAccountsOnSheet()
    let dates = this.getDates()

    for(let dataPaymentSystem of formatPaymentSystems){
      let currentAccount = dataPaymentSystem["счет"];
      let currentDate = dataPaymentSystem["дата"]
      currentDate = Utilities.formatDate(currentDate, "GMT+3", "dd.MM.yyyy")

      let indexAccount = accounts.indexOf(currentAccount)

      if(indexAccount === -1) {
        Browser.msgBox("Ошибка!", 'На листе "счета" нет счета: ' + currentAccount, Browser.Buttons.OK);
        continue;
      }

      let indexDate;

      dates.forEach((value,index) => {
        if(value instanceof Date){
          let dateStringRow = Utilities.formatDate(value, "GMT+3", "dd.MM.yyyy")
          if(dateStringRow === currentDate) indexDate = index
        } else if(value === currentDate){
          indexDate = index
        }
      })

      let row
      let column

      if(!indexDate){
        this.sheet.getRange(dates.length + 1, 1).setValue(currentDate)
        row = this.sheet.getLastRow()
        column = indexAccount+1
        let prevData = this.getPrevAmount(row, column)
        let endAmount = prevData.endAmount + dataPaymentSystem["сальдо"] * -1
        this.sheet.getRange(row, column, 1, 2).setValues([[prevData.endAmount, endAmount]])
      } else {
        row = +indexDate + 1
        column = indexAccount+1
        let prevData = this.getPrevAmount(row, column)
        let endAmount = prevData.endAmount + dataPaymentSystem["сальдо"] * -1
        this.sheet.getRange(row, column, 1, 2).setValues([[prevData.endAmount, endAmount]])
      }
    }
    return
  }

  getPrevAmount(row,column){
    let prevData = this.sheet.getRange(row-1, column, 1, 2).getValues()
    let outputObject = {
      startAmount:null,
      endAmount:null
    }
    outputObject.startAmount = prevData[0][0]
    outputObject.endAmount = prevData[0][1]
    return outputObject
  }
  /**
   * Создание нового счета
   * @param {string} account
   * @return {void}
   */
  addNewAccount(account){
    this.sheet.insertColumnsAfter(this.sheet.getLastColumn(), 2)
    this.sheet.getRange(1,this.sheet.getLastColumn()+1).setValue(account)

    this.sheet.getRange(1,this.sheet.getLastColumn(), 1, 2)
      .merge()
      .setBorder(false, false, true, true, false, false)
  }

  /**
   * Проверка предыдущих начальных остатков счета
   * @param {Format1CData} data Данные с текущими значении об остатках
   * @param {number} row Строка текущих данных
   * @param {number} column Колонка с текущим счетом (начальный остаток)
   * @return {void}
   */
  checkValuesOfAccounts(data, row, column){
    let currentDate = Utilities.formatDate(data["датаначала"], "GMT+3", "yyyy-MM-dd");
    let prevDay = new Date(currentDate)
    prevDay.setDate(prevDay.getDate() - 1)
    prevDay = Utilities.formatDate(prevDay, "GMT+3", "dd.MM.yyyy");

    if (this.sheet.getRange(row-1, 1).getValue() == prevDay){
      const lastValue = this.sheet.getRange(row-1, column+1).getValue();
      const value = data["начальныйостаток"];
      if (lastValue != value && lastValue !== ""){
        this.sheet.getRange(row-1, column+1)
          .setNote("Не совпадает со следующим днём")
          .setBackgroundRGB(255, 0, 0)

        this.sheet.getRange(row, column)
          .setNote("Не совпадает с прошлым днём")
          .setValue("")
          .setBackgroundRGB(255, 0, 0)

        Browser.msgBox("Ошибка!", 'На листе "счета" не совпадают данные. Cчет: ' + this.sheet.getRange(1, column).getValue(), Browser.Buttons.OK);
      }
    }
  }
}