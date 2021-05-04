/**
 * @class SheetDb
 */
class SheetDb{
  /**
   * @param {globalThis.SpreadsheetApp.Spreadsheet} ss
   */
  constructor(ss){
    this.sheet = ss.getSheetByName("БД")
    this.avaliableContractors = null
    this.avaliableTax = null
    this.avaliablePurposes = null
    this.avaliableFio = null
    this.avaliableAccount = null
  }

  /**
   * Записывает объект доступных контрагентов
   * @return {void}
   */
   setObjectContractors(){
    let contractorsHeader = this.sheet.getRange(2,1,1,8).getValues().flat()
    let values = this.sheet.getRange(3,1,this.sheet.getLastRow()-2,8).getValues()

    this.avaliableContractors = {}
    for(let row of values){
      if(row[0] === "") continue;
      this.avaliableContractors[row[0].toLowerCase()] = {}
      let objectContractor = {}
      for(let colIndex in row ){
        objectContractor[contractorsHeader[colIndex]] = row[colIndex]
      }
      this.avaliableContractors[row[0].toLowerCase()] = objectContractor
    }

    return
  }

  /**
   * Записывает объект НДС
   * @return {void}
   */
   setObjectTax(){
    let values = this.sheet.getRange(3,9,this.sheet.getLastRow()-2,2).getValues()

    this.avaliableTax = {}
    for(let row of values){
      if(row[0] === "") continue;
      this.avaliableTax[row[0].toLowerCase()] = {
        purpose: row[0],
        value: row[1]
      }
    }
    return
  }

   /**
   * Записывает объект по назначениям платежа
   * @return {void}
   */
  setDataByPurpose(){
    let purposeHeader = this.sheet.getRange(2,11,1,5).getValues().flat()
    let values = this.sheet.getRange(3,11,this.sheet.getLastRow()-2,5).getValues()

    this.avaliablePurposes = {}
    for(let row of values){
      if(row[0] === "") continue;
      this.avaliablePurposes[row[0].toLowerCase()] = {}
      let objectPurpose = {}
      for(let colIndex in row){
        objectPurpose[purposeHeader[colIndex]] = row[colIndex]
      }
      this.avaliablePurposes[row[0].toLowerCase()] = objectPurpose
    }
    return
  }
  
  /**
   * Записывает объект с данным для платежных систем для поиска по ФИО
   */
  setDataByFioForPaymentSystem(){
    let contractorsHeader = this.sheet.getRange(2,16,1,6).getValues().flat()
    let values = this.sheet.getRange(3,16,this.sheet.getLastRow(),6).getValues()

    this.avaliableFio = {}
    for(let row of values){
      if(row[1] == "") continue
      this.avaliableFio[row[1].toLowerCase()] = {}
      let objectContractor = {}
      for(let colIndex in row ){
        objectContractor[contractorsHeader[colIndex]] = row[colIndex]
      }
      this.avaliableFio[row[1].toLowerCase()] = objectContractor
    }
    return 
  }

  /**
   * Записывает объект с данным для платежных систем для поиска по номеру счета
   */
  setDataByAccountForPaymentSystem(){
    let contractorsHeader = this.sheet.getRange(2,16,1,6).getValues().flat()
    let values = this.sheet.getRange(3,16,this.sheet.getLastRow(),6).getValues()

    this.avaliableAccount = {}
    for(let row of values){
      if(row[0] == "") continue
      this.avaliableAccount[row[0]] = {}
      let objectContractor = {}
      for(let colIndex in row ){
        objectContractor[contractorsHeader[colIndex]] = row[colIndex]
      }
      this.avaliableAccount[row[0]] = objectContractor
    }
    return 
  }

  /**
   * Ищет контрагента,которые существует в БД. 
   * Важно, перед выполнением функции должна быть выполнена функция "setObjectContractors" текущего объекта,
   * иначе будет ошибка
   * @param {string} contractor
   * @return {object|undefined} Возвращает объект данных о контрагенте или ничего
   */
  findContractor(contractor){
    let objectContractors = this.avaliableContractors
    if(!objectContractors) throw new Error("Объект с контрагентами пустой, выполните функцию - setObjectContractors")

    return objectContractors[contractor.toLowerCase()];
  }

  /**
   * Ищет тип ндс,которые существует в БД. 
   * Важно, перед выполнением функции должна быть выполнена функция "setObjectTax" текущего объекта,
   * иначе будет ошибка
   * @param {string} purpose
   * @return {object|undefined} Возвращает объект данных о НДС или ничего
   */
  findTax(purpose){
    let objectTax = this.avaliableTax
    if(!objectTax) throw new Error("Объект с контрагентами пустой, выполните функцию - setObjectTax")
    purpose = purpose.toLowerCase()
    for(let prop in objectTax){
      if(purpose.indexOf(prop) === -1) continue
      return objectTax[prop].value
    }
  }

  /**
   * Ищет подходящее назначение платежа,которые существует в БД. 
   * Важно, перед выполнением функции должна быть выполнена функция "setDataByPurpose" текущего объекта,
   * иначе будет ошибка
   * @param {string} purpose
   * @return {object|undefined} Возвращает объект данных связынных с назначением платеа или ничего
   */
  findPurpose(purpose){
    let objectPurpose = this.avaliablePurposes
    if(!objectPurpose) throw new Error("Объект с контрагентами пустой, выполните функцию - setDataByPurpose")
    purpose = purpose.toLowerCase()
    for(let prop in objectPurpose){
      if(purpose.indexOf(prop) === -1) continue
      return objectPurpose[prop]
    }
  }

  /**
   * Поиск по фио для платежных систем
   * @param {string} fio
   * @return {object|undefined}
   */
  findDataByFio(fio){
    fio = fio.toLowerCase()
    if(!this.avaliableFio) throw new Error("Объект с фио пустой, выполните функцию - setDataByFioForPaymentSystem")
    let data = this.avaliableFio[fio];
    if(!data) return
    return data
  }

  /**
   * Поиск по счету для платежных систем
   * @param {string} account
   * @return {object|undefined}
   */
  findDataByAccount(account){
    if(!this.avaliableAccount) throw new Error("Объект с аккаунтами счетов пустой, выполните функцию - setDataByAccountForPaymentSystem")
    let data = this.avaliableAccount[account];
    if(!data) return
    return data
  }
}
