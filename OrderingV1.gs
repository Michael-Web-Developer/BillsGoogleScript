/**
 * @typedef {{
 *  "Дата": string,
 * 	"Счет": string,
 * 	"Сумма": number,
 * 	"Плательщик": string,
 * 	"Получатель": string,
 * 	"Назначние платежа": string,
 * 	"Статья": string,
 *  "Месяц": string,
 * 	"Счет получателя": string,
 * 	"Ответственный": string,
 * 	"Продукт": string,
 * 	"ЦФО": string,
 * 	"Специфика": string,
 * 	"НДС": string,
 * 	"Сумма без НДС": number,
 * 	"поля 1-3, запас": string,
 * }} FormatOrderingVersion1
 */

/**
 * @class OrderingV1
 */
class OrderingV1 {
  /**
   * @param {globalThis.SpreadsheetApp.Spreadsheet} ss
   */
  constructor(ss, objectDb){
    this.sheet = ss.getSheetByName("выписки")
  }

  /**
   * Получение массива формата для листа выписки из данных 1С
   * @param {Format1CData} format1CData Распарсенные данные из файла 1С
   * @param {globalThis.SheetDb} objectDb
   * @return {FormatOrderingVersion1[]}
  */
  transferFormat1CDataToOrderingVer1(format1CData, objectDb) {
    let outputArray = []

    for(let transaction of format1CData["транзакции"]){
      /**
       * @type {FormatOrderingVersion1}
       */
      let objectOutput = this.getEmptyFormatOrderingVersion1()
      
      let typeTransaction
      let dateTransaction

      if(transaction["датасписано"]){
        typeTransaction = "расход"
        dateTransaction = Utilities.formatDate(transaction["датасписано"], "GMT+3", "dd.MM.yyyy")
      } else {
        typeTransaction = "доход"
        dateTransaction = Utilities.formatDate(transaction["датапоступило"], "GMT+3", "dd.MM.yyyy")
      }

      //Данные напрямую из выписок
      objectOutput["Дата"] = dateTransaction
      objectOutput["Сумма"] = typeTransaction === "доход" ? transaction["сумма"] : transaction["сумма"] * -1
      objectOutput["Назначние платежа"] = transaction["назначениеплатежа"]
      objectOutput["Счет получателя"] = transaction["получательсчет"]
      objectOutput["Счет"] = format1CData["расчсчет"]
      objectOutput["Месяц"] = dateTransaction.split(".")[1]

      let reciepientObject = objectDb.findContractor(transaction["получатель1"])
      let senderObject = objectDb.findContractor(transaction["плательщик1"])
      let tax = objectDb.findTax(transaction["назначениеплатежа"])
      let purposeData = objectDb.findPurpose(transaction["назначениеплатежа"])


      if(reciepientObject){
        objectOutput["Получатель"] = reciepientObject["контрагент УУ"]
        objectOutput["Статья"] = reciepientObject["статья"]
        objectOutput["Продукт"] = reciepientObject["продукт"]
        objectOutput["ЦФО"] = reciepientObject["ЦФО"]
        objectOutput["Специфика"] = reciepientObject["Специфика"]
        objectOutput["НДС"] = reciepientObject["НДС"]
      } else {
        objectOutput["Получатель"] = transaction["получатель1"]
      }

      if(senderObject){
        objectOutput["Плательщик"] = senderObject["контрагент УУ"]
      } else {
        objectOutput["Плательщик"] = transaction["плательщик1"]
      }

      if(tax){
        objectOutput["НДС"] = tax
      }

      if(purposeData){
        objectOutput["Статья"] = purposeData["статья"]
        objectOutput["Продукт"] = purposeData["продукт"]
        objectOutput["ЦФО"] = purposeData["ЦФО"]
        objectOutput["Специфика"] = purposeData["Специфика"]
      }


      if(objectOutput["НДС"].toLowerCase() === "с ндс"){
          objectOutput["Сумма без НДС"] = Number((objectOutput["Сумма"]*5/6).toFixed(2))
        } else {
          objectOutput["Сумма без НДС"] = objectOutput["Сумма"]
        }

      outputArray.push(objectOutput)
    }

    return outputArray;
  }

  /**
   * Получение пустого объект формата для выписок 1
   * @returns {FormatOrderingVersion1}
   */
  getEmptyFormatOrderingVersion1(){
    return {
        "Дата": "",
        "Счет": "",
        "Сумма": 0,
        "Плательщик": "",
        "Получатель": "",
        "Назначние платежа": "",
        "Статья": "",
        "Месяц": "",
        "Счет получателя": "",
        "Ответственный": "",
        "Продукт": "",
        "ЦФО": "",
        "Специфика": "",
        "НДС": "",
        "Сумма без НДС": 0,
        "поля 1-3, запас": ""
      }
  }
  
  /**
   * Вставка данных в лист "выписки"
   * @param {FormatOrderingVersion1[]}
   * @return {void}
   */
  insertDataInOrderingVer1(arrayFormatOrderingV1){
    let header = this.getHeader()
    let arrayToInsert = []
    for(let objectFormat of arrayFormatOrderingV1){
      let arrayKeys = Object.keys(objectFormat)
      let arrayToAdd = new Array(arrayKeys.length)
      for(let property of arrayKeys){
        let indexCol = header.indexOf(property)
        if(indexCol === -1) throw new Error(`Не найдено столбца с именем ${property}`)
        arrayToAdd[indexCol] = objectFormat[property]
      }
      arrayToInsert.push(arrayToAdd)
    }
    arrayToInsert.sort((a,b) => {
      if (a[0] < b[0]) {
        return -1;
      }
      if (a[0] > b[0]) {
        return 1;
      }
      return 0
      }
    )

    if(arrayToInsert.length === 0) return 

    //Вставляем данные из выписок в лист Выписки
    this.sheet.getRange(this.sheet.getLastRow() + 1, 1, arrayToInsert.length, arrayToInsert[0].length).setValues(arrayToInsert)
  }

  /**
   * Получение шапки листа Выписки
   * @return {Array}
   */
  getHeader(){
    let values = this.sheet.getRange(1,1,1,this.sheet.getLastColumn()).getValues().flat()
    return values
  }
}


