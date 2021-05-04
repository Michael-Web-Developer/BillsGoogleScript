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
 * 	"Комиссия": number,
 * }} FormatOrderingVersion2
 */

/**
 * @class OrderingV2
 */
class OrderingV2{
  constructor(ss){
    this.sheet = ss.getSheetByName("выписки 2")
  }

  /**
   * Получение пустого объекта FormatOrderingVersion2
   * @return {FormatOrderingVersion2}
   */
  getEmptyFormatOrdering(){
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
        "Комиссия": 0
      }
  }

  /**
   * Добавляет в поля нужные значени из базы данных
   * @param {FormatPaymentSystem[]} dataPaymentSystem
   * @param {globalThis.SheetDb} dbObject
   * @return {FormatPaymentSystem[]}
   */
  transferDataPaymentSystems(dataPaymentSystem, dbObject){
    for(let data of dataPaymentSystem){
      let transactions = data["транзакции"]
      for(let transaction of transactions){
        let dataFromDb

        if(transaction["Получатель"] != ""){
          dataFromDb = dbObject.findDataByFio(transaction["Получатель"])
        }

        if(transaction["Счет получателя"] != ""){
          dataFromDb = dbObject.findDataByAccount(transaction["Счет получателя"])
        }

        if(dataFromDb){
          transaction["Получатель"] = dataFromDb["ФИО"]
          transaction["Статья"] = dataFromDb["статья"]
          transaction["Продукт"] = dataFromDb["продукт"]
          transaction["ЦФО"] = dataFromDb["ЦФО"]
          transaction["Специфика"] = dataFromDb["Специфика"]
          transaction["Счет получателя"] = dataFromDb["номер карты"]
        }

        transaction["Сумма"]  =  transaction["Сумма"] * -1
        transaction["Комиссия"]  =  transaction["Комиссия"] * -1
      }
    }
    return dataPaymentSystem
  }

  /**
   * Вставка данных в лист "выписки 2"
   * @param {FormatPaymentSystem[]}
   * @return {void}
   */
  insertDataInOrderingVer2(arrayFormatPaymentSystem){
    let header = this.getHeader()
    let arrayToInsert = []
    for(let objectFormat of arrayFormatPaymentSystem){
      for(let preparedDataInsert of objectFormat["транзакции"]){
        let arrayKeys = Object.keys(preparedDataInsert)
        let arrayToAdd = new Array(arrayKeys.length)
        for(let property of arrayKeys){
          let indexCol = header.indexOf(property)
          if(indexCol === -1) throw new Error(`Не найдено столбца с именем ${property}`)
          arrayToAdd[indexCol] = preparedDataInsert[property]
        }
        arrayToInsert.push(arrayToAdd)
      } 
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
   * Получение шапки листа Выписки2
   * @return {Array}
   */
  getHeader(){
    let values = this.sheet.getRange(1,1,1,this.sheet.getLastColumn()).getValues().flat()
    return values
  }
}
