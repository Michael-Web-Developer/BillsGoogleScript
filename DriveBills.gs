/**
 * @typedef {{
 *    "счет":string,
 *    "сальдо":number,
 *    "дата":Date|null,
 *    "транзакции": FormatOrderingVersion2[]
 *  }} FormatPaymentSystem Формат спарсеных данных платежных систем
 */

/**
 * @typedef {{
 *    "расчсчет":string|null,
 *    "датаначала":Date|null,
 *    "датаконца":Date|null,
 *    "начальныйостаток":number|null,
 *    "конечныйостаток":number|null,
 *    "транзакции": Transaction1С[]
 *  }} Format1CData Формат спарсеных данных 1С
 */

/**
 * @typedef {{
 *    "сумма":number|null,
 *    "плательщик1":string|null,
 *    "получатель1":string|null,
 *    "назначениеплатежа":string|null,
 *    "получательсчет":string|null,
 *    "датасписано":Date|null,
 *    "датапоступило":Date|null
 *  }} Transaction1C Распарсенные данные транзакции
 */

/**
 * @class DriveBills
 */
class DriveBills{
  constructor(){
    this.idFolderIn = "1w6kLHbSvUw47nR0nUJq31AhVOgfocu40";
    this.idFolderOut = "1LfXAOu3uhYMfQVFOUDvzuPEyBJ1CCtse"
    this.idFolder1C = "1y06HdG5OFsuKDihw1g3PE3zJoAMf6dtu"
    this.idFolderQiwi = "1ObpROj6OfSmQRsGhPIXXGbY3LxblQB2j"
    this.idFolderMandarin = "1cBq7asPkbuBj5SfBj1k0pGvDMSP4nrml"
    this.idFolderEkselent = "1XwYryeD8vpMgsGQSN_aJEMHUm95r91GD"
    this.idFolderCypix = "1mZb60DP9KkqFmHuJpnP8rCwJ7ufwToAp"
  }
  
  /**
   * Возвращает массив файлов 1С
   * @return {globalThis.DriveApp.File[]}
   */
  getFiles1C(){
    let folder = DriveApp.getFolderById(this.idFolder1C)
    let filesIterator = folder.getFiles()
    /**
     * @type {globalThis.DriveApp.File[]}
     */
    let arrayOutput = []

    while(filesIterator.hasNext()){
      let file = filesIterator.next()
      if(!file.getName().match(/([.]txt)/)) continue;
      arrayOutput.push(file)
    }

    return arrayOutput
  }
  /**
   * Возвращает массив файлов выписок Из Qiwi
   * @return {globalThis.DriveApp.File[]}
   */
  getFilesQiwi(){
    let folder = DriveApp.getFolderById(this.idFolderQiwi)
    let filesIterator = folder.getFiles()
    /**
     * @type {globalThis.DriveApp.File[]}
     */
    let arrayOutput = []

    while(filesIterator.hasNext()){
      let file = filesIterator.next()
      if(!file.getName().match(/([.]xlsx)/)) continue;
      arrayOutput.push(file)
    }

    return arrayOutput
  }

  /**
   * Возвращает массив файлов выписок Из папки Мандарин
   * @return {globalThis.DriveApp.File[]}
   */
  getFilesMandarin(){
    let folder = DriveApp.getFolderById(this.idFolderMandarin)
    let filesIterator = folder.getFiles()
    /**
     * @type {globalThis.DriveApp.File[]}
     */
    let arrayOutput = []

    while(filesIterator.hasNext()){
      let file = filesIterator.next()
      if(!file.getName().match(/([.]csv)/)) continue;
      arrayOutput.push(file)
    }

    return arrayOutput
  }

  /**
   * Возвращает массив файлов выписок Из папки Экселент
   * @return {globalThis.DriveApp.File[]}
   */
  getFilesEkselent(){
    let folder = DriveApp.getFolderById(this.idFolderEkselent)
    let filesIterator = folder.getFiles()
    /**
     * @type {globalThis.DriveApp.File[]}
     */
    let arrayOutput = []

    while(filesIterator.hasNext()){
      let file = filesIterator.next()
      if(!file.getName().match(/([.]csv)/)) continue;
      arrayOutput.push(file)
    }

    return arrayOutput
  }

  /**
   * Возвращает массив файлов выписок Из папки Cypix
   * @return {globalThis.DriveApp.File[]}
   */
  getFilesCypix(){
    let folder = DriveApp.getFolderById(this.idFolderCypix)
    let filesIterator = folder.getFiles()
    /**
     * @type {globalThis.DriveApp.File[]}
     */
    let arrayOutput = []

    while(filesIterator.hasNext()){
      let file = filesIterator.next()
      if(!file.getName().match(/([.]csv)/)) continue;
      arrayOutput.push(file)
    }

    return arrayOutput
  }

  /**
   * Парсинг файл выписки 1c
   * @param {globalThis.DriveApp.File} file
   */
  parse1CFile(file){
    let text = file.getBlob().getDataAsString("windows-1251")
    if(!text.trim()) return;

    let rows = text.split("\n");

    /** @type {Format1CData} */
    let outputObject = {
     "расчсчет":null,
     "датаначала":null,
     "датаконца":null,
     "начальныйостаток":null,
     "конечныйостаток":null,
     "транзакции": []
    }

    /** @type {Transaction1C} */
    let objectTransaction1C = this.getEmptyObjectTransaction1C()
    for(let row of rows){

      let parseRow  = row.split("=");

      let name = parseRow[0].toLowerCase().trim()
      let value = parseRow[1] ? parseRow[1].trim() : null
      
      if(name == "конецдокумента"){
        outputObject["транзакции"].push(objectTransaction1C)
        objectTransaction1C = this.getEmptyObjectTransaction1C()
        continue;
      }

      if(name === "расчсчет"){
        outputObject["расчсчет"] = value
        continue;
      }

      if(name === "датаначала"){
        let parseDate = value.split(".");    
        let date = new Date(`${parseDate[2]}-${parseDate[1]}-${parseDate[0]}`)
        outputObject["датаначала"] = date
        continue;
      }
      
      if(name === "датаконца"){
        let parseDate = value.split(".");
        let date = new Date(`${parseDate[2]}-${parseDate[1]}-${parseDate[0]}`)
        outputObject["датаконца"] = date
        continue;
      }

      if(name === "начальныйостаток"){
        outputObject["начальныйостаток"] = Number(value)
        continue;
      }

      if(name === "конечныйостаток"){
        outputObject["конечныйостаток"] = Number(value)
        continue;
      }

      if(name === "датасписано"){
        let parseDate = value.split(".");
        if(parseDate.length <= 1) continue
        let date = new Date(`${parseDate[2]}-${parseDate[1]}-${parseDate[0]}`)
        objectTransaction1C["датасписано"] = date
        continue;
      }

      if(name === "датапоступило"){
        let parseDate = value.split(".");
        if(parseDate.length <= 1) continue
        let date = new Date(`${parseDate[2]}-${parseDate[1]}-${parseDate[0]}`)
        objectTransaction1C["датапоступило"] = date
        continue;
      }

      if(name === "сумма"){
        objectTransaction1C["сумма"] = Number(value)
        continue;
      }
      
      if(name === "плательщик1" || name === "плательщик"){
        //if(!objectTransaction1C["плательщик1"]){
          objectTransaction1C["плательщик1"] = value
        //}
        continue;
      }

      if(name === "получатель1" || name === "получатель"){
        //if(!objectTransaction1C["получатель1"]){
          objectTransaction1C["получатель1"] = value
        //}
         
        continue;
      }

      if(name === "назначениеплатежа"){
        objectTransaction1C["назначениеплатежа"] = value
        continue;
      }

      if(name === "получательсчет"){
        objectTransaction1C["получательсчет"] = value
        continue;
      }
    }

    return outputObject
  }

  /**
   * Парсинг файл выписки Qiwi
   * @param {globalThis.DriveApp.File} file
   * @param {globalThis.OrderingV2} orderingV2
   * @return {FormatPaymentSystem}
   */
  parseQiwiFile(file, orderingV2){
    let blob = file.getBlob();
    let data = {
        title: file.getName().replace(".xlsx", ""),
        key: file.getId(),
        parents: [{id: this.idFolderQiwi}],
        mimeType: MimeType.GOOGLE_SHEETS
    };
    
    let fileInfo = Drive.Files.insert(data, blob, {convert: true})
    let ss = SpreadsheetApp.openById(fileInfo.id)
    let outputObject = this.getEmptyObjectFormatPaymentSystem()
    outputObject["счет"] = accountQiwi

    let sheet = ss.getSheets()[0]

    let values = sheet.getRange(3,1,sheet.getLastRow() - 2, sheet.getLastColumn()).getValues()

    for(let row of values){
      if(row[6].toLowerCase().match(/всего исходящие/)) break;
      let objectFormatOrderingV2 = orderingV2.getEmptyFormatOrdering()

      objectFormatOrderingV2["Дата"] = row[1]
      objectFormatOrderingV2["Счет"] = accountQiwi
      objectFormatOrderingV2["Комиссия"] = Number(row[12])
      objectFormatOrderingV2["Месяц"] = row[1].split(".")[1]
      objectFormatOrderingV2["Плательщик"] = row[21]
      objectFormatOrderingV2["Сумма"] = Number(row[10])
      objectFormatOrderingV2["Счет получателя"] = row[49]

      
      outputObject["сальдо"] += Number(row[10]) 
      if(!outputObject["дата"]){
        let parseDate = row[1].split(".")
        let date = new Date(`${parseDate[2]}-${parseDate[1]}-${parseDate[0]}`)
        outputObject["дата"] = date
      }
      outputObject["транзакции"].push(objectFormatOrderingV2)  
    }
    let fileSS = DriveApp.getFileById(fileInfo.id)
    fileSS.setTrashed(true)
    return outputObject
  }

  /**
   * Парсинг файл выписки Мандарин
   * @param {globalThis.DriveApp.File} file
   * @param {globalThis.OrderingV2} orderingV2
   * @return {FormatPaymentSystem}
   */
  parseMandarinFile(file, orderingV2){
    let content = file.getBlob().getDataAsString("windows-1251")
    let outputObject = this.getEmptyObjectFormatPaymentSystem()
    outputObject["счет"] = accountMandarin

    let rows = content.split("\n")
    for(let row of rows){
      let cols = row.split(";")
      if(cols[0] === "" || cols[4] !== "Успешно") continue
      let objectFormatOrderingV2 = orderingV2.getEmptyFormatOrdering()
      let date = cols[3].replace(/["]/g, "")
      let splitDateAndTime = date.split(" ")
      date = splitDateAndTime[0].split("-")
      let dateObject = new Date(`${date[2]}-${date[1]}-${date[0]}`)
      objectFormatOrderingV2["Дата"] = `${date[0]}.${date[1]}.${date[2]}`

      if(!outputObject["дата"]){
        outputObject["дата"] = dateObject
      }

      let sum = Number(cols[13].replace(",", "."))
      let comision = Number(cols[15].replace(",", "."))
      objectFormatOrderingV2["Счет"] = accountMandarin
      objectFormatOrderingV2["Сумма"] = sum
      outputObject["сальдо"] += sum
      objectFormatOrderingV2["Плательщик"] = payerMandarin
      objectFormatOrderingV2["Счет получателя"] = cols[16]
      objectFormatOrderingV2["Месяц"] = date[1]
      objectFormatOrderingV2["Комиссия"] = comision

      outputObject["транзакции"].push(objectFormatOrderingV2)
    }

    return outputObject
  }

  /**
   * Парсинг файл выписки Элекснет
   * @param {globalThis.DriveApp.File} file
   * @param {globalThis.OrderingV2} orderingV2
   * @return {FormatPaymentSystem}
   */
  parseEleksnetFile(file, orderingV2){
    let content = file.getBlob().getDataAsString()
    let outputObject = this.getEmptyObjectFormatPaymentSystem()
    outputObject["счет"] = accountEleksnet

    let rows = Utilities.parseCsv(content)

    for(let row of rows){
      let cols = row
      
      if(cols[0] === "" || cols[15] !== "approved") continue
      let objectFormatOrderingV2 = orderingV2.getEmptyFormatOrdering()
      let date = cols[3]
      date = date.split("-")
      let dateObject = new Date(`${date[0]}-${date[1]}-${date[2]}`)
      objectFormatOrderingV2["Дата"] = `${date[2]}.${date[1]}.${date[0]}`

      if(!outputObject["дата"]){
        outputObject["дата"] = dateObject
      }

      let sum = Number(cols[21].replace(",", "."))
      let comision = Number(cols[68].replace(",", "."))
      objectFormatOrderingV2["Счет"] = accountEleksnet
      objectFormatOrderingV2["Сумма"] = sum
      outputObject["сальдо"] += sum
      objectFormatOrderingV2["Плательщик"] = cols[6]
      objectFormatOrderingV2["Счет получателя"] = cols[69]
      objectFormatOrderingV2["Месяц"] = date[1]
      objectFormatOrderingV2["Комиссия"] = comision

      outputObject["транзакции"].push(objectFormatOrderingV2)
    }

    return outputObject
  }

    /**
   * Парсинг файл выписки Cypix
   * @param {globalThis.DriveApp.File} file
   * @param {globalThis.OrderingV2} orderingV2
   * @return {FormatPaymentSystem}
   */
  parseCypixFile(file, orderingV2){
    let content = file.getBlob().getDataAsString("windows-1251")
    let outputObject = this.getEmptyObjectFormatPaymentSystem()
    outputObject["счет"] = accountCypix

    let rows = content.split("\n")
    for(let row of rows){
      let cols = row.split(";")
      
      if(cols[0] === "" || cols[16] !== "DONE") continue
      let objectFormatOrderingV2 = orderingV2.getEmptyFormatOrdering()
      let date = cols[8].replace(/["]/g, "")
      let splitDateAndTime = date.split(" ")
      date = splitDateAndTime[0].split("-")
      let dateObject = new Date(`${date[2]}-${date[1]}-${date[0]}`)
      objectFormatOrderingV2["Дата"] = `${date[2]}.${date[1]}.${date[0]}`

      if(!outputObject["дата"]){
        outputObject["дата"] = dateObject
      }

      let sum = Number(cols[9].replace(",", "."))
      let comision = +(Number(cols[10].replace(",", ".")) - sum).toFixed(2)
      objectFormatOrderingV2["Счет"] = accountCypix
      objectFormatOrderingV2["Сумма"] = sum
      outputObject["сальдо"] += sum
      objectFormatOrderingV2["Плательщик"] = payerCypix
      objectFormatOrderingV2["Получатель"] = cols[18]
      objectFormatOrderingV2["Месяц"] = date[1]
      objectFormatOrderingV2["Комиссия"] = comision

      outputObject["транзакции"].push(objectFormatOrderingV2)
    }

    return outputObject
  }

  /**
   * Возвращает пустой объект типа Transaction1C
   * @returns {Transaction1C}
   */
  getEmptyObjectTransaction1C(){
    return {
     "сумма":null,
     "плательщик1":null,
     "получатель1":null,
     "назначениеплатежа":null,
     "получательсчет":null,
     "датасписано":null,
     "датапоступило":null
    }
  }

  /**
   * Возвращает пустой объект типа FormatPaymentSystem
   * @return {FormatPaymentSystem}
   */
  getEmptyObjectFormatPaymentSystem(){
    return {
      "счет":"",
      "сальдо":0,
      "дата":null,
      "транзакции": []
    }
  }

  /**
   * Перенос файлов в архив
   * @param {globalThis.DriveApp.File[]} files
   * @return {void}
   */
  moveToArchive(files){
    let folder = DriveApp.getFolderById(this.idFolderOut)
    for(let file of files){
      file.moveTo(folder)
    }
  }
}

function testDrive(){
  let test = new Drive()
  let files = test.getFiles1C();
  for(let file of files){
    let result = test.parse1CFile(file)
  }
}
