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
 * @class Drive
 */
class Drive{
  constructor(){
    this.idFolderIn = "1w6kLHbSvUw47nR0nUJq31AhVOgfocu40";
    this.idFolderOut = "1LfXAOu3uhYMfQVFOUDvzuPEyBJ1CCtse"
    this.idFolder1C = "1y06HdG5OFsuKDihw1g3PE3zJoAMf6dtu"
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
        let date = new Date(`${parseDate[2]}-${parseDate[1]}-${parseDate[0]}`)
        objectTransaction1C["датасписано"] = date
        continue;
      }

      if(name === "датапоступило"){
        let parseDate = value.split(".");
        let date = new Date(`${parseDate[2]}-${parseDate[1]}-${parseDate[0]}`)
        objectTransaction1C["датапоступило"] = date
        continue;
      }

      if(name === "сумма"){
        objectTransaction1C["сумма"] = Number(value)
        continue;
      }
      
      if(name === "плательщик1"){
        objectTransaction1C["плательщик1"] = value
        continue;
      }

      if(name === "получатель1"){
        objectTransaction1C["получатель1"] = value
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
}

function testDrive(){
  let test = new Drive()
  let files = test.getFiles1C();
  for(let file of files){
    let result = test.parse1CFile(file)
  }
}
