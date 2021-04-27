// Запрос на получение выписок из банка.
function getObjectsModulBank(companies) { 
  const fromDay = Utilities.formatDate(new Date(new Date() - amountDaysForAPI*oneDay), "GMT+3", "yyyy-MM-dd");
  
  let resObjs = [];
  
  for (let i = 0; i < companies.length; i++){ // Пробегаемся по всем компаниям
    for (let j = 0; j < companies[i].bankAccounts.length; j++){ // Пробегаемся по всем счетам компании
      let bankAccountIdModulBank = companies[i].bankAccounts[j].id;
      
      const url = 'https://api.modulbank.ru/v1/operation-history/' + bankAccountIdModulBank;
      
      let Data;
      let objsData = [];
      let amount = 0;
      
      while(typeof Data == "undefined" || objsData.length == 50){ // Максимальное количество выписок в запросе = 50. Бегаем циклом, чтобы получить все.
        let options = {
          headers: {'Authorization': 'Bearer ' + tokenOfAuthorizationModulBank},
          method: 'POST',
          payload: JSON.stringify({
            records: 50,
            skip: amount,
            from: fromDay,
          }),
          contentType: 'application/json',
        };
        
        Data = UrlFetchApp.fetch(url, options);
        
        objsData = JSON.parse(Data);
        
        for (let i = 0; i < objsData.length; i++) {
          if (objsData[i].status == "Executed" || objsData[i].status == "Received")
            resObjs.push(objsData[i]);
        }
        amount += 50;
      }
    }
  }
  resObjs.sort(function(a, b){ //Убрать
    if (a.executed < b.executed) return -1;
    if (a.executed == b.executed) return 0;
    if (a.executed > b.executed) return 1;
  });
  return resObjs;
}

// Запрос на получение компаний для определения названия компании по companyId.
function getCompaniesModulBank(){ 
  
  const url = 'https://api.modulbank.ru/v1/account-info/';
  
  let options = {
    headers: {'Authorization': 'Bearer ' + tokenOfAuthorizationModulBank},
    method: 'POST',
    contentType: 'application/json',
  };
  
  let Data = UrlFetchApp.fetch(url, options);
  
  if (Data.getResponseCode() == 503) { // 503 ошибка - Сервер на технических работах.
    Browser.msgBox("Ошибка!", 'Сервер ModulBank проводит технические работы', Browser.Buttons.OK);
    writeError('Сервер ModulBank проводит технические работы');
    return Data.getResponseCode();
  }
  
  let companies = JSON.parse(Data);
  
  return companies;
}