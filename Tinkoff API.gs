// Запрос на получение счетов по инн компании
function getAccountsTinkoff(inn){
  let url = 'https://business.tinkoff.ru/openapi/api/v1/company/'+ inn + '/accounts';
  
  let options = {
    headers: {'Authorization': 'Bearer ' + tokenOfAuthorizationTinkoff},
    method: 'GET',
    contentType: 'application/json',
  };
  
  let Data = UrlFetchApp.fetch(url, options);
  
  if (Data.getResponseCode() == 503) { // 503 ошибка - Сервер на технических работах.
    Browser.msgBox("Ошибка!", 'Сервер ModulBank проводит технические работы', Browser.Buttons.OK);
    writeError('Сервер ModulBank проводит технические работы');
    return Data.getResponseCode();
  }
  
  let accounts = JSON.parse(Data);
  
  return accounts;
}

// Запрос на получение компаний.
function getCompaniesTinkoff(){ 
  let companies = [];
  const inns = [];
  
  for (let i = 0; i < inns.length; i++){
    let url = 'https://business.tinkoff.ru/openapi/api/v1/company/{inn}/info';
  
    let options = {
      headers: {'Authorization': 'Bearer ' + tokenOfAuthorizationTinkoff},
      method: 'GET',
      contentType: 'application/json',
    };
  
    let Data = UrlFetchApp.fetch(url, options);
  
    if (Data.getResponseCode() == 503) { // 503 ошибка - Сервер на технических работах.
      Browser.msgBox("Ошибка!", 'Сервер ModulBank проводит технические работы', Browser.Buttons.OK);
      writeError('Сервер ModulBank проводит технические работы');
      return Data.getResponseCode();
    }
    
    Data = JSON.parse(Data);
    
    let company = {
      companyName: companies[i].name,
      companyId: i,
      bankAccounts: [],
    };
    
    const accounts = getAccountsTinkoff(inns[i]); // Проверить
    
    for (let j = 0; j < accounts.length; j++){
      let account = {
        balance: accounts[j].currency, // Проверить
        number: accounts[j].accountNumber
      }
    }
  }
  
  return companies;
}

// Запрос на получение выписок из банка.(Одной компании)
function getPaymentsTinkoff(inn, accounts) { 
  const resObjs = [];
  
  for(let i = 0; i < accounts.length; i++){
    const url = 'https://business.tinkoff.ru/openapi/api/v1/company/' + inn + '/bankStatement'; 
    const fromDay = Utilities.formatDate(new Date(new Date() - amountDaysForAPI*oneDay), "GMT+3", "yyyy-MM-dd");
    
    let options = {
      headers: {'Authorization': 'Bearer ' + tokenOfAuthorizationTochka},
      method: 'GET',
      payload: JSON.stringify({
        accountNumber: accounts[i].accountNumber,
        from: fromDay,
      }),
      accept: 'application/json',
    };
    
    let payments =  UrlFetchApp.fetch(url, options);
    
    payments = JSON.parse(payments);
    
    for (let j = 0; j < payments.operation.length; j++){
      resObjs.push(payments.operation[j]);
    }
  }
  
  return resObjs;
}