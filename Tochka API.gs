// Запрос на получение компаний для определения названия компании по customer_code.
function getCompaniesTochka(){ 
  
  const url = 'https://enter.tochka.com/api/v1/organization/list';
  
  let options = {
    headers: {'Authorization': 'Bearer ' + tokenOfAuthorizationTochka},
    method: 'GET',
    accept: "application/json",
  };
  
  let Data = UrlFetchApp.fetch(url, options);
  
  if (Data.getResponseCode() == 503) { // 503 ошибка - Сервер на технических работах.
    Browser.msgBox("Ошибка!", 'Сервер Tochka проводит технические работы', Browser.Buttons.OK);
    writeError('Сервер Tochka проводит технические работы');
    return Data.getResponseCode();
  }
  
  Data = JSON.parse(Data);
  
  return Data;
}

// Запрос на получение request_id. Нужен для получения выписок.
function getRequestId(account){
  let request_id;
  
  let url = 'https://enter.tochka.com/api/v1/statement';
  
  const today = Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd");
  const fromDay = Utilities.formatDate(new Date(new Date() - amountDaysForAPI*oneDay), "GMT+3", "yyyy-MM-dd");
  
  let options = {
    headers: {'Authorization': 'Bearer ' + tokenOfAuthorizationTochka},
    method: 'POST',
    payload: JSON.stringify({
      "account_code": account.account_code,
      "bank_code": account.bank_code,
      "date_end": today,
      "date_start": fromDay
    }),
    contentType: 'application/json',
    accept: "application/json",
  };
  
  request_id = UrlFetchApp.fetch(url, options);
  
  request_id = JSON.parse(request_id);
  
  return request_id.request_id;
}

// Запрос на получение статуса запроса на получение выписок
function checkStatusRequest(request_id){
  url = 'https://enter.tochka.com/api/v1/statement/status/'+ request_id;
  
  let status = {status: ""};
  
  options = {
    headers: {'Authorization': 'Bearer ' + tokenOfAuthorizationTochka},
    method: 'GET',
    contentType: 'application/json',
  };
  
  while (status.status != "ready" && status.status != "Bad JSON" && status.status != "Operation not allowed") status = JSON.parse(UrlFetchApp.fetch(url, options));
}

// Возвращает выписки перестроенные по форме выписок из ModulBank.
function getConvertedPayments(payments){
  let convertedPayments = [];
  for (let i = 0; i < payments.length; i++){
    let payment_bank_system_id = payments[i].payment_bank_system_id; // ID платежного документа
    let counterparty_name = payments[i].counterparty_name; // Имя контрагента
    let counterparty_account_number = payments[i].counterparty_account_number; // Счет контрагента
    let operation_type;// Тип операции
    let payment_amount;// Сумма платежа
    let companyId = payments[i].companyId;
    if (payments[i].payment_amount > 0) {
      operation_type = "Debet";
      payment_amount = payments[i].payment_amount;
    }
    else {
      operation_type = "Credit";
      payment_amount = -payments[i].payment_amount;
    }
    let code = payments[i].code; // Счет
    let payment_purpose = payments[i].payment_purpose; // Назначение платежа
    let payment_date = payments[i].payment_date;
    convertedPayments.push({
      id: payment_bank_system_id,
      companyId: companyId,
      category: operation_type,
      contragentName: counterparty_name,
      contragentBankAccountNumber: counterparty_account_number,
      amount: payment_amount,
      bankAccountNumber: code,
      paymentPurpose: payment_purpose, 
      executed: payment_date,
    });
  }
  return convertedPayments;
}

// Запрос на получение выписок из банка.
function getPaymentsTochka(companies) { 
  const resObjs = [];
  const convertedCompanies = [];
  
  for (let i = 0; i < companies.length; i++){
    let company = {
      companyName: companies[i].full_name,
      companyId: companies[i].customer_code,
      bankAccounts: [],
    };
    convertedCompanies.push(company);
    for(let j = 0; j < companies[i].accounts.length; j++){
      // Получение request_id. Для получения payments 
      let request_id = getRequestId(companies[i].accounts[j]);
      
      //Проверка статуса
      checkStatusRequest(request_id);
      
      // Получение payments
      url = 'https://enter.tochka.com/api/v1/statement/result/'+ request_id;
      
      let payments;
      
      options = {
        headers: {'Authorization': 'Bearer ' + tokenOfAuthorizationTochka},
        method: 'GET',
        accept: 'application/json',
      };
      
      payments =  UrlFetchApp.fetch(url, options);
      
      payments = JSON.parse(payments);
      
      for (let k = 0 ; k < payments.payments.length; k++) {
        let payment = payments.payments[k];
        payment.code = companies[i].accounts[j].account_code; // Добавление номер счета в каждую выписку
        payment.companyId = companies[i].customer_code; // Чтобы не менять код в sheet payments
        resObjs.push(payments.payments[k]);
      }
      let account = {
        balance: payments.balance_closing,
        number: companies[i].accounts[j].account_code
      }
      convertedCompanies[i].bankAccounts.push(account);
    }
  }
  let result = {
    payments: getConvertedPayments(resObjs).sort(function(a, b){
      if (a.executed < b.executed) return -1;
      if (a.executed == b.executed) return 0;
      if (a.executed > b.executed) return 1;
    }),
    companies: convertedCompanies,
  }
  return result;
}

function test(){
  let companiesTochka = getCompaniesTochka();
  let DataTochka;
  
  if (companiesTochka != 503) DataTochka = getPaymentsTochka(companiesTochka.organizations);
  return;
}
