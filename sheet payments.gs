//Вписывание выписок в лист "выписки".
function setPayments(sheet, payments, companies) {
  const lastRow = sheet.getMaxRows();
  
  let Data = [];
  Data.push([]);
  
  for (let i = 0; i < payments.length; i++){
    for (let j = 0; j < companies.length; j++){
      if (companies[j].companyId == payments[i].companyId){
        payments[i].companyName = companies[j].companyName;
        break;
      }
    }
  }
  
  payments = replaceValues(payments);
  
  for (let i = 0; i < payments.length; i++){
    
    Data[i].push(Utilities.formatDate(new Date(payments[i].executed), "GMT+3", "dd.MM.yyyy"));
    if (payments[i].category == "Debet"){
      Data[i].push(payments[i].contragentBankAccountNumber);
      Data[i].push(payments[i].amount);
      Data[i].push(payments[i].contragentName);
      Data[i].push(payments[i].companyName);
    }
    if (payments[i].category == "Credit"){
      Data[i].push(payments[i].bankAccountNumber);
      Data[i].push(-payments[i].amount);
      Data[i].push(payments[i].companyName);
      Data[i].push(payments[i].contragentName);
    }
    Data[i].push(payments[i].paymentPurpose);
    Data[i].push(payments[i].article);
    Data[i].push(Utilities.formatDate(new Date(payments[i].executed), "GMT+3", "MM"));
    if (payments[i].category == "Debet"){
      Data[i].push(payments[i].bankAccountNumber);
    }
    if (payments[i].category == "Credit"){
      Data[i].push(payments[i].contragentBankAccountNumber);
    }
    
    if (i+1 < payments.length) Data.push([]);
  }
  
  if(Data[0].length != 0) sheet.getRange(lastRow+1, 1, Data.length, Data[0].length).setValues(Data);
}
