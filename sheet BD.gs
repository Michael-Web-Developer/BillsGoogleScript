function replaceValues(payments) {
  const sheetPayment = SpreadsheetApp.openById(tableId);
  let dbSheet = sheetPayment.getSheetByName("БД");
  let dbMatches = dbSheet.getRange(3, 1, dbSheet.getLastRow() - 2, 4).getValues();
  
  for (let i = 0; i < payments.length; i++){
    for (let j = 2; j < dbMatches.length; j++){
      if (payments[i].contragentName.toLowerCase().indexOf(dbMatches[j][0].toLowerCase()) > -1 && dbMatches[j][0] !== "") payments[i].contragentName = dbMatches[j][1];
      if (payments[i].companyName.toLowerCase().indexOf(dbMatches[j][0].toLowerCase()) > -1 && dbMatches[j][0] !== "") payments[i].companyName = dbMatches[j][1];
      if (payments[i].paymentPurpose.toLowerCase().indexOf(dbMatches[j][2].toLowerCase()) > -1 && dbMatches[j][0] !== "") payments[i].article = dbMatches[j][3];
    }
  }
  
  return payments;
}
