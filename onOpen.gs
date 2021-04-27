function onOpen(e) {
    SpreadsheetApp.getUi()
        .createMenu('Доп. функции')
        .addItem('Проверить выписки', 'billsFromBank')
        .addItem('Обновить выписки через API', 'main')
        .addToUi();
}