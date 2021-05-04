function onOpen(e) {
    SpreadsheetApp.getUi()
        .createMenu('Доп. функции')
        .addItem('Проверить выписки', 'runUpdateBillsFromDrive')
        .addItem('Обновить выписки через API', 'main')
        .addToUi();
}