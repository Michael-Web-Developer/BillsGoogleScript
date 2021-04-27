var id_table = '1dS8IyPnp_z9UzRJXUKtSDBYgy2XZTI_c9-1ZvNK7vDE';
let excelTableId = "1dS8IyPnp_z9UzRJXUKtSDBYgy2XZTI_c9-1ZvNK7vDE";

let actualFolderId = "1Mqi23hWmMyVkYZoKqdZEtomwh3VxZNZ7";
let archiveFolderId = "1ZnV8x_l51JS5I3SSUYtWhMp03889tbv6";
//let actualFolderId = "1Th7Uo-kYulxvTvKIVoB-0vf_YNG1_4Zz";
//let archiveFolderId = "19bwUZyJ1-BcJ6B96G68-WBNiYfAnC0pZ";

function parseDate(date) {
    let month = "0" + (date.getMonth() + 1);
    let day = "0" + date.getDate();
    return {
        year: date.getFullYear(),
        month: month.substring(month.length - 2),
        day: day.substring(day.length - 2)
    };
}

function billsFromBank() {
    var folder = DriveApp.getFolderById(actualFolderId);
    var archiveFolder = DriveApp.getFolderById(archiveFolderId);
    var sheetPayment = SpreadsheetApp.openById(id_table);
    var forExcelSpread = SpreadsheetApp.openById(excelTableId);
    // var excelDbSheet = forExcelSpread.getSheetByName("БД");
    // if (!excelDbSheet) excelDbSheet = forExcelSpread.insertSheet("БД");
    var excelWriteSheet = forExcelSpread.getSheetByName("выписки 2");
    if (!excelWriteSheet) excelWriteSheet = forExcelSpread.insertSheet("выписки 2");
    excelWriteSheet.getRange(1, 2, excelWriteSheet.getMaxRows(), 2).setNumberFormat("0.00");
    excelWriteSheet.getRange(1, 4, excelWriteSheet.getMaxRows(), 1).setNumberFormat("@");

    var dbSheet = sheetPayment.getSheetByName("БД");
    var dbMatches = dbSheet.getRange(3, 1, dbSheet.getLastRow() - 2, 4).getValues();


    let sheet = sheetPayment.getSheetByName('выписки');
    sheet.getRange(1, 3, sheet.getMaxRows(), 1).setNumberFormat("0.00");
    sheet.getRange(1, 8, sheet.getMaxRows(), 2).setNumberFormat("@");

    let paymentSheet = sheetPayment.getSheetByName("счета");
    if (!paymentSheet) {
        paymentSheet = sheetPayment.insertSheet("счета");
        paymentSheet.getRange(1, 1, paymentSheet.getMaxRows(), paymentSheet.getMaxColumns()).setNumberFormat("@");

        paymentSheet.getRange(1, 1, 1, 1).setValues([["Дата"]]);
        paymentSheet.setFrozenRows(1);
        paymentSheet.setRowHeight(1, 50);
        let style = SpreadsheetApp.newTextStyle()
            .setFontSize(10)
            .setBold(true)
            .setItalic(true)
            .build();
        paymentSheet.getRange(1, 1, 1, paymentSheet.getMaxColumns())
            .setHorizontalAlignment("center").setVerticalAlignment("middle").setTextStyle(style)
            .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

    }
    paymentSheet.getRange(2, 2, paymentSheet.getMaxRows() - 1, paymentSheet.getMaxColumns() - 1).setNumberFormat("0.00");


    let files = folder.getFiles();

    failLoop:
        while (files.hasNext()) {
            let file = files.next();
            let fileName = file.getName();
            if (fileName.includes(".xls")) {
                parseExcel(file);
                continue;
            }
            let text = file.getBlob().getDataAsString("windows-1251");
            if (!text.trim()) continue;


            let names = ["сумма", "плательщик1", "получатель1", "назначениеплатежа", "получательсчет"];

            let arr = [];
            let currentRow = 0;
            let firstName = null;

            let paymentAccount = null;
            /**
             * @type {null|Array}
             */
            let paymentDate = null;
            let startSum = null; // начальный и конечный остаток
            let endSum = null;
            let paymentIndex = null; // индекс счета в строке


            let insertIntoPayments = () => {
                if (startSum === null || endSum === null || !paymentDate || paymentIndex === null) return true;

                let dateIndex = -1;
                if (paymentSheet.getLastRow() > 1) {
                    dateIndex = paymentSheet.getRange(2, 1, paymentSheet.getLastRow() - 1, 1)
                        .getValues().findIndex(row => row[0] === paymentDate.join("."));
                }
                if (!~dateIndex) {
                    let count = paymentSheet.getLastColumn();
                    if (count % 2 === 0) count += 1;
                    let range = paymentSheet.getRange(2, 1, paymentSheet.getLastRow(), count);
                    let values = range.getValues();
                    let lastArr = (new Array(count)).fill("");
                    lastArr[0] = paymentDate.join(".");
                    lastArr[paymentIndex - 1] = startSum;
                    lastArr[paymentIndex] = endSum;
                    values[values.length - 1] = lastArr;
                    lastArr = null;

                    /// сортируем по дате
                    values = values.sort((current, row) => {
                        let currentDate = current[0].split(".");
                        let currentDateTime = (new Date(currentDate[2], currentDate[1] - 1, currentDate[0])).getTime();
                        let rowDate = row[0].split(".");
                        let rowDateTime = (new Date(rowDate[2], rowDate[1] - 1, rowDate[0])).getTime();
                        return currentDateTime - rowDateTime;
                    });

                    range.setValues(values);
                } else {
                    dateIndex += 2;
                    paymentSheet.getRange(dateIndex, paymentIndex, 1, 2).setValues([[startSum, endSum]]);
                }
            };


            let obj = {};
            for (let row of text.split("\n")) {
                if (!row) continue;

                let rowArr = row.split("=");
                let name = rowArr[0].toLowerCase().trim();

                if (name in obj) {
                    currentRow++;
                    arr.push([
                        obj["дата"],
                        paymentAccount,
                        obj["сумма"],
                        obj["плательщик1"],
                        obj["получатель1"],
                        obj["назначениеплатежа"],
                        obj["short"],
                        obj["дата"].split(".")[1],
                        obj["получательсчет"]
                    ]);
                    obj = {};
                }

                if (rowArr[1]) rowArr[1] = rowArr[1].trim();

                if (name === "расчсчет") {
                    paymentAccount = rowArr[1];

                    let add = () => {
                        paymentSheet.getRange(1, paymentSheet.getLastColumn() + 1, 1, 2)
                            .merge().setValues([[paymentAccount, paymentAccount]]);
                    };

                    if (paymentSheet.getLastColumn() < 2) add();
                    else {
                        let paymentsArr = paymentSheet.getRange(1, 2, 1, paymentSheet.getLastColumn() - 1).getValues()[0];
                        if (!paymentsArr.includes(paymentAccount)) add();
                    }
                    continue;
                }

                if (name === "датаконца") paymentDate = rowArr[1].split(".");

                if (name === "начальныйостаток") {
                    let date = new Date(paymentDate[2], paymentDate[1] - 1, paymentDate[0]);
                    date.setDate(date.getDate() - 1);
                    let parsedDate = parseDate(date);
                    let dateIndex = -1;
                    if (paymentSheet.getLastRow() > 1) {
                        let datesRange = paymentSheet.getRange(2, 1, paymentSheet.getLastRow() - 1, 1).getValues();
                        dateIndex = datesRange.findIndex(row => row[0] === parsedDate.day + "." + parsedDate.month + "." + parsedDate.year);
                    }

                    paymentIndex = -1;
                    if (paymentSheet.getLastColumn() > 1) {
                        paymentIndex = paymentSheet.getRange(1, 2, 1, paymentSheet.getLastColumn() - 1)
                            .getValues()[0].indexOf(paymentAccount);
                    }
                    paymentIndex += 2;
                    // если есть предыдущая дата
                    if (~dateIndex && paymentIndex > 1) {
                        let paymentsData = paymentSheet.getRange(dateIndex + 2, paymentIndex, 1, 2).getValues()[0];
                        if (+paymentsData[1] !== +rowArr[1]) {
                            /// если не совпадают суммы
                            SpreadsheetApp.getUi().alert(file.getName());
                            break failLoop;
                        }

                    }

                    startSum = +rowArr[1];
                    if (startSum !== null && endSum !== null) insertIntoPayments();
                }

                if (name === "конечныйостаток") {
                    endSum = +rowArr[1];
                    if (startSum !== null && endSum !== null) insertIntoPayments();
                }


                if (["датасписано", "датапоступило"].includes(name)) {
                    obj.type = name === "датасписано" ? "расход" : "приход";
                    obj["дата"] = rowArr[1];
                    if (obj.type === "расход" && "сумма" in obj) obj["сумма"] *= -1;

                    continue;
                }


                if (!names.includes(name)) continue;


                if (firstName === null) firstName = name;

                if (name === "сумма") {
                    rowArr[1] = +rowArr[1].replace(/[.].{0,}/, '');
                    if (obj.type === "расход") rowArr[1] *= -1;
                }

                if (name === "плательщик1" || name === "получатель1") {
                    let index = dbMatches.findIndex(row => rowArr[1].toLowerCase().includes(row[0].toLowerCase()) && row[0].trim().length > 0);

                    if (~index) rowArr[1] = dbMatches[index][1];
                }

                if (name === "назначениеплатежа") {
                    let index = dbMatches.findIndex(row => rowArr[1].toLowerCase().includes(row[2].toLowerCase()) && row[2].trim().length > 0);
                    let short = "";
                    if (~index) short = dbMatches[index][3];
                    obj[name] = rowArr[1];
                    obj["short"] = short;
                    continue;

                }


                obj[name] = rowArr[1];
            }
            if (Object.keys(obj).length > 0) {
                arr.push([obj["дата"], paymentAccount, obj["сумма"], obj["плательщик1"], obj["получатель1"], obj["назначениеплатежа"], obj["short"], obj["дата"].split(".")[1], obj["получательсчет"]]);
            }

            for (let i = 0; i < arr.length; i++) {
                if (arr[i].length !== 9) {
                    arr.splice(i, 1);
                    i--;
                }
            }

            if (arr.length > 0) {
                let range = sheet.getRange(sheet.getLastRow() + 1, 1, arr.length, arr[0].length);
                range.setValues(arr);
            }

            let currentDate = parseDate(new Date());
            let formattedDate = currentDate.day + "." + currentDate.month + "." + currentDate.year;
            let dateFolders = archiveFolder.getFoldersByName(formattedDate);
            let dateFolder;
            if (dateFolders.hasNext()) dateFolder = dateFolders.next();
            else dateFolder = archiveFolder.createFolder(formattedDate);
            file.makeCopy(dateFolder);
            folder.removeFile(file);
        }


}

function parseExcel(file) {
    var forExcelSpread = SpreadsheetApp.openById(excelTableId);
    var excelDbSheet = forExcelSpread.getSheetByName("БД");
    if (!excelDbSheet) excelDbSheet = forExcelSpread.insertSheet("БД");
    var excelWriteSheet = forExcelSpread.getSheetByName("выписки 2");
    if (!excelWriteSheet) excelWriteSheet = forExcelSpread.insertSheet("выписки 2");
    var excelDbRange = [];
    if (excelDbSheet.getLastRow() > 0) {
        excelDbRange = excelDbSheet.getRange(3, 5, excelDbSheet.getLastRow() - 2, 2).getValues();
    }

    var folder = DriveApp.getFolderById(actualFolderId);
    var archiveFolder = DriveApp.getFolderById(archiveFolderId);


    let blob = file.getBlob();
    let data = {
        title: file.getName().replace(".xlsx", "").replace(".xls", ""),
        key: file.getId(),
        parents: [{id: actualFolderId}]
    };
    let fileInfo = Drive.Files.insert(data, blob, {convert: true});

    let excelSheet = SpreadsheetApp.openById(fileInfo.id).getActiveSheet();
    let typeRange = excelSheet.getRange(1, 1, 1, 1).getValues();
    let type = typeRange[0][0].toLowerCase() === "номер перевода" ? 1 : 2;

    let arr = [];
    if (type === 1) {
        let range = excelSheet.getRange(3, 1, excelSheet.getLastRow() - 4, excelSheet.getLastColumn()).getValues();
        for (let row of range) {
            let paymentAccount = "";
            let fioIndex = excelDbRange.findIndex(r => r[1].toLowerCase() === row[37].toLowerCase());
            if (~fioIndex) paymentAccount = excelDbRange[fioIndex][0];
//            let fioMatches = dbMatches.findIndex(r => row[37].toLowerCase().includes(r[0].toLowerCase()));
//            if(~fioMatches) row[37] = dbMatches[fioMatches][1];
            let a = [row[1], +row[10], +row[16], String(paymentAccount), row[37]];
            arr.push(a);
        }

    } else {
        let range = excelSheet.getRange(11, 1, excelSheet.getLastRow() - 12, excelSheet.getLastColumn()).getValues();
        for (let row of range) {
            let fio = "";
            let paymentIndex = excelDbRange.findIndex(r => String(r[0]) === String(row[5]));
            if (~paymentIndex) {
                fio = excelDbRange[paymentIndex][1];
//                let fioMatches = dbMatches.findIndex(r => fio.toLowerCase().includes(r[0].toLowerCase()));
//                if (~fioMatches) fio = dbMatches[fioMatches][1];
            }
            let a = [row[3].trim().split(" ")[0], +row[6], +row[8], String(row[5]), fio];
            arr.push(a);
        }
    }

    if (arr.length > 0) {
        excelWriteSheet.getRange(excelWriteSheet.getLastRow() + 1, 1, arr.length, 5).setValues(arr);
    }


    let currentDate = parseDate(new Date());
    let formattedDate = currentDate.day + "." + currentDate.month + "." + currentDate.year;
    let dateFolders = archiveFolder.getFoldersByName(formattedDate);
    let dateFolder;
    if (dateFolders.hasNext()) dateFolder = dateFolders.next();
    else dateFolder = archiveFolder.createFolder(formattedDate);
    file.makeCopy(dateFolder);
    folder.removeFile(file);
    folder.removeFile(DriveApp.getFileById(fileInfo.id));
}