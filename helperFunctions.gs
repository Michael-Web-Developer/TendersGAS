/**
* Получить буквенный индекс Столца, по числовому индексу
* return {string}
*/
function getCol(index) {
    var array_alp = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
    var last_index_array = array_alp.length;
    index--;
    if (index <= last_index_array) {
        return array_alp[index];
    } else {
        var col = '';
        var iter = 0;
        while (index > last_index_array) {
            col = array_alp[iter];
            index -= last_index_array;
            iter++;
        }
        col += array_alp[index];
        return col;
    }
}


/*function reverse() {
    let spread = SpreadsheetApp.openById(id_ss_Final);
    let respSheet = spread.getSheetByName("Пастушков И");
    let range = respSheet.getRange(3, 1, respSheet.getLastRow() - 2, respSheet.getLastColumn());
    let arr = range.getValues();
    range.setValues(arr.reverse());
}*/



/**
* Сформировать массив, по существующим заказам у переданного Листа.
* return {array}
*/
function getArrayOrdersExistResponsible(sheetResponse) {
    let indexColOrder = templateCol.indexOf(orderNameCol);
    var arrayOrders = [];
    var rangeOrders = sheetResponse.getRange(3,indexColOrder + 1, sheetResponse.getLastRow(), 1).getValues();
    for (var row in rangeOrders) {
        if (rangeOrders[row][0] != '') {
            arrayOrders.push(rangeOrders[row][0]);
        }
    }
    return arrayOrders;
}

/**
* Получить массив всех админов.
* return {array}
*/
function getArrayAdmins() {
    var ssMails = SpreadsheetApp.openById(id_ss_Mails);
    var ssFinal = SpreadsheetApp.openById(id_ss_Final);

    var admins = [];
    var admin = ssMails.getSheetByName(sheetNameResponse).getRange('A2:B' + ssMails.getSheetByName(sheetNameResponse).getLastRow()).getValues();
    for (var row in admin) {
        if (admin[row][0] == 'Администратор' && ssFinal.getSheetByName(admin[row][1])) {
            admins.push(admin[row][1]);
        }
    }
    return admins;
}

/**
 * Перемещает в архив запись
 * @param row {GoogleAppsScript.Spreadsheet.Range} Строка администратора
 */
function archive(row, oldValue) {
    /*let ui = SpreadsheetApp.getUi();
    let statuses = ['Подготовка заявки', 'Участие в закупке', 'Участие не целесообразно', 'Подача заявки', 'Просчет рентабельности', 'Отправлена на согласование', 'Бухгалтеру', 'Закупка проиграна', 'Закупка выиграна'];
    var response = ui.prompt('Введите предыдущий статус');  
    var status = response.getResponseText();  
    if(status == ''){
      return ui.alert('Нужно ввести статус');
    }  
    if(statuses.indexOf(status) == -1){
      return ui.alert('Вы ввели неверный статус');
    }*/
    
    
    let spread = SpreadsheetApp.openById(id_ss_Final);
    let stuff = getAllStuff();
    let indexColStatus = templateCol.indexOf(statusNameCol);
    let indexColOrder = templateCol.indexOf(orderNameCol);
    let archiveSheet = spread.getSheetByName("Архив");
    if (!archiveSheet) return;

    row[indexColStatus] = oldValue;
    Logger.log('Старое значение: ' + oldValue);
    for (let sheetName of stuff) {
        let sheet = spread.getSheetByName(sheetName);
        if (!sheet || sheet.getLastRow() === 2) continue;

        let rowIndex = searchRowByOrder(row[indexColOrder], sheet.getRange(3, 1, sheet.getLastRow() - 2, 2).getValues());
        if (rowIndex === null) continue;
        sheet.deleteRow(rowIndex);
    }

    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, 1, archiveSheet.getLastColumn()).setValues([row]);
}

/**
 * Возвращает номер строки по номеру закупки
 * @param order {Number}
 * @param values {Array[]} Массив значений всех строк, начиная с 3
 * @return {Number|Null}
 */
function searchRowByOrder(order, values) {
    let index = values.findIndex(row => String(order) === String(row[1]));
    if (!~index) return null;
    return index + 3;
}


/**
 * Возвращает всех ответственных
 * @return {String[]}
 */
function getArrayResponsibles() {
    let stuffSpread = SpreadsheetApp.openById(id_ss_Mails);
    let sheet = stuffSpread.getSheetByName(sheetNameResponse);
    let arr = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    return arr
        .filter(row => row[0].toLowerCase() === "ответственный")
        .map(row => row[1]);
}


/**
 * Возвращает массив с ролью сотрудник
 * @return {Array}
 */
function getArrayStuff() {
    let stuffSpread = SpreadsheetApp.openById(id_ss_Mails);
    let sheet = stuffSpread.getSheetByName(sheetNameStaff);
    let arr = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    return arr
        .filter(row => row[1])
        .map(row => row[1]);
}


/**
 * Возвращает массив с ролью бухгалтер
 * @return {Array}
 */
function getArrayAccountants() {
    let spread = SpreadsheetApp.openById(id_ss_Mails);
    let sheet = spread.getSheetByName(sheetNameResponse);
    let arr = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    return arr
        .filter(row => row[0].toLowerCase() === "бухгалтер")
        .map(row => row[1]);
}


/**
 * Возвращает всех сотрудников и ответственных
 * @returns {[string]}
 */
function getAllStuff() {
    let spread = SpreadsheetApp.openById(id_ss_Mails);
    let sheets = [sheetNameStaff, sheetNameResponse];
    let arr = [];
    for (let sheetName of sheets) {
        let sheet = spread.getSheetByName(sheetName);
        if (!sheet) continue;
        let values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
        arr = arr.concat(values.filter(row => row[0]).map(row => row[0]));
    }

    return arr;
}


/**
 * Поиск одинаковых записей
 * @param event
 */
function searchSame(event) {
    let spread = SpreadsheetApp.openById(id_ss_Final);
    let responsiblesArr = getArrayResponsibles();
    let colIndexStatus = templateCol.indexOf(statusNameCol);
    let range = event.range;
    let sheet = range.getSheet();
    if(!responsiblesArr.includes(sheet.getName())) return; // если не ответственный

    if (range.getColumn() !== 2) return;

    let order = range.getValues()[0][0];
    let stuff = getArrayStuff();

    let isFound = false;
    for (let sheetName of stuff) {
        let stuffSheet = spread.getSheetByName(sheetName);
        if (!stuffSheet) continue;

        let index = searchRowByOrder(+order, stuffSheet.getRange(
            3, 1,
            stuffSheet.getLastRow() - 2, stuffSheet.getLastColumn()).getValues()
        );
        if (index === null) continue;
        let rowArr = stuffSheet.getRange(index, 1, 1, stuffSheet.getLastColumn()).getValues()[0];

        // если статус не совпадает или стоит чекбокс
        if(rowArr[colIndexStatus].toLowerCase() !== "просчет рентабельности" || rowArr.includes(true)) continue;

        isFound = true;
        // stuffSheet.getRange(index, 2, 1, 1).setBackground("#FF0000");
    }

    sheet.getRange(range.getRow(), 2, 1, 1).setBackground("#FF0000");

}

/**
 * Добавляет валидацию данных, в необходимые листы
 * @param event
 */
function addValidation() {
    let admin = getArrayAdmins();
    let spread = SpreadsheetApp.openById(id_ss_Final);
    for (let adminName of admin) {
        let adminSheet = spread.getSheetByName(adminName);
        if (!adminSheet) continue;

        let range = adminSheet.getRange(3, 13, adminSheet.getMaxRows() - 2, 1);
        let currentValidation = range.getDataValidation();
        let rule = currentValidation.getCriteriaValues()[0];
        rule.push("Бухгалтеру", "В архив");
        range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(rule).build());
    }
}