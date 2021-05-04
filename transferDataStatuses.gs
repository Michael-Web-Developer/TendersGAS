var id_ss_Mails = '1b0afNKBGhjlvLtfyXzEQ6tQF2rBVjrCrGQQQaKFgY-0';
var id_ss_Final = '1LHPpQ_aZoBX0fokOtVf35TKRVN2l3vWgeDZVg6iM8BM';
var rangeFromOrdersToStatus = 'B3:M';
var rangeFromFirstColToNote = 'A3:O';

var sheetNameStaff = 'Сотрудники';
var sheetNameResponse = 'Ответственные';
var rangeStaffInDb = 'A2:B';
var templateCol = ['№ п/п', 'Номер закупки', 'Тип закона', 'Вид закупки', 'Наименование', 'Ссылка на конкурсную процедуру', 'Площадка', 'Дата подачи документов', 'Дата торга', 'НМЦК', 'Себестоимость', 'Маржинальность', 'Рекомендованная цена', 'Обеспечение контракта', 'Итоговая цена', 'Статус *', 'Ссылка на гугл диске', 'Примечание'];
var orderNameCol = 'Номер закупки';
var statusNameCol = 'Статус *';
var hrefGoogleDrive = 'Ссылка на гугл диске';
var hrefTender = 'Ссылка на конкурсную процедуру';
var type_law = 'Тип закона';
var type_tender = 'Вид закупки';
var extraCharge = 'Маржинальность';
var cost = 'НМЦК';
var primeCost = 'Себестоимость';
var nameOrder = 'Наименование';
var protectContract = 'Обеспечение контракта';
var dateTender = 'Дата торга';
var dateDocs = 'Дата подачи документов';
var spaceTender = 'Площадка';
var numberItemOrder = '№ п/п';

/*
* Переменные для стилей
*
*/
var textType = ['Тип закона', 'Вид закупки', 'Наименование', 'Площадка', 'Рекомендованная цена', 'Статус *', 'Примечание', 'Номер закупки', 'Итоговая цена'];
var numberType = ['НМЦК', 'Обеспечение контракта', 'Себестоимость'];
var hrefs = ['Ссылка на конкурсную процедуру', 'Ссылка на гугл диске'];
var procent = ['Маржинальность'];
var date = ['Дата подачи документов', 'Дата торга'];



class TransferDataByStatus{
    constructor(id_work_table, id_db_table) {
        this.work_table = SpreadsheetApp.openById(id_work_table);
        this.db_table = SpreadsheetApp.openById(id_db_table);
        this.indexColOrder = templateCol.indexOf(orderNameCol);
        this.indexColStatus = templateCol.indexOf(statusNameCol);
        this.lastIndexInTemplate = templateCol.length - 1;
        this.indexColForNameSheet = 0;
    }

    /**
     * Функция отправляет заявки от рабочего персонала
     * к необходимым лицам на согласование
     */
    requestFromStaff(){

        var sheetStaff = this.db_table.getSheetByName(sheetNameStaff);

        var admins = getArrayAdmins();

        //Получаем массив с сотрудников, которые заполняют номера закупок.
        var rangeMails = sheetStaff.getRange(rangeStaffInDb + sheetStaff.getLastRow()).getValues();
        var arrayExistOrders = {};
        let listSheetNameStaff = [];

        for (var row in rangeMails) {
            if (rangeMails[row][0] != '' && rangeMails[row][1] != '') {
                var checkTable = this.work_table.getSheetByName(rangeMails[row][1]);
                if (checkTable == null) {
                    continue;
                }
                var dataOrders = checkTable.getSheetValues(3, this.indexColOrder + 1, checkTable.getLastRow(), 1);
                arrayExistOrders[rangeMails[row][1]] = [];
                for (let rowData of dataOrders) {
                    if(rowData[this.indexColOrder] == '') continue;
                    arrayExistOrders[rangeMails[row][1]].push(rowData[this.indexColOrder]);
                }
                listSheetNameStaff.push(rangeMails[row][1]);
            }
        }

        for (let row of listSheetNameStaff) {
            let checkTable = this.work_table.getSheetByName(row);

            //Берем range из листа текущего сотрудника, который вносит данный по закупкам
            let rangeTableStaff = checkTable.getRange(3,1,checkTable.getLastRow(), checkTable.getLastColumn()).getValues();

            for (let rowStaff in rangeTableStaff) {

                //Поиск текущего номера закупки у других сотрудников
                let statusExist = false;
                let checkExistStaffWithOrders = arrayExistOrders.hasOwnProperty(row);
                if (checkExistStaffWithOrders){
                    let indexInArrayOrders = arrayExistOrders[row].indexOf(rangeTableStaff[rowStaff][1]);
                    if (indexInArrayOrders !== -1) {
                        if (arrayExistOrders[row][indexInArrayOrders] == '') continue;
                        statusExist = true;
                    }
                }

                //Если такой номер существует, то помечаем ячейку красного цвета.
                if (statusExist) {
                    checkTable.getRange((+rowStaff + 3), this.indexColOrder + 1, 1, 1).setBackground('red');
                    continue;
                } else {
                    checkTable.getRange((+rowStaff + 3), this.indexColOrder + 1, 1, 1).setBackground('white')
                }

                //Преверяем каждую строку, на пустоту и то что сотрудник изменил статус на Отправлено на согласование
                if (rangeTableStaff[rowStaff][this.indexColOrder] != '' && rangeTableStaff[rowStaff][this.indexColStatus] == 'Отправлена на согласование') {
                    let sheetName = checkTable.getSheetName();
                    let arraySheetsResp = [];
                    for (let col in rangeTableStaff[rowStaff]) {

                        //Ищем, каким Ответсвтенным отправлено на согласование.
                        if (typeof (rangeTableStaff[rowStaff][col]) == 'boolean' && rangeTableStaff[rowStaff][col]) {
                            arraySheetsResp.push(checkTable.getRange(getCol(+col + 1) + '2').getValue());
                        }

                    }


                    if (arraySheetsResp.length === 0) continue;

                    //Дублируем данные ответсвенным.
                    this.copyOrderToResponse(arraySheetsResp, rangeTableStaff[rowStaff]);

                    //Ищем листы с Администраторами и дублируем, те же данные.
                    this.copyOrderToAdmin(getArrayAdmins(), rangeTableStaff[rowStaff], sheetName);
                }
            }
        }
    }

    /**
     * Функция копирует строку в листы с ответсвенными
     */
    copyOrderToResponse(arrayResponse, rowData, sheetNameStaff = undefined){
        for (let itemResp of arrayResponse) {
            let sheetResp = this.work_table.getSheetByName(itemResp);
            if (sheetResp == null) continue;

            //Делаем проверку, что текущего номера закупке нет у текущего Ответсвенного лица.
            if (getArrayOrdersExistResponsible(sheetResp).indexOf(rowData[this.indexColOrder]) != -1) continue;

            let lastRowResp = sheetResp.getLastRow() + 1;

            let rangeResponse = sheetResp.getRange(3,this.indexColOrder + 1,lastRowResp,1).getValues();

            //Находим первую пустую строку у ответсвенного лица и вставляем туда строку с информазицей по закупке
            for (let rowResponse in rangeResponse) {
                if (rangeResponse[rowResponse][this.indexColOrder - 1] == '') {
                    let dataOrderRow = rowData.slice(0, this.lastIndexInTemplate);
                    let row = sheetResp.getRange((+rowResponse + 3),1,1, dataOrderRow.length);
                    row.setValues([dataOrderRow]);
                    styleForStatuses(row);
                    break;
                }
            }
        }
    }

    /**
     * Функция копирует строку в листы с админами
     */
    copyOrderToAdmin(arrayAdmins, rowData, sheetNameStaff = undefined){
        for (let rowAdmin in arrayAdmins) {
            let adminSheet = this.work_table.getSheetByName(arrayAdmins[rowAdmin]);

            if (adminSheet == null) continue;

            if (getArrayOrdersExistResponsible(adminSheet).indexOf(rowData[this.indexColOrder]) != -1) continue;

            adminSheet.insertRowBefore(3);
            rowData[this.indexColForNameSheet] = sheetNameStaff;
            let arrayStaff = rowData.slice(0, this.lastIndexInTemplate);
            adminSheet.getRange(3,1,1,arrayStaff.length).setValues([arrayStaff]);
        }
    }

    /**
     * Функция переносит статусы с ответственных к админам
     */
    editActionToAdminByResponsible(){
        let admins = getArrayAdmins();

        for (let admin of admins) {
            let sheetAdmin = this.work_table.getSheetByName(admin);

            if (sheetAdmin == null) continue;

            let indexStartResponseCol = templateCol.length + 1;
            let responsibles = {};
            let rangeRespons = sheetAdmin.getRange(2,indexStartResponseCol, 1, sheetAdmin.getLastColumn()).getValues();

            for (let row in rangeRespons) {
                for (let col in rangeRespons[row]) {
                    if (rangeRespons[row][col] == '') continue;
                    responsibles[rangeRespons[row][col]] = indexStartResponseCol + (+col);
                }
            }

            let statusByRespons = {};
            for (let respons in responsibles) {
                let sheetRespons = this.work_table.getSheetByName(respons);
              
                if (sheetRespons == null) continue;

                rangeRespons = sheetRespons.getRange(3,1,sheetRespons.getLastRow(), this.indexColStatus + 1).getValues();
                for (let row in rangeRespons) {

                    let order = rangeRespons[row][this.indexColOrder];
                    let status = rangeRespons[row][this.indexColStatus];
                    if (order == '') break;

                    if (statusByRespons[order] == null) {
                        statusByRespons[order] = {};
                        statusByRespons[order][respons] = status;
                    } else {
                        statusByRespons[order][respons] = status;
                    }
                    styleForStatuses(sheetRespons.getRange(+row + 3, 1, 1, sheetRespons.getLastColumn()));
                }
            }

            let rangeOrders = sheetAdmin.getRange(3, this.indexColOrder + 1, sheetAdmin.getLastRow(), 1).getValues();
            for (let row in rangeOrders) {
                let orderFromObject = statusByRespons[rangeOrders[row][0]];
                if (orderFromObject != null) {
                    for (let respons in orderFromObject) {
                        let col = responsibles[respons];
                        sheetAdmin.getRange((+row + 3), col, 1,1).setValue(orderFromObject[respons]);
                    }
                }
            }
        }
    }

    /**
     * Функция переносит заказы от ответсвенных к администраторам, которые они сами нашли
     */
    addOrderToAdminByResponsible(){

        let sheetStaff = this.db_table.getSheetByName(sheetNameResponse);

        let admins = getArrayAdmins();
        let arrayExistOrders = {};
        for (let admin of admins){
            let sheet = this.work_table.getSheetByName(admin);
            let range = sheet.getRange(3, this.indexColOrder + 1, sheet.getLastRow() ,1).getValues();
            for(let row of range){
                if(row[0] === '') continue;
                arrayExistOrders[row[0]] = true;
            }
        }

        //Получаем массив с ответсвенными, которые заполняют номера закупок.
        let rangeRespons = sheetStaff.getRange(rangeStaffInDb + sheetStaff.getLastRow()).getValues();

        for (var row of rangeRespons) {
            if (row[0] == '' || row[1] == '' || row[0] != 'Ответственный') continue;

            let checkTable = this.work_table.getSheetByName(row[1]);
            if (checkTable == null) {
                continue;
            }

            let rangeOrdersResponse = checkTable.getRange(3,1,checkTable.getLastRow(), checkTable.getLastColumn()).getValues();
            for (let rowOrder of rangeOrdersResponse){
                if(rowOrder[this.indexColOrder] == '' || rowOrder[this.indexColStatus] != 'Отправлена на согласование' || arrayExistOrders[rowOrder[this.indexColOrder]]) continue;

                for (let admin of admins){
                    let adminSheet = this.work_table.getSheetByName(admin);
                    adminSheet.insertRowBefore(3);
                    rowOrder[this.indexColForNameSheet] = row[1];
                    let arrayOrder = rowOrder.slice(0, this.lastIndexInTemplate);
                    adminSheet.getRange(3,1,1,arrayOrder.length).setValues([arrayOrder]);
                }
            }
        }
    }

    /**
     * Функция проставляет статусы у всех участников, которые поставил администратор.
     * Также создает строки у нужных сотрудников с заявками
     */
    editActionAdminToAll() {
        let sheetStaff = this.db_table.getSheetByName(sheetNameStaff);
        let sheetRespons = this.db_table.getSheetByName(sheetNameResponse);

        let listSheets = [];

        let rangeStaff = sheetStaff.getRange(rangeStaffInDb + sheetStaff.getLastRow()).getValues();

        for (let row in rangeStaff) {
            if (rangeStaff[row][1] == '') {
                continue;
            } else {
                listSheets.push({status: rangeStaff[row][0], sheet: rangeStaff[row][1]});
            }
        }

        let rangeRespons = sheetRespons.getRange(rangeStaffInDb + sheetStaff.getLastRow()).getValues();
        for (let row in rangeRespons) {
            if (rangeRespons[row][1] == '') continue;
            listSheets.push({status: rangeRespons[row][0], sheet: rangeRespons[row][1]});
        }

        let admins = getArrayAdmins();
        let objectOrderAction = {};
        let readyToMoveAccountant = {};
        let readyToMoveManeger = {};

        for (let admin in admins) {
            let sheetAdmin = this.work_table.getSheetByName(admins[admin]);

            if (sheetAdmin == null) continue;

            let orderActionRangeObject = sheetAdmin.getRange(rangeFromFirstColToNote + sheetAdmin.getLastRow());
            let rangeOrderAction = orderActionRangeObject.getValues();
            for (let row in rangeOrderAction) {

                let order = rangeOrderAction[row][this.indexColOrder];
                let status = rangeOrderAction[row][this.indexColStatus];

                if (order == '' || status == '') continue;

                /*if (status.toLowerCase() === "в архив") {
                    archive(sheetAdmin.getRange(+row + 3, 1, 1, sheetAdmin.getLastColumn()));
                    return;
                }*/

                styleForStatuses(orderActionRangeObject.offset(+row, 0, 1));
                if (status == 'Отправлена на согласование') continue;

                if (status === 'Подготовка заявки') {
                    readyToMoveManeger[order] = {status: status, row: rangeOrderAction[row]};
                    continue;
                }

                if (status === "Бухгалтеру") {
                    readyToMoveAccountant[order] = {status: status, row: rangeOrderAction[row]};
                    continue;
                }
                objectOrderAction[order] = {status: status, row: rangeOrderAction[row]};
            }
        }

        for (let sheet in listSheets) {
            let getSheet = this.work_table.getSheetByName(listSheets[sheet].sheet);

            if (getSheet == null) continue;

            let state = listSheets[sheet].status == 'Менеджер';
            let isAccountant = listSheets[sheet].status === "Бухгалтер";

            let lastRow = getSheet.getLastRow();

            if (lastRow == 2) lastRow = 3;

            let rangeObject = getSheet.getRange(3,this.indexColOrder + 1,getSheet.getLastRow(), 1);
            let getRange = rangeObject.getValues();
            for (let row in getRange) {
                let checkExistStatus = objectOrderAction[getRange[row][0]];

                if (checkExistStatus != null || state || isAccountant) {
                    if (state && checkExistStatus != null) {
                        getSheet.getRange((+row + 3), this.indexColStatus + 1, 1, 1).setValue(checkExistStatus.status);
                    }
                    if (state) {
                        let checkOrder = readyToMoveManeger[getRange[row][0]];
                        if (checkOrder != null) {
                            delete readyToMoveManeger[getRange[row][0]];
                            continue;
                        }
                        continue;
                    }
                    if (isAccountant) {
                        let checkOrder = readyToMoveAccountant[getRange[row][0]];
                        if (checkOrder != null) {
                            delete readyToMoveAccountant[getRange[row][0]];
                        }
                    }
                    if (checkExistStatus !== null && checkExistStatus !== undefined) {
                        getSheet.getRange((+row + 3), this.indexColStatus + 1, 1, 1).setValue(checkExistStatus.status);

                        if (listSheets[sheet].status.toLowerCase() === "ответственный") {
                            styleForStatuses(getSheet.getRange(+row + 3, 1, 1, getSheet.getLastColumn()));
                        }
                    }
                }

            }

            if (state) {
                for (let orderRow in readyToMoveManeger) {
                    getSheet.appendRow(readyToMoveManeger[orderRow].row);
                }
            }

            if (isAccountant) {
                for (let row of Object.values(readyToMoveAccountant)) {
                    getSheet.appendRow(row.row);
                }
            }

        }

    }

    /**
     * Функция меняет статус у всех участников, которые изменил мендеджер
     */
    editStatusFromManagerToAll() {
        let sheetStaff = this.db_table.getSheetByName(sheetNameStaff);
        let sheetRespFromSSMails = this.db_table.getSheetByName(sheetNameResponse);
        let range = sheetRespFromSSMails.getRange(rangeStaffInDb + sheetRespFromSSMails.getLastRow()).getValues();
        let orderStatusReadyRequest = {};

        for (let row in range) {
            if (range[row][0] == 'Менеджер' && range[row][1] != '') {
                let sheetManager = this.work_table.getSheetByName(range[row][1]);

                if (sheetManager == null) continue;

                let rangeOrderAndStatus = sheetManager.getRange(3,1,sheetManager.getLastRow(), this.indexColStatus + 1).getValues();

                for (let rowManager in rangeOrderAndStatus) {
                    if (rangeOrderAndStatus[rowManager][this.indexColOrder] != '' && rangeOrderAndStatus[rowManager][this.indexColStatus] == 'Подача заявки') {
                        orderStatusReadyRequest[rangeOrderAndStatus[rowManager][this.indexColOrder]] = rangeOrderAndStatus[rowManager][this.indexColStatus];
                    }
                }
            }
        }

        let arraySheets = [sheetStaff, sheetRespFromSSMails];
        for (let row in arraySheets) {
            let getRangeSheets = arraySheets[row].getRange(rangeStaffInDb + arraySheets[row].getLastRow()).getValues();

            for (let rowSheet in getRangeSheets) {
                if (getRangeSheets[rowSheet][1] == '' || getRangeSheets[rowSheet][0] == 'Менеджер') continue;

                let sheetInFinalTable = this.work_table.getSheetByName(getRangeSheets[rowSheet][1]);
                if (sheetInFinalTable == null) continue;

                let rangeOrders = sheetInFinalTable.getRange(3,this.indexColOrder + 1, sheetInFinalTable.getLastRow(), 1).getValues();
                for (let rowOrder in rangeOrders) {
                    let checkStatus = orderStatusReadyRequest[rangeOrders[rowOrder][0]];
                    if (checkStatus != null) {
                        sheetInFinalTable.getRange((+rowOrder + 3), this.indexColStatus + 1, 1, 1).setValue(checkStatus);
                    }
                }
            }
        }
    }
}



function onEditByAccountant(e) {
    let range = e.range;
    let accountants = getArrayAccountants();
    let accountantSheet = range.getSheet();
    let indexColStatus = templateCol.indexOf(statusNameCol);
    let indexColOrder = templateCol.indexOf(orderNameCol);
    if (!accountants.includes(accountantSheet.getName())) return;

    if (range.getColumn() !== (indexColStatus + 1) || range.getRow() < 3) return;

    let status = range.getValue();
    let order = accountantSheet.getRange(range.getRow(), indexColOrder + 1, 1, 1).getValues()[0][0];
    if (!order || !status) return;

    let stuff = getAllStuff();
    let spread = SpreadsheetApp.openById(id_ss_Final);
    for (let sheetName of stuff) {
        let sheet = spread.getSheetByName(sheetName);
        if (!sheet || accountantSheet.getName() === sheet.getName()) continue;

        let index = searchRowByOrder(order, sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn()).getValues());
        if (index === null) continue;
        let range = sheet.getRange(index, 1, 1, sheet.getLastColumn());
        let arr = range.getValues();
        arr[0][indexColStatus] = status;
        range.setValues(arr);
    }
}

function onEditArchive(e){
  let admins = new OrdersRowStaff(id_ss_Mails, id_ss_Final).getObjectTypeStaff()['Администратор'];
  if(admins.indexOf(e.source.getActiveSheet().getName()) === -1 || e.range.getValue().toLowerCase() !== "в архив") return;  
  archive(e.source.getActiveSheet().getRange(e.range.getRow(), 1, 1, e.source.getLastColumn()).getValues()[0], e.oldValue);
}

function runTransferStatuses(){
  let objectTranferData = new TransferDataByStatus(id_ss_Final, id_ss_Mails);  
  objectTranferData.requestFromStaff();
  objectTranferData.editActionToAdminByResponsible();
  objectTranferData.addOrderToAdminByResponsible();
  objectTranferData.editActionAdminToAll();
  objectTranferData.editStatusFromManagerToAll();  
}

function testStatuses (){
  new TransferDataByStatus(id_ss_Final, id_ss_Mails).editStatusFromManagerToAll();
}

function testtest(){
  let ss_work = SpreadsheetApp.openById(id_ss_Final);
  let ss_test = SpreadsheetApp.openById('1RMIrRLmy-eQ_iS1UZdHbOVCiYea823c9CfI10QCTJzI');
  let array_sheet_name_staff = getAllStuff();
  let objectOrderStaffRow = {};
  
  for (let staffSheetName of array_sheet_name_staff){
    let getSheet = ss_work.getSheetByName(staffSheetName);
    
    let getRange = getSheet.getRange(3,2,getSheet.getLastRow(),1).getValues();
    objectOrderStaffRow[staffSheetName] = {};
    
    for (let row in getRange){
      if(getRange[row][0] == '') continue;
      
      objectOrderStaffRow[staffSheetName][getRange[row]] = +row + 3;
    }
  }
  
  for (let staffSheetName of array_sheet_name_staff){
    let getSheet = ss_test.getSheetByName(staffSheetName);
    let sheetWork = ss_work.getSheetByName(staffSheetName);
    let getRange = getSheet.getRange(3,2,getSheet.getLastRow(),getSheet.getLastColumn()).getValues();
    
    for (let row in getRange){
      if(getRange[row][0] == '' || getRange[row][7] == '') continue;
      
      if(!objectOrderStaffRow[staffSheetName].hasOwnProperty(getRange[row][0])) continue;
         
      sheetWork.getRange(objectOrderStaffRow[staffSheetName][getRange[row][0]], 11, 1, 1).setValue(getRange[row][7]);
      Logger.log(staffSheetName + ' ' + objectOrderStaffRow[staffSheetName][getRange[row][0]] + ' ' + getRange[row][7]);
    }
  }
  
}