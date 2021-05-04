class StyleTable {
    constructor() {
        this.type_law = templateCol.indexOf(type_law);
        this.type_tender =templateCol.indexOf(type_tender);
        this.status = templateCol.indexOf(statusNameCol);
        this.spaceTender = templateCol.indexOf(spaceTender);
        this.extra_charge = templateCol.indexOf(extraCharge);
        this.order_col = templateCol.indexOf(orderNameCol);
        this.cost = templateCol.indexOf(cost);
        this.prime_cost = templateCol.indexOf(primeCost);
        this.typeDataValidationByRole = {
            'Ответственный':{
                status:{
                    col:this.status,
                    values: ['Согласовано на участие','Участие не целесообразно', 'Просчет рентабельности', 'Отправлена на согласование']
                },
                type_law:{
                    col:this.type_law,
                    values: ['44 ФЗ', '223 ФЗ']
                },
                type_tender:{
                    col:this.type_tender,
                    values: ['А', 'ЗК', 'ЗК', 'МЗ', 'ПР']
                },
              spaceTender:{
                col:this.spaceTender,
                values: ['РТС Тендер', 'Сбербанк-АСТ', 'Росэлторг', 'ЭТП Заказ РФ', 'ТЭК-ТОРГ', 'ЭТП ГПБ']
              }
            },
            'Администратор':{
                status:{
                    col:this.status,
                    values: ['Подготовка заявки', 'Участие в закупке', 'Участие не целесообразно', 'Подача заявки', 'Просчет рентабельности', 'Отправлена на согласование', 'Бухгалтеру', 'В архив', 'Закупка выиграна', 'Закупка проиграна', 'Заявка отклонена']
                },
                type_law:{
                    col:this.type_law,
                    values: ['44 ФЗ', '223 ФЗ']
                },
                type_tender:{
                    col:this.type_tender,
                    values: ['А', 'ЗК', 'ЗК', 'МЗ', 'ПР']
                },
                spaceTender:{
                   col:this.spaceTender,
                   values: ['РТС Тендер', 'Сбербанк-АСТ', 'Росэлторг', 'ЭТП Заказ РФ', 'ТЭК-ТОРГ', 'ЭТП ГПБ']
                }
            },
            'Менеджер':{
                status:{
                    col:this.status,
                    values: ['Подача заявки', 'Участие в закупке']
                },
                type_law:{
                    col:this.type_law,
                    values: ['44 ФЗ', '223 ФЗ']
                },
                type_tender:{
                    col:this.type_tender,
                    values: ['А', 'ЗК', 'ЗК', 'МЗ', 'ПР']
                },
                spaceTender:{
                   col:this.spaceTender,
                   values: ['РТС Тендер', 'Сбербанк-АСТ', 'Росэлторг', 'ЭТП Заказ РФ', 'ТЭК-ТОРГ', 'ЭТП ГПБ']
                }
            },
            'Бухгалтер':{
                status:{
                    col:this.status,
                    values: ['На оплату']
                },
                type_law:{
                    col:this.type_law,
                    values: ['44 ФЗ', '223 ФЗ']
                },
                type_tender:{
                    col:this.type_tender,
                    values: ['А', 'ЗК', 'ЗК', 'МЗ', 'ПР']
                },
                spaceTender:{
                   col:this.spaceTender,
                   values: ['РТС Тендер', 'Сбербанк-АСТ', 'Росэлторг', 'ЭТП Заказ РФ', 'ТЭК-ТОРГ', 'ЭТП ГПБ']
                }
            },
            'Сотрудник':{
                status:{
                    col:this.status,
                    values: ['Просчет рентабельности', 'Отправлена на согласование']
                },
                type_law:{
                    col:this.type_law,
                    values: ['44 ФЗ', '223 ФЗ']
                },
                type_tender:{
                    col:this.type_tender,
                    values: ['А', 'ЗК', 'ЗК', 'МЗ', 'ПР']
                },
                spaceTender:{
                   col:this.spaceTender,
                   values: ['РТС Тендер', 'Сбербанк-АСТ', 'Росэлторг', 'ЭТП Заказ РФ', 'ТЭК-ТОРГ', 'ЭТП ГПБ']
                }
            }
        }
    }

    styleForEvery() {
        var ssMails = SpreadsheetApp.openById(id_ss_Mails);
        var ssFinal = SpreadsheetApp.openById(id_ss_Final);

        var sheetStaff = ssMails.getSheetByName('Сотрудники');
        var sheetRespFromSSMails = ssMails.getSheetByName('Ответственные');

        var urlCriteria = SpreadsheetApp.newDataValidation().requireTextIsUrl().build();
        var style = SpreadsheetApp.newTextStyle()
            .setFontSize(9)
            .setUnderline(false)
            .setBold(false)
            .build();

        let ruleCheckbox =  SpreadsheetApp.newDataValidation().requireCheckbox().build();
        var arraySheets = [sheetStaff, sheetRespFromSSMails];
        for (var sheet in arraySheets) {
            var rangeSheetsAll = arraySheets[sheet].getRange('A2:B' + arraySheets[sheet].getLastRow()).getValues();
              
            for (var row in rangeSheetsAll) {
              if (rangeSheetsAll[row][1] == '' || rangeSheetsAll[row][0] == '') continue;


              var getSheetFromFinalTable = ssFinal.getSheetByName(rangeSheetsAll[row][1]);
              if (getSheetFromFinalTable == null) continue;
              
              
              
              //стили для текста
                for (let index of textType){
                    let getIndexCol = templateCol.indexOf(index);
                    getSheetFromFinalTable.getRange(3, getIndexCol + 1, getSheetFromFinalTable.getLastRow() + 10, 1)
                        .clear({formatOnly: true})
                        .clearDataValidations()
                        .setFontColor("black")
                }
              
              //Скрытие столбца с номером пункта для Сотрудников
              if(sheet === '0') {
                getSheetFromFinalTable.hideColumns(templateCol.indexOf(numberItemOrder) + 1);
              }
              
              //Валидация данных для разных ролей таблицы
                if(this.typeDataValidationByRole.hasOwnProperty(rangeSheetsAll[row][0]) || sheet === '0'){                    
                    let objectStaff;
                    sheet === '0' ? objectStaff = this.typeDataValidationByRole['Сотрудник'] : objectStaff = this.typeDataValidationByRole[rangeSheetsAll[row][0]];
                    for (let property in objectStaff){
                        let rule = SpreadsheetApp.newDataValidation().requireValueInList(objectStaff[property].values).build();
                        getSheetFromFinalTable.getRange(3,objectStaff[property].col+ 1, getSheetFromFinalTable.getLastRow(), 1).setDataValidation(rule);
                    }
                }
              
              
              
              //стили для денежнего формата
                for (let index of numberType){
                    let getIndexCol = templateCol.indexOf(index);
                    getSheetFromFinalTable.getRange(3, getIndexCol + 1, getSheetFromFinalTable.getLastRow() + 10, 1)
                        .setNumberFormat('0.00')
                        .setFontColor("black")
                }
              
              //стили для даты
                for (let index of date){
                    let getIndexCol = templateCol.indexOf(index);
                    getSheetFromFinalTable.getRange(3, getIndexCol + 1, getSheetFromFinalTable.getLastRow() + 10, 1)
                        .setFontColor("black")
                        .setNumberFormat('hh:mm dd.mm.yyy');
                }
              
              //стили для ссылок
                for (let index of hrefs){
                    let getIndexCol = templateCol.indexOf(index);
                    getSheetFromFinalTable.getRange(3, getIndexCol + 1, getSheetFromFinalTable.getLastRow() + 10, 1)
                        .setFontColor("blue")
                        .setDataValidation(urlCriteria);
                }

                if(sheet == '0'){
                    getSheetFromFinalTable.getRange(3,templateCol.length + 1, getSheetFromFinalTable.getLastRow(), (getSheetFromFinalTable.getLastColumn() - templateCol.length)).setDataValidation(ruleCheckbox);
                }

                getSheetFromFinalTable.getRange(3, 1, getSheetFromFinalTable.getLastRow(), getSheetFromFinalTable.getLastColumn())
                    .setTextStyle(style)
                    .setWrap(true)
                    .setBorder(true, true, true, true, true, true)
                    .setVerticalAlignment("middle")
                    .setHorizontalAlignment("center");
              
              let getRangeSheet = getSheetFromFinalTable.getRange(3, this.order_col + 1, getSheetFromFinalTable.getLastRow(), 1).getValues();
              for (let rowRange in getRangeSheet){
                if(getRangeSheet[rowRange][0] == '') continue;
                
                let primeCostCol = getCol(this.prime_cost + 1);
                let costCol = getCol(this.cost + 1);
                getSheetFromFinalTable.getRange(+rowRange+3, this.extra_charge + 1, 1, 1).setFormula('=IF(' + primeCostCol + (+rowRange+3) + '="";0;100-(' + primeCostCol + (+rowRange+3) + '/' + costCol + (+rowRange+3) + '*100))');
              }
            }
        }
    }

}


function setFormulaExtraCharge() {
    var ssMails = SpreadsheetApp.openById(id_ss_Mails);
    var ssFinal = SpreadsheetApp.openById(id_ss_Final);

    var sheetStaff = ssMails.getSheetByName('Сотрудники');
    var sheetRespFromSSMails = ssMails.getSheetByName('Ответственные');

    var arraySheets = [sheetStaff, sheetRespFromSSMails];
    for (var sheet in arraySheets) {
        var rangeSheetsAll = arraySheets[sheet].getRange('A2:B' + arraySheets[sheet].getLastRow()).getValues();
        for (var row in rangeSheetsAll) {
            if(rangeSheetsAll[row][0].toLowerCase() === "бухгалтер") continue;
            var getSheetFromFinalTable = ssFinal.getSheetByName(rangeSheetsAll[row][1]);
            if (getSheetFromFinalTable == null) continue;
            var rangeSheet = getSheetFromFinalTable.getRange('K3:K' + ssFinal.getSheetByName(rangeSheetsAll[row][1]).getLastRow()).getValues();
            for (var rowSheet in rangeSheet) {
                getSheetFromFinalTable.getRange('K' + (+rowSheet + 3)).setFormula('=100 - (I' + (+rowSheet + 3) + '/H' + (+rowSheet + 3) + '*100)');
            }
        }
    }
}


/**
 * Фон для новых зписей
 * @param rangeRow {GoogleAppsScript.Spreadsheet.Range} Строка в админ листе
 */
function styleForStatuses(rangeRow) {
    let stutusCol = templateCol.indexOf(statusNameCol);
    let values = rangeRow.getValues()[0];
    let color = values[stutusCol].toLowerCase() === "отправлена на согласование" ? "#00FF00" : "#FFFFFF";
    rangeRow.getCell(1, 2).setBackground(color);
}

function runStylesForEvery(){
  new StyleTable().styleForEvery();
}