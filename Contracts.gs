class Contracts {
    constructor(id_table_work, id_table_staff, sheetName) {
        this.rangeTemplate = 'A1:E16';
        this.template = ['Номер договора', 'Наименование объекта / название работ (кратко)', 'Дата заключения договора', 'Дата исполнения по договору / доп. соглашению	Номер процедуры	Заказчик по договору',	'Номер процедуры', 'Заказчик по договору', 'Сумма договора, тыс. руб.', 'Обеспечение',	'Себестоимость', 'Стоимость доставки', 'Чистая прибыль',	'Рентабельность',	'Менеджер',	'Гарантийные обязательства', 'Статус'];
        this.sheetContracts = SpreadsheetApp.openById(id_table_work).getSheetByName(sheetName);
        this.arrayOrdersRows = new OrdersRowStaff(id_table_staff, id_table_work).getOrdersByStaff('Администратор');
    }

    appendContract(){      
        let indexColStatusNameColInWorkSheet = templateCol.indexOf(statusNameCol);
        let indexColOrderInWorkSheet = templateCol.indexOf(orderNameCol);
        let indexColCostInWorkSheet = templateCol.indexOf(cost);
        let indexColPrimeCostInWorkSheet = templateCol.indexOf(primeCost);
        let indexColNameOrderCostInWorkSheet = templateCol.indexOf(nameOrder); 
        let indexColNameProtectContractInWorkSheet = templateCol.indexOf(protectContract);
      
        let indexColOrderInTemplateContract = this.template.indexOf('Номер договора');
        let indexColStatusInTemplateContract = this.template.indexOf('Статус');
        let indexColProfInTemplateContract = this.template.indexOf('Рентабельность');
        let indexColProffitInTemplateContract = this.template.indexOf('Чистая прибыль');
        let indexColExCharInTemplateContract = this.template.indexOf('Себестоимость');
      
      let objectMatchTemplateAndWorkSheet = {
        'Номер договора': indexColOrderInWorkSheet,
        'Наименование объекта / название работ (кратко)': indexColNameOrderCostInWorkSheet,
        'Сумма договора, тыс. руб.': indexColCostInWorkSheet,
        'Обеспечение': indexColNameProtectContractInWorkSheet,
        'Себестоимость': indexColPrimeCostInWorkSheet,
        'Менеджер': 0        
      }
    
        let getRangeContracts = this.sheetContracts.getRange(2, indexColOrderInTemplateContract + 1, this.sheetContracts.getLastRow(), 1).getValues();
      
        /*** В массиве для поиска статуса ***/
        let arrayExistOrders = [];
      
        for (let row of getRangeContracts){       
            if(row[0] == '') continue;
            arrayExistOrders.push(row[0]);
        }
        
      
        let requireParams = SpreadsheetApp.newDataValidation().requireValueInList(['Контракт заключен', 'Контракт исполнен', 'Акт подписан', 'Бухгалтеру', 'Оплачено']);
      
        let arrayFinal = [];
        for (let admin in this.arrayOrdersRows){
          for (let rowArray in this.arrayOrdersRows[admin]){
            let row = this.arrayOrdersRows[admin][rowArray];
            if(row[indexColStatusNameColInWorkSheet] !== 'Участие в закупке' || arrayExistOrders.indexOf(row[indexColOrderInWorkSheet]) !== -1) continue;
            
            
            let array = [];
            for (let elem of this.template){
              if(objectMatchTemplateAndWorkSheet.hasOwnProperty(elem)){
                array.push(row[objectMatchTemplateAndWorkSheet[elem]]);
              } else {
                array.push('');
              }
            }
            arrayFinal.push(array);
          }
        }   
        
        if(arrayFinal.length === 0) return;
        
        this.sheetContracts.getRange(this.sheetContracts.getLastRow() + 1, 1, arrayFinal.length, arrayFinal[0].length).setValues(arrayFinal);
        this.sheetContracts.getRange(2, indexColStatusInTemplateContract + 1, 1, 1).setDataValidation(requireParams);
      
        this.sheetContracts.getRange(2, 1, this.sheetContracts.getLastRow(), this.sheetContracts.getLastColumn())
        .setFontSize(9)
        .setWrap(true)
        .setBorder(true, true, true, true, true, true)
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center");
      
        let range = this.sheetContracts.getRange(2, indexColProfInTemplateContract + 1, this.sheetContracts.getLastRow(), 1).getValues();
      
      for (let rowRange in range){
        let primeCostCol = getCol(indexColExCharInTemplateContract + 1);
        let costCol = getCol(indexColProffitInTemplateContract + 1);
        this.sheetContracts.getRange(2, indexColProfInTemplateContract + 1, 1, 1).setFormula('=' + primeCostCol + (+rowRange+2) + '/' + costCol + (+rowRange+2)).setNumberFormat('%');
      }
    }
}

function runContracts(){
  new Contracts(id_ss_Final, id_ss_Mails, 'Лист контрактов').appendContract();
}
