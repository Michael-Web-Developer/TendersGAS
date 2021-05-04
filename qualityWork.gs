var rangeQualityInDB = 'A2:D';

function getObjectQuality(sheet) {
  let startDate = sheet.getRange(2,2).getValue();
  let endDate = sheet.getRange(2,3).getValue();
  
  if(startDate == '' || endDate == '' || startDate instanceof Date === false || endDate instanceof Date === false) return;
    
  let objectQuality = {};
  var ssMails = SpreadsheetApp.openById(id_ss_Mails);
  let getRangeQuality = ssMails.getSheetByName('Показатели').getRange(rangeQualityInDB + ssMails.getSheetByName('Показатели').getLastRow()).getValues();
  
  
   
  for (let row of getRangeQuality){
    if(row[0] == '') continue;         
    let checkDate = startDate <= row[2] && endDate >= row[2];
    if(!checkDate) continue;
    
    if(objectQuality[row[3]] == null){
      objectQuality[row[3]] = 
      {
        'PR' : 0,
        'UZ' : 0,
        'ZV' : 0,
        'PZ' : 0
      }
    }

    switch (row[1]) {
      case 'Просчет рентабельности':
        objectQuality[row[3]]['PR'] += 1;            
        break;
      case 'Участие в закупке':
        objectQuality[row[3]]['UZ'] += 1;
        break;
      case 'Закупка выиграна':
        objectQuality[row[3]]['ZV'] += 1;
        break;
      case 'Подача заявки':
        objectQuality[row[3]]['PZ'] += 1;
        break;
      default:
        continue;
    }
  }
  return objectQuality;
}


function setAllQuality(){
  var ssFinal = SpreadsheetApp.openById(id_ss_Final);  
  var sheetQual = ssFinal.getSheetByName('Показатели работы');
  
  let objectQuality = getObjectQuality(sheetQual);
  
  let templateQuality = ['№ п/п', 'ФИО', 'Тендеров в обработке', 'Подготовлено заявок', 'Закупок выиграно', 'Поданые заявки', 'Итого'];
  let finalArray = [templateQuality];
  
  let lastRow = sheetQual.getLastRow();
  if(lastRow < 4) lastRow = 4;
  sheetQual.getRange(4,1, lastRow, sheetQual.getLastColumn()).clear({commentsOnly:true, contentsOnly:true, formatOnly:true, validationsOnly:true, skipFilteredRows:true});
  
  let itemRow = 5;
  let item = 1;
  for (let staff in objectQuality){
    finalArray.push([
      item, staff, 
      objectQuality[staff]['PR'], 
      objectQuality[staff]['UZ'], 
      objectQuality[staff]['ZV'], 
      objectQuality[staff]['PZ'],
      '=(C' + itemRow +  '/20+D' + itemRow + '/10' + '+E' + itemRow + '/3)/3'
    ])
    itemRow++;
    item++;
  }
  
  let style = SpreadsheetApp.newTextStyle()
            .setFontSize(10)
            .setUnderline(false)
            .setBold(true)
            .build();
  
  sheetQual.getRange(4,1,finalArray.length, templateQuality.length)
  .setValues(finalArray)
  .setWrap(true)
  .setBorder(true, true, true, true, true, true)
  .setVerticalAlignment("middle")
  .setHorizontalAlignment("center");
  
  sheetQual.getRange(4,1,sheetQual.getLastRow(), 1).setTextStyle(style);
  sheetQual.getRange(4,1,1, sheetQual.getLastColumn()).setTextStyle(style);
  sheetQual.getRange(4,sheetQual.getLastColumn(),sheetQual.getLastRow(), 1).setTextStyle(style).setNumberFormat('0.00');
}

function addOrderToQualityDb(){
  var ssMails = SpreadsheetApp.openById(id_ss_Mails);
  var ssFinal = SpreadsheetApp.openById(id_ss_Final);
  
  let qualitySheetFromDb = ssMails.getSheetByName('Показатели');
  let rangeOrderCalendar = qualitySheetFromDb.getRange(rangeQualityInDB + qualitySheetFromDb.getLastRow()).getValues();
  
  let objectAllOrders = {};
  for(let row of rangeOrderCalendar){
    if(row[0] == '') continue;
    if(objectAllOrders[row[1]] == null){
      objectAllOrders[row[1]] = {};
      objectAllOrders[row[1]][row[0]] = row[1];
    } else {
      objectAllOrders[row[1]][row[0]] = row[1];
    }
    
  }
  
  let sheetStaff = ssMails.getSheetByName(sheetNameStaff);
  let rangeList = sheetStaff.getRange(rangeStaffInDb + sheetStaff.getLastRow()).getValues();
  for (let row of rangeList){
    if(row[0] == '' && row[1] == '') continue;
    
    let checkSheet = ssFinal.getSheetByName(row[1]);
    if(checkSheet == null) continue;
    
    let getRange = checkSheet.getRange(rangeFromFirstColToNote + checkSheet.getLastRow()).getValues();
    objectAllOrders = appendRowToQuality(getRange, row[1], qualitySheetFromDb, objectAllOrders)
  }
  
  let sheetResponse = ssMails.getSheetByName(sheetNameResponse);
  
  rangeList = sheetResponse.getRange(rangeStaffInDb + sheetStaff.getLastRow()).getValues();
  for (let row of rangeList){
    if(row[1] == '') continue;
    
    if(row[0] === 'Ответственный' || row[0] === 'Менеджер'){
      let checkSheet = ssFinal.getSheetByName(row[1]);
      if(checkSheet == null) continue;
      
      let getRange = checkSheet.getRange(rangeFromFirstColToNote + checkSheet.getLastRow()).getValues();
      
      if(row[0] === 'Менеджер'){
        objectAllOrders = appendRowToQualityByManager(getRange, row[1], qualitySheetFromDb, objectAllOrders)
        continue;
      }
      
      objectAllOrders = appendRowToQuality(getRange, row[1], qualitySheetFromDb, objectAllOrders);
    }    
  }
  
}

function appendRowToQuality(range, sheetName, qualitySheetFromDb, objectAllOrders){
  let indexColStatus = templateCol.indexOf(statusNameCol);
  let indexColOrder = templateCol.indexOf(orderNameCol);
  
  for (let rowItem of range){
    if(rowItem[indexColOrder] == '' || objectAllOrders.hasOwnProperty(rowItem[indexColStatus]) && objectAllOrders[rowItem[indexColStatus]].hasOwnProperty(rowItem[indexColOrder])) continue;
    if(rowItem[indexColStatus] == 'Участие в закупке' || rowItem[indexColStatus] == 'Закупка выиграна' || rowItem[indexColStatus] == 'Просчет рентабельности'){
      let date = new Date();
      let array_to_copy = [rowItem[indexColOrder], rowItem[indexColStatus], date, sheetName];
      qualitySheetFromDb.appendRow(array_to_copy);
      
      if(objectAllOrders.hasOwnProperty(rowItem[indexColStatus])){
        objectAllOrders[rowItem[indexColStatus]][rowItem[indexColOrder]] = rowItem[indexColStatus];
      } else {
        objectAllOrders[rowItem[indexColStatus]] = {};
        objectAllOrders[rowItem[indexColStatus]][rowItem[indexColOrder]] = rowItem[indexColStatus];
      }
    }            
  }
  return objectAllOrders;
}

function appendRowToQualityByManager(range, sheetName, qualitySheetFromDb, objectAllOrders){
  let indexColStatus = templateCol.indexOf(statusNameCol);
  let indexColOrder = templateCol.indexOf(orderNameCol);
  
  for (let rowItem of range){
    if(rowItem[indexColOrder] == '' || objectAllOrders.hasOwnProperty(rowItem[indexColStatus]) && objectAllOrders[rowItem[indexColStatus]].hasOwnProperty(rowItem[indexColOrder])) continue;
    if(rowItem[indexColStatus] == 'Подача заявки'){
      let date = new Date();
      let array_to_copy = [rowItem[indexColOrder], rowItem[indexColStatus], date, sheetName];
      qualitySheetFromDb.appendRow(array_to_copy);
      
      if(objectAllOrders.hasOwnProperty(rowItem[indexColStatus])){
        objectAllOrders[rowItem[indexColStatus]][rowItem[indexColOrder]] = rowItem[indexColStatus];
      } else {
        objectAllOrders[rowItem[indexColStatus]] = {};
        objectAllOrders[rowItem[indexColStatus]][rowItem[indexColOrder]] = rowItem[indexColStatus];
      }      
    }            
  }
  
  return objectAllOrders;
}

