function objectAllDateByStaff() {
  var ssMails = SpreadsheetApp.openById(id_ss_Mails);
  var ssFinal = SpreadsheetApp.openById(id_ss_Final);
  
  let objectDateByStaff = {};
  
  let calendarSheetFromDb = ssMails.getSheetByName('Календарь');
  let rangeOrderCalendar = calendarSheetFromDb.getRange('A2:D' + calendarSheetFromDb.getLastRow()).getValues();
  
  for (let row of rangeOrderCalendar){
    if(row[0] == '' || row[1] instanceof Date !== true || row[2] != 'Участие в закупке') continue;
    
    let date = Utilities.formatDate(row[1], "GMT", "dd.MM.yyyy");
    
    if(objectDateByStaff[date] == null){
      objectDateByStaff[date] = {};
      objectDateByStaff[date][row[3]] = 1;
      continue;
    }
      
    if(objectDateByStaff[date][row[3]] == null){
      objectDateByStaff[date][row[3]] = 1
    } else {
      objectDateByStaff[date][row[3]] += 1;
    }
    
  }   
  return objectDateByStaff;
}


function calendar(){
  var ssMails = SpreadsheetApp.openById(id_ss_Mails);
  var ssFinal = SpreadsheetApp.openById(id_ss_Final);
  
  let objectDAteBySteff = objectAllDateByStaff();
  let sheetCalendar = ssFinal.getSheetByName('Календарный план');
  let staffObject = {};
  let staffRange = sheetCalendar.getRange(1,2,1,sheetCalendar.getLastColumn()).getValues();
  
  for (let row of staffRange){
    for (let col in row){
      if(row[col] != ''){
        staffObject[row[col]] = (+col + 2);
      }
    }
  }
  
  let date = new Date();
  
  let rangeCalendar = sheetCalendar.getRange('A2:A' + sheetCalendar.getLastRow()).getValues();
  if(date.getDay() === 1){
    date = Utilities.formatDate(date, "GMT", "dd.MM.yyyy");    
    for (let row in rangeCalendar){
      if(formatDate(rangeCalendar[row][0], 'full') == date){
        sheetCalendar.getRange(2, 1, (+row + 1), sheetCalendar.getLastColumn()).setBackground('white');
        sheetCalendar.getRange((+row + 2), 1, 7, sheetCalendar.getLastColumn()).setBackground("#A4C2F4");
        break;
      }
    }
  }
  
  for(let dateRow in rangeCalendar){
    date = Utilities.formatDate(rangeCalendar[dateRow][0], "GMT", "dd.MM.yyyy");
    let objectDateOfStaff = objectDAteBySteff[date];
    if(objectDateOfStaff != null){
      for (let staffName in objectDateOfStaff){
        let getColStaffInCalendar = staffObject[staffName];
        if(getColStaffInCalendar != null){
          sheetCalendar.getRange((+dateRow + 2), getColStaffInCalendar, 1, 1).setValue(objectDateOfStaff[staffName]);
        }
      }
    }
  }
}

function activeCurrentDay(){
  var ssMails = SpreadsheetApp.openById(id_ss_Mails);
  var ssFinal = SpreadsheetApp.openById(id_ss_Final);
  
  let sheetCalendar = ssFinal.getSheetByName('Календарный план');
  let getRange = sheetCalendar.getRange('A2:A' + sheetCalendar.getLastRow()).getValues();
  let date = formatDate(new Date(), 'full');
  for (let row in getRange){
    var test = formatDate(getRange[row][0], 'full');
    if(date == formatDate(getRange[row][0], 'full')){
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Календарный план').getRange('A' + (+row + 2)).activate();
    }
  }
}

//Добавляем заявки в Базу Данных,для полсчета общего кол-ва завок на определенную дату
function addOrderForCalendar(){
  var ssMails = SpreadsheetApp.openById(id_ss_Mails);
  var ssFinal = SpreadsheetApp.openById(id_ss_Final);
  
  let calendarSheetFromDb = ssMails.getSheetByName('Календарь');
  let rangeOrderCalendar = calendarSheetFromDb.getRange('A2:A' + calendarSheetFromDb.getLastRow()).getValues();
  
  let objectAllOrders = {};
  for(let row of rangeOrderCalendar){
    if(row[0] == '') continue;
    objectAllOrders[row[0]] = true;
  }
  
  let sheetStaff = ssMails.getSheetByName('Сотрудники');
  let rangeList = sheetStaff.getRange('A2:B' + sheetStaff.getLastRow()).getValues();
  let objectDateByStaff = {};
  for (let row of rangeList){
    if(row[0] == '' && row[1] == '') continue;
    
    let checkSheet = ssFinal.getSheetByName(row[1]);
    if(checkSheet == null) continue;
    
    let getRangeDates = checkSheet.getRange(3, 1, checkSheet.getLastRow(), templateCol.length).getValues();
    appendRowToCalendar(getRangeDates, calendarSheetFromDb, objectAllOrders, row[1])
  }
    
  let responseStaff = ssMails.getSheetByName('Ответственные');
  rangeList = sheetStaff.getRange('A2:B' + sheetStaff.getLastRow()).getValues();
  objectDateByStaff = {};
  for (let row of rangeList){
    if(row[0] == '' && row[1] == '') continue;
    
    let checkSheet = ssFinal.getSheetByName(row[1]);
    if(checkSheet == null) continue;
    
    let getRangeDates = checkSheet.getRange(3, 1, checkSheet.getLastRow(), templateCol.length).getValues();
    appendRowToCalendar(getRangeDates, calendarSheetFromDb, objectAllOrders, row[1])
  }
}

function appendRowToCalendar(range, calendarSheetFromDb, objectAllOrders, nameStaff){
  let indexColStatus = templateCol.indexOf(statusNameCol);
  let indexColOrder = templateCol.indexOf(orderNameCol);
  let indexcolDateTender = templateCol.indexOf(dateTender);
  
  for (let rowItem of range){
      if(rowItem[indexColStatus] != 'Участие в закупке' || rowItem[indexcolDateTender] instanceof Date !== true || rowItem[indexColOrder] == '' || objectAllOrders[rowItem[indexColOrder]] != null) continue;      
      let array_to_copy = [rowItem[indexColOrder], rowItem[indexcolDateTender], rowItem[indexColStatus], nameStaff];
      calendarSheetFromDb.appendRow(array_to_copy);      
    }
}

function addCurrentCols(){
  var ssMails = SpreadsheetApp.openById(id_ss_Mails);
  var ssFinal = SpreadsheetApp.openById(id_ss_Final);
  
  let arrayStaffs = [];
  
  let getSheet = ssMails.getSheetByName('Сотрудники');
  let range = getSheet.getRange('A2:B' + getSheet.getLastRow()).getValues();
  for (let row of range){
    if(row[0] === '' || row[1] === '') continue;
    arrayStaffs.push(row[1]);    
  }
  
  getSheet = ssMails.getSheetByName('Ответственные');
  range = getSheet.getRange('A2:B' + getSheet.getLastRow()).getValues();
  for (let row of range){
    if(row[0] === '' || row[1] === '' || row[0] !== 'Ответственный') continue;
    arrayStaffs.push(row[1]);    
  }
  
  let getCalendar = ssFinal.getSheetByName('Календарный план');
  let rangeCols = getCalendar.getRange(1,2,1,getCalendar.getLastColumn()).getValues()[0];
  
  let arrayAddStaffCols = [];

  for(let staff of arrayStaffs){
    if(rangeCols.indexOf(staff) === -1) arrayAddStaffCols.push(staff);
  }
  
  if(arrayAddStaffCols.length === 0) return;
  var boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
  getCalendar.getRange(1,(getCalendar.getLastColumn() + 1),1, arrayAddStaffCols.length)
    .setValues([arrayAddStaffCols])
    .setTextStyle(boldStyle)
    .setBorder(true, true, true, true, true, true)
    .setBackgroundRGB(217, 217, 217)
    .setVerticalAlignment("middle")
    .setHorizontalAlignment("center")
    .setWrap(true);
}