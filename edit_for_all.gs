let nameSheetRoles = [sheetNameStaff,sheetNameResponse]

function edit_data() {  
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Введите номер/номера строк для выгрузки');
  var rowsUser = response.getResponseText();
  //Форматы ответа (Номер строки (1), Диапозон строк через :(1:4), Несколько диапозонов (1:4;6:7))
  if(rowsUser == ''){
    return ui.alert('Нужно ввести диапозон строк, примеры:6, 6:10, 6:10;14:15');
  }
  
  var startRow = []; 
  var endRow = []; 
  
  if(rowsUser.match(/;/)){
    var arraySplit = rowsUser.split(';');    
    for (var key in arraySplit){
      if(arraySplit[key].match(/:/)){
        var arraySplit2 = arraySplit[key].split(':');
        startRow.push(+arraySplit2[0]);
        endRow.push(+arraySplit2[1]);
      } else {
        startRow.push(+arraySplit[key]);
        endRow.push(+arraySplit[key]);
      }
    }  
  } else if(rowsUser.match(/:/)){
    var arraySplit2 = rowsUser.split(':');
    startRow.push(+arraySplit2[0]);
    endRow.push(+arraySplit2[1]);
  } else if (rowsUser.match(/^\d{1,}$/)){
    startRow.push(+rowsUser);
    endRow.push(+rowsUser);
  } else {
    return ui.alert('Неверный формат, попробуйте еще раз');
  }
  
  let data = get_data(startRow,endRow);
  let roles = get_roles();
  edit_for_all(data,roles);
}

//Для каждого листа-роли ищется номер заявки из массива выбранных строк, 
//при совпадении происходит замена строки на выбранную пользователем
function edit_for_all(data,roles){
  let columnStart = templateCol.indexOf(orderNameCol) + 1;
  let columnEnd = templateCol.length;  
  let indexColOrder = templateCol.indexOf(orderNameCol);
      
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  for (let valR of roles){
    let sheet = ss.getSheetByName(valR[0]);
    for (let valD of data){
      var found = sheet.getRange(2,indexColOrder + 1,sheet.getLastRow()).createTextFinder(valD[0][0]).matchCase(false).matchEntireCell(true).findAll();
      for (let valF of found){
        sheet.getRange(valF.getRow(), columnStart, 1, columnEnd-columnStart+1).setValues(valD);
      }
    }
  }
}

//Получение массива данных из выбранных строк
function get_data(startRow,endRow){
  let columnStart = templateCol.indexOf(orderNameCol) + 1;
  let columnEnd = templateCol.length;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let rows=[];
  let range;
  for (let i=0;i<startRow.length;i++){
      for (let j = startRow[i];j <= endRow[i]; j++){
        range = sheet.getRange(j,columnStart,1, columnEnd-columnStart+1);
        rows.push(range.getValues())
      }
  }
return rows;  
}

//Получение ролей из таблицы "Данные сотрудников" лист "Ответсвенные" 
function get_roles(){
  let ss = SpreadsheetApp.openById(id_ss_Mails);
  let sheet = ss.getSheetByName(nameSheetRoles[0]);
  let roles=sheet.getRange(2, 2, sheet.getLastRow()-1).getValues();
  
  for (let i=1;i<nameSheetRoles.length;i++){
    sheet = ss.getSheetByName(nameSheetRoles[i]);
    roles = roles.concat(sheet.getRange(2, 2, sheet.getLastRow()-1).getValues());
  }
  
  return roles.filter(role => role[0] != "");
}