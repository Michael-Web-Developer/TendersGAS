function favourite_row() {
  let rowsUser = getResponseUser();
  let sheet = SpreadsheetApp.getActiveSheet();
  let range = sheet.getRange(rowsUser, 1, 1, sheet.getLastColumn());
  sheet.moveRows(range, 3);
  sheet.getRange(3,1,1, sheet.getLastColumn()).setBackground('#FF8C00');
}

function getResponseUser(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Введите номер строки для закрепления');
  var rowsUser = response.getResponseText();
  //Форматы ответа (Номер строки (1))
  if(rowsUser == ''){
    return ui.alert('Нужно ввести строку в формате: 2, где 2 это номер строки');
  }
  if (rowsUser.match(/^\d{1,}$/)){
    rowsUser = +rowsUser;
  } else {
    return ui.alert('Неверный формат, попробуйте еще раз');
  }
  return rowsUser;
}

function remove_favourite_row(){
  let rowUser = getResponseUser();
  let sheet = SpreadsheetApp.getActiveSheet();
  let range = sheet.getRange(rowUser, 1, 1, sheet.getLastColumn()).setBackground('white');
  let lastRow = sheet.getLastRow() + 1;
  sheet.moveRows(range, lastRow);
}