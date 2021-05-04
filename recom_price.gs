
/**
* Вывод сайд бара
*/
function openSideBar() {
  var html = HtmlService.createTemplateFromFile('sideBarForPrice');
  html.listOrders = getListOrders();
  let template = html.evaluate();
  SpreadsheetApp.getUi().showSidebar(template)
}


/**
* Вывод модального окна
*/
function modelDialog(){
  var html = HtmlService.createTemplateFromFile('sideBarForPrice');
  html.listOrders = getListOrders();
  let template = html.evaluate();
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(template, 'Расчет рек. цены');
}

/**
* Получение списка заказов на текущем листе
* return {object}
*/
function getListOrders(){
  let ss = SpreadsheetApp.getActiveSheet();
  let objectNameCols = getColumnIndexByName({'номер закупки':''}, ss);
  let rangeOrders = ss.getRange(3,objectNameCols['номер закупки'] + 1, ss.getLastRow(), 1).getValues();
  let objectOrders = {};
  for(let row in rangeOrders){
    if(objectOrders[rangeOrders[row]] == null || rangeOrders[row] != ''){
      objectOrders[rangeOrders[row]] = +row+3;
    }
  }
  return objectOrders;
}

/**
* Получение значение из обпределной строки
* input(array)
* return {object}
*/
function getDataOrder(order_row){
  let ss = SpreadsheetApp.getActiveSheet();
  let objectNameCols = getColumnIndexByName({'нмцк':'', 'себестоимость':''}, ss);
  let getRange = ss.getRange(order_row,1,ss.getLastRow(), ss.getLastColumn()).getValues()[0];
  objectNameCols['нмцк'] = getRange[objectNameCols['нмцк']];
  objectNameCols['себестоимость'] = getRange[objectNameCols['себестоимость']];
  return objectNameCols
}

/**
* Получение объекта с полями и их индексом строки
* input(object, objectSheet)
* return {object}
*/
function getColumnIndexByName(objectCol, sheet){
  let rangeCols = sheet.getRange(2,1,sheet.getLastRow(), sheet.getLastColumn()).getValues()[0];
  for (let col in rangeCols){
    let nameCol = rangeCols[col].toLowerCase();
    if(objectCol[nameCol] != null) objectCol[nameCol] = +col;
  }
  return objectCol
}

/**
* Расчет рек. цены и занесение ее в ячейку.
* input(object, int, float)
*/
function getPrice(objectData, row, coefPrice = 3.5){
  
  let price = coefPrice * objectData['time']/100*objectData['price']+objectData['price'];
  let profitContr = (100-100*(+objectData['price'])/(+objectData['nmck'])).toFixed(2);
  let profitInv = (12/(+objectData['time'])*profitContr).toFixed(2);
  
  if(profitContr>=10 && profitContr <100){
    profitContr = profitContr/10;
  }
  if(profitInv>=10 && profitInv <100){
    profitInv = profitInv/10;
  }
  
  if(profitContr === 100){
    profitContr = profitContr/100;
  }
  if(profitInv === 100){
    profitInv = profitInv/100;
  }
  
  let coefProfit = (profitInv*0.5+profitContr*0.3+(+objectData['complication'])).toFixed(2);
  
  let result = 'Реком. цена: ' + price + "\n" + 'Коэф. выгод.: ' + coefProfit;
  let ss = SpreadsheetApp.getActiveSheet();
  let colObject = getColumnIndexByName({"рекомендованная цена":''}, ss);
  ss.getRange(row, colObject['рекомендованная цена'] + 1).setValue(result);
  return;
}