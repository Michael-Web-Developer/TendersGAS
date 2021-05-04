var DateDiff = {    
    inMin: function(d1, d2) {
        var t1 = d1.getTime();  
        var t2 = d2.getTime();
        
        return parseInt((t1 - t2)/(1000*60));
    }
}

function getListAllMails(){
  var lists = ['Сотрудники', 'Ответственные'];
  let statusMailsStaffs = {};
  var ss_mails = SpreadsheetApp.openById(id_ss_Mails);
  for (let item of lists){
    var getSheet = ss_mails.getSheetByName(item);
    if (getSheet == null) continue;
    
    if(item == 'Сотрудники'){
      var getRange = getSheet.getRange(rangeStaffInDb + getSheet.getLastRow()).getValues();
      for (var row of getRange){
        if(row[0] == '' || row[1] == '') continue;
        if(statusMailsStaffs['Сотрудники'] == null){
          statusMailsStaffs['Сотрудники'] = {};
          statusMailsStaffs['Сотрудники'][row[1]] = row[0];
        } else {
          statusMailsStaffs['Сотрудники'][row[1]] = row[0];
        }
      }
    } else {
      var getRange = getSheet.getRange('A2:C' + getSheet.getLastRow()).getValues();
      for (var row of getRange){
        if(row[0] == '' || row[1] == '' || row[2] == '') continue;
        if(statusMailsStaffs[row[0]] == null){
          statusMailsStaffs[row[0]] = {};
          statusMailsStaffs[row[0]][row[1]] = row[2];
        } else {
          statusMailsStaffs[row[0]][row[1]] = row[2];
        }
      }
    }    
  }
  return statusMailsStaffs;
}

function getSentOrders(){
   var ss_mails = SpreadsheetApp.openById(id_ss_Mails).getSheetByName('Отправленные');
   let arrayOrders = [];
  let getRange = ss_mails.getRange('A2:A' + ss_mails.getLastRow()).getValues();
  for (let row of getRange){
    if(row[0] == '') continue;
    arrayOrders.push(row[0]);
  }
  return arrayOrders;
}

function sendMail( name, row, email){
  let indexColStatus = templateCol.indexOf(statusNameCol);
  let indexColOrder = templateCol.indexOf(orderNameCol);
  let indexColDateDocs = templateCol.indexOf(dateDocs);
  let indexColDateTender = templateCol.indexOf(dateTender);
  let indexColNameTender = templateCol.indexOf(nameOrder);
  let indexColSpaceTender = templateCol.indexOf(spaceTender);
  let indexColHrefTender = templateCol.indexOf(hrefTender);
  let indexColCostTender = templateCol.indexOf(cost);
  let indexColExChargeTender = templateCol.indexOf(extraCharge);
  
  var header, text
    
  text  = '\n Документы для подачи Заявки по закупке № - '+row[indexColStatus].toString()+' - + - '+row[indexColNameTender]+' - готовы.'
  text += '\n Площадка - '+row[indexColSpaceTender]+' -'
  text += '\n Ссылка на конкурсную процедуру - '+row[indexColHrefTender]+' - '
  text += '\n НМЦК/Маржинальность - '+row[indexColCostTender]+ ' -/- '+row[indexColExChargeTender]+' - '

  text += '\n\n Напоминаю, что Заявка должна быть подана до  - '+formatDate(row[indexColDateDocs], 'full', 'time')+' -! '
  text += '\n Дата проведения торгов - '+formatDate(row[indexColDateTender], 'full', 'time')+' - '
  
  
  header = 'Уважаемый '+name+'!'
  MailApp.sendEmail(email, 'тема', header + text)  

  
}

function checkForSendMails(staff = undefined, order = undefined){
  let indexColStatus = templateCol.indexOf(statusNameCol);
  let indexColOrder = templateCol.indexOf(orderNameCol);
  let indexcolDateDocs = templateCol.indexOf(dateDocs);
  
  let work_ss = SpreadsheetApp.openById(id_ss_Final);
  let mailsTable = SpreadsheetApp.openById(id_ss_Mails);
  
  let listStaffData = getListAllMails();
  let listSentOrders = getSentOrders();
  
  let listStaffs = [];
  if(staff != null){
    listStaffs = [staff];
  } else {
    for(let staffItem in listStaffData['Сотрудники']){
      listStaffs.push(staffItem);
    }    
  }
  
  var curDate = new Date();
  for (let staffSheet of listStaffs){
    let getSheet = work_ss.getSheetByName(staffSheet);
    if(getSheet == null) continue;
    let getRange = getSheet.getRange(3, 1, getSheet.getLastRow(), getSheet.getLastColumn()).getValues();
    
    for (let row of getRange){
      if(listSentOrders.indexOf(row[indexColOrder]) != -1) continue;
      
      let arraySheetsResp = [];
      if(row[indexcolDateDocs] instanceof Date !== true) continue;
      let dtDiff = DateDiff.inMin(row[indexcolDateDocs], curDate);

      if (dtDiff > 0 && dtDiff <= 5*60 || staff != null && order == row[0]){
        let startColResp = templateCol.length + 1;
        let countResp = templateCol.length + 2 - getSheet.getLastColumn();
        
        for (let i = startColResp; i <= countResp; i++){
          if(typeof(row[i]) == 'boolean' && row[i]){
            arraySheetsResp.push(getSheet.getRange(2,i,1,1).getValue());
          }          
        }
        let mail = listStaffData['Сотрудники'][staffSheet];
        sendMail(staffSheet, row, mail);
        for (let resp of arraySheetsResp){
          mail = listStaffData['Ответственный'][resp];
          sendMail(resp, row, mail);
        }
        
        for(let admin in listStaffData['Администратор']){
          sendMail(admin, row, listStaffData['Администратор'][admin]);
        }
        mailsTable.getSheetByName('Отправленные').appendRow([row[indexColOrder], curDate]);        
      }
       
    }
  }  
}

function checkManagerStatus(){
  let indexColStatus = templateCol.indexOf(statusNameCol);
  let indexColOrder = templateCol.indexOf(orderNameCol);
  let nameStaffIndex = 0;
  
  let work_ss = SpreadsheetApp.openById(id_ss_Final);
  let listSentOrders = getListAllMails();
  for(let sheetName in listSentOrders['Менеджер']){
    let checkSheetMan = work_ss.getSheetByName(sheetName);
    if(checkSheetMan == null) continue;
    
    let getRange = checkSheetMan.getRange(3,1, checkSheetMan.getLastRow(), templateCol.length).getValues();
    for (let row of getRange){
      if(row[indexColStatus] == 'Подача заявки' && row[indexColOrder] != '' && row[nameStaffIndex] != ''){
        checkForSendMails(row[nameStaffIndex], row[indexColOrder]);
      }
    }
  }
}