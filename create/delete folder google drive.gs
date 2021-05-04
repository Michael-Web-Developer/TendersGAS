
function create_folder() 
{ 
  var sheet_staff = SpreadsheetApp.openById(id_ss_Mails);
  var sheet_final_table = SpreadsheetApp.openById(id_ss_Final);
  
  var folder = DriveApp.getFoldersByName('Тендеры');
  var checkFolder = folder.hasNext();
  
  if(checkFolder !== true){
    folder = DriveApp.createFolder('Тендеры')
  } else {
    folder = folder.next();
  } 
  
  let listSheets = [];
  var sheetStaff = sheet_staff.getSheetByName(sheetNameStaff);
  var rangeListSheets = sheetStaff.getRange('B2:B' + sheetStaff.getLastRow()).getValues();
  
  for (let row of rangeListSheets){
    if(row[0] != ''){
      listSheets.push(row[0]);
    }
  }
  
  var responsebleStaff = sheet_staff.getSheetByName(sheetNameResponse);
  var getRange = responsebleStaff.getRange(rangeStaffInDb + responsebleStaff.getLastRow()).getValues();
  
  for (let row of getRange){
    if(row[0] == 'Ответственный' && row[1] != ''){
      listSheets.push(row[1]);
    }
  }
  
  var folderDate = folder.getFoldersByName(formatDate(new Date));
  var folderIterator = folderDate.hasNext();
  if(folderIterator !== true){
    folderDate = folder.createFolder(formatDate(new Date));
  } else {
    folderDate = folderDate.next();
  }
  
  let indexColStatus = templateCol.indexOf(statusNameCol);
  let indexColOrder = templateCol.indexOf(orderNameCol);
  let indexColHrefGoogleDrive = templateCol.indexOf(hrefGoogleDrive);
  
  for (var rowSheet of listSheets){
    var getSheet = sheet_final_table.getSheetByName(rowSheet);
    if(getSheet == null) continue;
    
    var rangeOrders = getSheet.getRange(3,1,getSheet.getLastRow(), indexColHrefGoogleDrive + 1).getValues();
    for (var row in rangeOrders){
      if (rangeOrders[row][indexColOrder] != '' && rangeOrders[row][indexColStatus] == 'Просчет рентабельности' && rangeOrders[row][indexColHrefGoogleDrive] == ''){        
        folder = folderDate.getFoldersByName(rangeOrders[row][indexColOrder]);
        if(folder.hasNext() !== true){
          folder = folderDate.createFolder(rangeOrders[row][indexColOrder]);
          getSheet.getRange(+row + 3, indexColHrefGoogleDrive + 1, 1, 1).setValue(folder.getUrl());
        }        
      }            
    }    
  }
}

function delete_folder()
{
  var sheet_staff = SpreadsheetApp.openById(id_ss_Mails);
  var sheet_final_table = SpreadsheetApp.openById(id_ss_Final);
  
  var folder = DriveApp.getFoldersByName('Тендеры');
  var checkFolder = folder.hasNext();
  
  if(checkFolder !== true){
    folder = DriveApp.createFolder('Тендеры')
  } else {
    folder = folder.next();
  } 
  
  let indexColHrefGoogleDrive = templateCol.indexOf(hrefGoogleDrive);
  let indexColStatus = templateCol.indexOf(statusNameCol);
  let indexColOrder = templateCol.indexOf(orderNameCol);  
  let rangeListAdmin = sheet_staff.getSheetByName(sheetNameResponse).getRange(rangeStaffInDb + sheet_staff.getSheetByName(sheetNameResponse).getLastRow()).getValues();
  for (let row of rangeListAdmin){
    if (row[0] == 'Администратор' && row[1] != ''){
      let getSheet = sheet_final_table.getSheetByName(row[1]);
      
      if (getSheet == null) continue;
      
      let getRangeStatusAndListStaff = getSheet.getRange(3, 1, getSheet.getLastRow(), indexColHrefGoogleDrive + 1).getValues();
      for (let rowRange in getRangeStatusAndListStaff){
        if(getRangeStatusAndListStaff[rowRange][indexColStatus] == 'Участие не целесообразно' && getRangeStatusAndListStaff[rowRange][indexColOrder] != ''){          
          var date = new Date();
          let folderDate = folder.getFoldersByName(formatDate(date));
          let checkFolderDate = folderDate.hasNext();
          if(checkFolderDate){
            folderDate = folderDate.next();
            let folderIterator = folderDate.getFoldersByName(getRangeStatusAndListStaff[rowRange][indexColOrder]);
            if(folderIterator.hasNext()){
              var folder = folderIterator.next();
            } else {
              date = new Date(date.setMonth(date.getMonth() - 1));
              folderDate = folder.getFoldersByName(formatDate(date));
              checkFolderDate = folderDate.hasNext();
              if(checkFolderDate){
                folderDate = folderDate.next();
                folderIterator = folderDate.getFoldersByName(getRangeStatusAndListStaff[rowRange][indexColOrder]);
                if(folderIterator.hasNext()){
                  var folder = folderIterator.next();
                } else {
                  continue;
                }
              } else {
                continue
              }
            } 
          }
                   
          let ui = SpreadsheetApp.getUi();
          let response = ui.prompt('Вы действительно хотите удалить папку - ' + getRangeStatusAndListStaff[rowRange][indexColOrder], ui.ButtonSet.YES_NO);
          
          if (response.getSelectedButton() == ui.Button.YES) {            
            folderDate.removeFolder(folder);
            getSheet.getRange((+rowRange + 3), indexColHrefGoogleDrive + 1, 1, 1).setValue('');
          }
        }
      }
    }
  }  
}

