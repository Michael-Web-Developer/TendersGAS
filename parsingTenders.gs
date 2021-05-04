class ParsingTenders {
    constructor(event){
        this.event = event;
        this.spreadsheet = event.source;
        this.sheetName;
        this.listSheetsStaff = getAllStuff();
        this.indexColHrefTender = templateCol.indexOf(hrefTender);
        this.indexColOrder = templateCol.indexOf(orderNameCol);
        this.responseHTML;
    }

    parsingTenders(){
        if(!this.checkNewTenders()) return;
        if(this.checkAlreadyExistHref()){
            this.event.range.setBackground('red');
            return;
        }
        this.event.range.setBackground('white');
        let typeTender = this.event.range.getValue().match(/[\/]223[\/]|[\/]epz[\/]|[\/]ea44[\/]/);
        if(typeTender == undefined) return;
      
        let response;
        switch (typeTender[0]) {
            case '/223/':
                return;
                break;
            case '/epz/':
                response = this.parsingHTML('epz');
                break;
            case '/ea44/':
                response = this.parsingHTML('ea44');
                break;

        }
      
        if(response['type'] === 'error') return;
      
        let indexColLaw =  templateCol.indexOf(type_law) + 1;
        let indexColOrder = templateCol.indexOf(orderNameCol) + 1;
        let indexColPrice = templateCol.indexOf(cost) + 1;
        let indexColNameContract = templateCol.indexOf(nameOrder) + 1;
        let indexColDateDocs = templateCol.indexOf(dateDocs) + 1;
        let indexColSpace = templateCol.indexOf(spaceTender) + 1;
        let indexColDateTender = templateCol.indexOf(dateTender) + 1;
        let indexColSupport = templateCol.indexOf(protectContract) + 1;
        //let indexColSupportPrice = templateCol.indexOf(dateTender) + 1;
      
        let sheet = this.spreadsheet.getSheetByName(this.sheetName);
      
        sheet.getRange(this.event.range.getRow(), indexColLaw, 1).setValue(response['law']);
        sheet.getRange(this.event.range.getRow(), indexColOrder, 1).setValue(response['contract_number']);
        sheet.getRange(this.event.range.getRow(), indexColPrice, 1).setValue(response['start-price']);
        sheet.getRange(this.event.range.getRow(), indexColDateDocs, 1).setValue(response['name_contract']);
        sheet.getRange(this.event.range.getRow(), indexColPrice, 1).setValue(response['indexColDateDocs']);
        sheet.getRange(this.event.range.getRow(), indexColSpace, 1).setValue(response['type_space']);
        sheet.getRange(this.event.range.getRow(), indexColDateTender, 1).setValue(response['date_end_tender']); 
        sheet.getRange(this.event.range.getRow(), indexColSupport, 1).setValue(response['indexColSupport']);
    }

    checkNewTenders(){
        Logger.log(this.event.range.getValue());
        if(this.event.range.getColumn() !== this.indexColHrefTender + 1) return false;
        if(this.event.range.getValue().match(/^(http|https)/) == undefined) return false;
        return true;
    }

    checkAlreadyExistHref(){
        let ss = SpreadsheetApp.getActive();
        for (let sheetName of this.listSheetsStaff){
            let checkSheet = ss.getSheetByName(sheetName);
            if(checkSheet == null) continue;
            let range = checkSheet.getRange(3,this.indexColHrefTender + 1, checkSheet.getLastRow(), 1).getValues();
            for(let row in range){
                if(this.event.range.getValue() === range[row][0] && this.event.range.getRow() === (+row + 3)) this.sheetName = sheetName;
                if(this.event.range.getValue() === range[row][0] && this.event.range.getRow() != (+row + 3)) return true;
            }
        }
        return false;
    }

    parsingHTML(typeHref){
        if(typeHref === 'epz' || typeHref === 'ea44'){
            this.parsingEa44AndEpz(typeHref, this.event.range.getValue());
        }

    }

    parsingEa44AndEpz(type, href){
        Logger.log('https://websitewizard.ru/test/parsing-tenders.php?type-tender=' + type + '&href=' + href);
        let response = UrlFetchApp.fetch('https://websitewizard.ru/test/parsing-tenders.php?type-tender=' + type + '&href=' + href);
        let result = response.getContentText();
      Logger.log(result['type']);
        return JSON.parse(result);
               
    }
}


function runParsingTendrer(e){
  new ParsingTenders(e).parsingTenders();
}


function testParsing(){
  let response = UrlFetchApp.fetch('https://websitewizard.ru/test/parsing-tenders.php?type-tender=ea44&href=https://zakupki.gov.ru/epz/order/notice/ea44/view/common-info.html?regNumber=0366200035620002154');
  let result = response.getContentText();
  let object = JSON.parse(result);
  let any = 'efwfefw';

}