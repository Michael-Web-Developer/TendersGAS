class OrdersRowStaff {
    constructor(id_table_staff, id_table_work) {
        this.ss_staff = SpreadsheetApp.openById(id_table_staff);
        this.ss_table_work = SpreadsheetApp.openById(id_table_work);            
        this.rangeDataStaff = "A2:B";
        this.array_staff_sheet = ['Сотрудники', 'Ответственные'];
    }

    getObjectTypeStaff(){
        let objectType = {};
        for (let sheetName of this.array_staff_sheet){
            let getSheet = this.ss_staff.getSheetByName(sheetName);
            let getRange = getSheet.getRange(this.rangeDataStaff + getSheet.getLastRow()).getValues();
            for (let row of getRange){
                if (row[0] === '' || row[1] === '') continue;

                if (sheetName === 'Сотрудники'){
                    objectType['Сотрудники'] == null ? objectType['Сотрудники'] = [row[1]] : objectType['Сотрудники'].push(row[1]);
                } else {
                    objectType[row[0]] == null ? objectType[row[0]] = [row[1]] : objectType[row[0]].push(row[1]);
                }
            }
        }
        return objectType;
    }

    getOrdersByStaff(){
        let objectStaffs = this.getObjectTypeStaff();
        let objectRowEveryStaff = {};
        let listTypeStaffs = [];
        if(arguments.length === 0) {
            for(let staff in objectStaffs) {
                listTypeStaffs.push(staff);
            }
        } else {
            for (let i = 0; i < arguments.length; i++){
                listTypeStaffs.push(arguments[i]);
            }
        }

        for (let typeStaff of listTypeStaffs){
            let getSheetNameByStaff = objectStaffs[typeStaff];
            for (let staff of getSheetNameByStaff){
                let getSheet = this.ss_table_work.getSheetByName(staff);
                if(getSheet == null) continue;

                let getRange = getSheet.getRange(3,1,getSheet.getLastRow(),getSheet.getLastColumn()).getValues();
                for (let row in getRange){
                    let currentRow = (+row + 3).toString();
                    objectRowEveryStaff[staff] == null ? objectRowEveryStaff[staff] = {currentRow:getRange[row]} : objectRowEveryStaff[staff][currentRow] = getRange[row];
                }
            }
        }
        return objectRowEveryStaff;
    }
}
