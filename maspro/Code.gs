function getEmployeeSheet() {
    return getSheetByName(EMPLOYEE_SHEET_NAME);
}

function getDashboardSheet() {
    return getSheetByName(DASHBOARD_SHEET_NAME);
}

function getNumberOfEmployees() {
    var ss = getEmployeeSheet();
    var column = ss.getRange('A:A');
    var values = column.getValues();
    var ct = 0;
    while ( values[ct] && values[ct][0] != "" ) {
        ct++;
    }
    return ct;
}

function loadEmployeeData() {
    var numOfEmployees = getNumberOfEmployees();
    var ss = getEmployeeSheet();
    var dataRangeDef = "A2:D" + numOfEmployees;
    var data = ss.getRange(dataRangeDef).getValues();

    return data;
}

function computeCbuBlock2CheckDigit(fullAccount) {
    var fullAccountDigits = fullAccount.split("");
    var cbuBlock2Weights = "3971397139713".split("");

    var sum = 0;
    for (var i = 0; i < fullAccountDigits.length; i++) {
        sum += parseInt(fullAccountDigits[i],10) * parseInt(cbuBlock2Weights[i],10);
    }
  
    var mod10 = sum % 10;
    return mod10 != 0 ? 10 - mod10 : 0; 
}

function buildCbu(account) {
    var accParts = account.split("/");
    var accountWithoutBar = accParts[0] + accParts[1];
    var fullAccount = CBU_BLOCK_2_ACCNT_TYPE_AND_CURRENCY + lpad(accountWithoutBar, "0", 11);
    var cbuBlock2 = fullAccount + computeCbuBlock2CheckDigit(fullAccount);
    
    var cbu = CBU_ASIPRO_BLOCK1 + cbuBlock2
  
    return cbu;
}

function getDueDate() {
    var theoricDueDate = new Date(new Date().getTime() + 7*24*60*60*1000);
    var lastBusinessDay = lastBusinessDayOfMonth();
    return theoricDueDate > lastBusinessDay ? lastBusinessDay : theoricDueDate;
}

function buildIdent3Field() {
    var today = new Date();

    var nextYearNum = today.getFullYear();
    var nextMonthNum = (today.getMonth() + 1) % 12;
    if (nextMonthNum === 0) {
        nextYearNum++;
    }
  
    var firstDayNextMonth = new Date(nextYearNum, nextMonthNum, 1);
    return rpad("CUOTA " + Utilities.formatDate(firstDayNextMonth, "GMT-3", "MM/yyyy"), " ", 15);
}

function buildFileObject(data) {
    fileStr = "";
    var totalAmount = 0;
    var batchNumber = "000001";
    var movementCode = "1040";
    var presentationDate = Utilities.formatDate(new Date(), "GMT-3", "yyyyMMdd");
    var dueDate = Utilities.formatDate(getDueDate(), "GMT-3", "yyyyMMdd");
  
    for (var i = 0; i < data.length; i++) {
        var sequence = Utilities.formatString("%06d", i+1);
        var movementCode = "1040";
      
        var amount = data[i][COL_AMOUNT];
        var amountNoDecimalPoint = amount * 100.00;
        var amountStr = Utilities.formatString("%011d", amountNoDecimalPoint.toFixed(0));
        totalAmount += amount;
      
        var account = data[i][COL_ACCOUNT_NUM];
        var cbu = buildCbu(account);
      
        var ident2 = rpad(lpad(stripNonNumeric(data[i][COL_DOCUMENT_NUM]), "0", 8), " ", 22);
        var ident3 = buildIdent3Field();
      
        var rejectCode = filler(" ", 3);
        var idOriginalMsg = filler("0", 15);
        var rest = filler(" ", 94);
      
        var line = COELSA_REG_TYPE_DETAIL + batchNumber + sequence + movementCode + amountStr + cbu + ident2 + ident3;
        line += rejectCode + idOriginalMsg + rest;
      
        fileStr +=  line + "\r\n";
    }
  
    var totalAmountNoDecimalPoint = totalAmount * 100.00;
    var totalAmountStr = Utilities.formatString("%013d", totalAmountNoDecimalPoint.toFixed(0));
  
    var batch = "000001"; // Assume less than 10,000 trxs forever :D
    var previousDebitDate = filler("0", 8);
    var restHeader = filler(" ", 131);
  
    var batchHeaderLine = COELSA_REG_TYPE_BATCH_HEADER + batchNumber + movementCode + dueDate + Utilities.formatString("%06d", data.length);
    batchHeaderLine += totalAmountStr + COELSA_BATCH_HEADER_SERVICE + COELSA_BATCH_HEADER_TAX + batch + previousDebitDate + restHeader;
  
    var filler1 = filler("0", 6);
    var filler2 = filler(" ", 137);
    var numOfBatches = "000001";
  
    var initialHeaderLine = COELSA_REG_TYPE_HEADER + filler1 + presentationDate + numOfBatches + totalAmountStr;
    initialHeaderLine += COELSA_ASIPRO_COMPANY_CODE + rpad(COELSA_ASIPRO_COMPANY_DESC, " ", 20) + filler2;
  
    fileStr = initialHeaderLine + "\r\n" + batchHeaderLine + "\r\n" + fileStr; 
  
    var result = {
        fileContents:fileStr,
        numOfEmployees:data.length,
        totalAmount:totalAmount
    };
  
    return result;
}

function getWorkingFolder() {
    var iter = DriveApp.getRootFolder().getFoldersByName(WORKING_FOLDER_NAME);
    return iter.hasNext() ? iter.next() : DriveApp.getRootFolder().createFolder(WORKING_FOLDER_NAME);
}

function buildFileName() {
    var presentationPeriod = Utilities.formatDate(new Date(), "GMT-3", "ddMM");
    return COELSA_FILE_NAME_PREFIX + presentationPeriod + ".txt"; 
}

function updateDashboard(result) {
    var ss = getDashboardSheet();
    var cellLastGenerationDate = ss.getRange("E3");
    var cellNumOfEmployees = ss.getRange("F3");
    var cellTotalAmount = ss.getRange("G3");
  
    cellLastGenerationDate.setValue(new Date());
    cellNumOfEmployees.setValue(result.numOfEmployees);
    cellTotalAmount.setValue(result.totalAmount);
}

function generateFile() {
    var data = loadEmployeeData();
    var fileObject = buildFileObject(data);
    getWorkingFolder().createFile(buildFileName(), fileObject.fileContents);
  
    var result = {
        numOfEmployees:fileObject.numOfEmployees,
        totalAmount:fileObject.totalAmount
    };
  
    return result;
}

function init() {
    var spr = SpreadsheetApp.getActiveSpreadsheet();
    spr.setSpreadsheetTimeZone("America/Argentina/Buenos_Aires");
}

function main() {
    init();
    var result = generateFile();
    updateDashboard(result);
}


