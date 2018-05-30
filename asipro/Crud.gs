function sortEmployees() {
    var numOfEmployees = getNumberOfEmployees();
    var ss = getEmployeeSheet();
    var dataRangeDef = "A2:" + LAST_COL + numOfEmployees;
    ss.getRange(dataRangeDef).sort(1);
}

function createEmployee(emp) {
    var ss = getEmployeeSheet();
    ss.insertRowAfter(2);
  
    var cellDest = ss.getRange(LAST_COL + "3");
    var cellSrc  = ss.getRange(LAST_COL + "2");
    cellSrc.copyTo(cellDest);
  
    var cellFullName = ss.getRange("A3");
    var cellDocument = ss.getRange("B3");
    var cellAccount  = ss.getRange("C3");
    var cellAmount   = ss.getRange("D3");
  
    cellFullName.setValue((emp.lastname + ", " + emp.name).toUpperCase());
    cellDocument.setValue(emp.document.toUpperCase());
    cellAccount.setValue(emp.account);
    cellAmount.setValue(emp.amount);
  
    sortEmployees();
}

function showSidebarForNewEmployee() {
    SpreadsheetApp.setActiveSheet(getEmployeeSheet());
    var html = HtmlService.createHtmlOutputFromFile('coelsa-form').setTitle('Nuevo Socio');
    SpreadsheetApp.getUi().showSidebar(html);
}

function loadAmounts() {
    var numOfEmployees = getNumberOfEmployees();
    var ss = getEmployeeSheet();
    var dataRangeDef = "D2:D" + numOfEmployees;
    var data = ss.getRange(dataRangeDef).getValues();

    return data;
}

function modifyAmount(dlgTitle, dlgSubtitle, operation) {
    SpreadsheetApp.setActiveSheet(getEmployeeSheet());
    var ui = SpreadsheetApp.getUi();
    var promptResult = null;
    var percStr = "";
    var percentage = 0;

    do {
        promptResult = ui.prompt(dlgTitle, dlgSubtitle, ui.ButtonSet.OK_CANCEL);
      
        percStr = promptResult.getResponseText();
        if (percStr && percStr.length > 0 && !isNaN(percStr)) {
            percentage = Number(percStr);
        } else {
            percentage = 0;
        }
      
    } while (promptResult.getSelectedButton() == ui.Button.OK && Math.abs(percentage) > 100);
  
    if (percentage == 0)
        return;
 
    var amounts = loadAmounts();
    for (var i = 0; i < amounts.length; i++) {
        if (operation == "MUL") {
            amounts[i][0] = amounts[i][0] * (1 + percentage / 100);
        } else {
            amounts[i][0] = amounts[i][0] / (1 + percentage / 100);
        }
    }
  
    var ss = getEmployeeSheet();
    var dataRangeDef = "D2:D" + (amounts.length + 1);
    ss.getRange(dataRangeDef).setValues(amounts);
}

function incDecPercentage() {
    modifyAmount("Variación de Importe", "Porcentaje de Variación [-100 a 100]:", "MUL");
}

function reset() {
    modifyAmount("Variación de Importe", "Reestablecer en % [-100 a 100]:", "DIV");
}


