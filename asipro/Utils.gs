function lpad(s, c, l) {
    var padLen = l - s.length;

    if (padLen <= 0) 
        return s;

    var r = "";
    for (var i = 0; i < padLen; i++) {
        r += c;
    }
  
    return r + s;
}


function rpad(s, c, l) {
    var padLen = l - s.length;

    if (padLen <= 0) 
        return s;

    var r = "";
    for (var i = 0; i < padLen; i++) {
        r += c;
    }
  
    return s + r;
}

function filler(c, l) {
    return Array(l+1).join(c);
}

function getSheetByName(sheetName) {
    var spr = SpreadsheetApp.getActiveSpreadsheet();
    var ss = spr.getSheetByName(sheetName);
    return ss;
}

function lastBusinessDayOfMonth(year, month) {
    var date = new Date();
    var offset = 0;
    var result = null;
    
    if ('undefined' === typeof year || null === year) {
        year = date.getFullYear();
    }
    
    if ('undefined' === typeof month || null === month) {
        month = (date.getMonth() + 1) % 12;
    } else {
        month = (month + 1) % 12;
    }
  
    if (month === 0) {
        year++;
    }

    do {
        result = new Date(year, month, offset);
        
        offset--;
    } while (0 === result.getDay() || 6 === result.getDay());

    return result;
}

function stripNonNumeric(s) {
    return s.replace(/\D/g,'');
}

function test() {
    var t = rpad(lpad(stripNonNumeric("D.N.I. 21.402.979"), "0", 8), " ", 22);
    var u = rpad(lpad(stripNonNumeric("C.I.      6.121.936"), "0", 8), " ", 22);
    var d1 = lastBusinessDayOfMonth();
    var d2 = lastBusinessDayOfMonth(2018, 7); // august
    var d3 = lastBusinessDayOfMonth(2018,11); // december
    var t = rpad(COELSA_ASIPRO_COMPANY_DESC, "@", 20);
    var s = Array(10).join("  ");
    var l = s.length;
    var blanks = filler(" ", 5);
    var asterisks = filler("*", 5);
}


