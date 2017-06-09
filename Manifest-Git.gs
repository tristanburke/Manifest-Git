function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{name: "export as csv files", functionName: "saveAsCSV"}];
  var gitMenuEntries = [{name: "Branch", functionName: "branch"},{name: "Diff Current", functionName: "diff_current"}, {name: "Diff All", functionName: "diff_all"}];
  ss.addMenu("CSV", csvMenuEntries);
  ss.addMenu("Manifest-Git", gitMenuEntries);
};

function branch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
 
  // Copy of Master Manifest
  var new_ss = ss.copy(Session.getActiveUserLocale() + new Date().getTime());
  var url = new_ss.getUrl();
  
  SpreadsheetApp.setActiveSpreadsheet(new_ss);
  showurl(url);
}
function diff_all() {
  // Import Master Sheet
  var master = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/11j392Y2P2tJ8LUgTGVklejheYmTYTQlfPG8aXekVcbc/edit#gid=995555814");
  
  // Retrieve Sheets of both Master and Current Sheets
  var current = SpreadsheetApp.getActive();
  
  var current_sheets = current.getSheets();
  var master_sheets = master.getSheets();  
  
  var built_string = "";
  for (var i=0; i < current_sheets.length; i++) {
      var current_sheet = current_sheets[i]
      var curr_name = current_sheet.getName();
      for (var j=0; j < current_sheets.length; j++) {
        var master_sheet = master_sheets[j]
        var master_name = master_sheet.getName();
        if (curr_name.equals(master_name)){
          built_string += diff_sheet(current_sheet, master_sheet);
        }
      }
  }
  
  var html_string = append(prepend() + built_string);
  Browser.msgBox(html_string);
  var html = HtmlService.createHtmlOutput(html_string).setTitle('Diff Display');
  SpreadsheetApp.getUi().showSidebar(html);
}

function diff_current() {
  
  // Import Daff.js 
  // var url = "http://tristan-burke.com/js/daff.js";
  // var javascript = UrlFetchApp.fetch(url).getContentText();
  
  // Import Master Sheet
  var master = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/11j392Y2P2tJ8LUgTGVklejheYmTYTQlfPG8aXekVcbc/edit#gid=995555814");
  
  // Retrieve Sheets of both Master and Current Sheets
  var current = SpreadsheetApp.getActive();
  
  var c1 = current.getActiveSheet();
  var sheet_name = c1.getName();
  var m1 = master.getSheetByName(sheet_name);
  
  var html_string = append(prepend() + diff_sheet(m1,c1));
  // Browser.msgBox(html_string);
  var html = HtmlService.createHtmlOutput(html_string).setTitle('Diff Display');
  SpreadsheetApp.getUi().showSidebar(html);
}

function diff_sheet(sheet_a, sheet_b) {
  var html = "<h3> On Sheet: <bold>" + sheet_a.getName() + "</bold></h3>";
  var data_a = sheet_a.getDataRange();
  var data_b = sheet_b.getDataRange();
  var values_a = data_a.getValues();
  var values_b = data_b.getValues();
  var found_diff = false; 
  
  try {
    for (var i=0; i < values_a.length; i++) {
      for (var j=0; j < values_a[i].length; j++) {
        var a_value = values_a[i][j].toString();
        var b_value = values_b[i][j].toString();
        
        if (!(a_value.equals(b_value))) {
          if (!found_diff) {
            found_diff = true;
          }
          html += "<p>Difference on coord: (" + (i+1).toString() + ", " + toCol(j) + ") <br></p>";
          html += "<p style=\"color:green;text-align:center;\">" + a_value +
            "<br><\p><p style=\"text-align:center;\">      <===============>      <\p><p style=\"color:red;text-align:center;\"><br> " + b_value + "</p>";
        }
      }
    }
    if (found_diff) {
      return html;
    } else {
      return html += "<p style=\"display:inline-block;\"> No Difference </p>"
    }
  } 
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }  
}

var Alpha = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
function toCol(num){
  return Alpha[num];
}

function prepend() {
  var html = "<!DOCTYPE html><html><head><style type=\"text/css\">.p{line-height:110%;margin-top:0px;margin-bottom:0px;bottom:0;top:0;}.h3{margin:0;line-height:110%;.body{padding:0;margin:0;}</style><base target=\"_top\"></head><body>"
  return html;
}

function append(html) {
   html += "</body></html>";
   return html;
}

function showurl(url) {
  var app = UiApp.createApplication().setHeight('60').setWidth('200');
  app.setTitle("Branch");
  var panel = app.createPopupPanel()
  var link = app.createAnchor('Click Here for Copy', url);
  panel.add(link);
  app.add(panel);
  var doc = SpreadsheetApp.getActive();
  doc.show(app);
}

function saveAsCSV(Spreadsheet, name) {
  var ss = Spreadsheet;
  var sheets = ss.getSheets();
  // create a folder from the name of the spreadsheet
  var folder = DriveApp.createFolder(name);
  for (var i = 0 ; i < sheets.length ; i++) {
    var sheet = sheets[i];
    // append ".csv" extension to the sheet name
    fileName = sheet.getName() + ".csv";
    // convert all available sheet data to csv format
    var csvFile = convertRangeToCsvFile_(fileName, sheet);
    // create a file in the Docs List with the given name and the csv data
    folder.createFile(fileName, csvFile);
  }
}

function convertRangeToCsvFile_(csvFileName, sheet) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}