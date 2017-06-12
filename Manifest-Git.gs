// Global Variables
var master_Spreadsheet; 
var stylesheet = "<style>p{margin: 0;font-size: 15px;}h3{margin:10px;}.center{text-align:center;}.deletion{color: red;}.insertion{color: green;}</style>"
var prepend = "<!DOCTYPE html><html><head>"+ stylesheet + "</head><body>";
var append = "</body></html>";

function onOpen() { 
  // Create Menu Item 'Manifest Git' and sub entries
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Manifest Git')
      .addItem('Branch', 'branch')
      .addSubMenu(ui.createMenu('Diff')
          .addItem('Diff Sheet With Master', 'diff_current')
          .addItem('Diff All With Master', 'diff_all')
          .addItem('Diff Sheet With Other Version', 'diff_with_other'))
      .addItem("Validate Against Ontology", "validate")
      .addItem("Merge", "merge")
      .addToUi();
};
// Include file -> used in Html file to include stylesheet
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function retrieve_master() {
  // Currently retrieve master version - may have to alter later - possibly retrieve by Unique Google Sheet ID 
  try { 
    master_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/11j392Y2P2tJ8LUgTGVklejheYmTYTQlfPG8aXekVcbc/edit#gid=995555814");
  } catch(err) {
    Browser.msgBox("Could not retrieve Master Manifest: " + err);
  }
}

function check_branch(current) {
  // Check to see if current Branch is already Master Manifest - True if so, false otherwise
  if (current.getId().equals(master_Spreadsheet.getId())) {
      Browser.msgBox("You already appear to be on the master copy");
      return true;
  }
  return false; 
}

function branch() {
  // Get Active SpreadSheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if branch, prompt to continue? 
  if (check_branch(current)) {
    Browser.msgBox("You already appear to be on a Branch of the master. Close to continue.");
  }
  
  // Copy of Active SpreadSheet #TODO -> Optimize, taking too long 
  var new_ss = ss.copy(Session.getActiveUserLocale() + new Date().getTime());
  var url = new_ss.getUrl();
  
  // Prompt with link to new copy of Manifest. 
  SpreadsheetApp.setActiveSpreadsheet(new_ss);
  showurl(url);
}
function diff_all() {  
  // Import Master Sheet
  retrieve_master(); 
   
  // Gett Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Check to see if already on Master
  if (check_branch(current)) {return;}
  
  // Get Sheets
  var current_sheets = current.getSheets();
  var master_sheets = master_Spreadsheet.getSheets();  
  
  var built_string = "";
  for (var i=0; i < current_sheets.length; i++) {
      var current_sheet = current_sheets[i]
      var curr_name = current_sheet.getName();
      var found = false;
      for (var j=0; j < master_sheets.length; j++) {
        var master_sheet = master_sheets[j]
        var master_name = master_sheet.getName();
        if (curr_name.equals(master_name)){
          found = true;
          built_string += diff_sheet(master_sheet, current_sheet);
        }
      }
      if (!found) {
        built_string += "<h3>Sheet Not Found: " + curr_name + "</h3><p class=\"center deletion\">Sheet does not exist in Master Version, or is differently named</p>"
      }
  }
  
  var html_string = prepend + built_string + append;
  // Browser.msgBox(html_string);
  var html = HtmlService.createHtmlOutput(html_string).setTitle('Diff Display');
  SpreadsheetApp.getUi().showSidebar(html);
}

function diff_current() {
  // Import Master Sheet
  retrieve_master(); 
   
  // Gett Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Check to see if already on Master
  if (check_branch(current)) {return;}
  
  // Get current sheet on Current spreadsheet, then get get corresponding sheet for 
  var c1 = current.getActiveSheet();
  var sheet_name = c1.getName();
  try {
    var m1 = master_Spreadsheet.getSheetByName(sheet_name);
  } catch (err) {
    Browser.msg("It appears this sheet does not exist or has a different name on Master Copy");
  }
  
  var html_string = prepend + diff_sheet(m1,c1) + append;
  Browser.msgBox(html_string);
  var html = HtmlService.createHtmlOutput(html_string).setTitle('Diff Display');
  SpreadsheetApp.getUi().showSidebar(html);
}

function diff_with_other() { 
  var ui = DocumentApp.getUi();
  var response = ui.prompt('Getting to know you', 'May I know your name?', ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.YES) {
    Logger.log('The user\'s name is %s.', response.getResponseText());
  } else if (response.getSelectedButton() == ui.Button.NO) {
    Logger.log('The user didn\'t want to provide a name.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
  
  // Import Master Sheet
  retrieve_master(); 
   
  // Gett Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Check to see if already on Master
  if (check_branch(current)) {return;}
  
  // Get current sheet on Current spreadsheet, then get get corresponding sheet for 
  var c1 = current.getActiveSheet();
  var sheet_name = c1.getName();
  try {
    var m1 = master_Spreadsheet.getSheetByName(sheet_name);
  } catch (err) {
    Browser.msg("It appears this sheet does not exist or has a different name on Master Copy");
  }
  
  var html_string = prepend + diff_sheet(m1,c1) + append;
  Browser.msgBox(html_string);
  var html = HtmlService.createHtmlOutput(html_string).setTitle('Diff Display');
  SpreadsheetApp.getUi().showSidebar(html);
}

function diff_sheet(sheet_a, sheet_b) {
  //Master Sheet always as a, current as b
  var insertions = 0;
  var deletions = 0;
  var modifications = 0;
  var title = "<h3><normal> On Sheet: </normal><bold>" + sheet_a.getName() + "</bold></h3>";
  var data_a = sheet_a.getDataRange();
  var data_b = sheet_b.getDataRange();
  var values_a = data_a.getValues();
  var values_b = data_b.getValues();
  var found_diff = false; 
  var html = "";
  
  try {
    for (var i=0; i < values_a.length; i++) {
      for (var j=0; j < values_a[i].length; j++) {
        var a_value = values_a[i][j].toString();
        var b_value = values_b[i][j].toString();
        
        if (!(a_value.equals(b_value))) {
          if (!found_diff) {
            found_diff = true;
          }
          
          // Current no Longer has Value -> Deletion
          if (b_value.equals("")) {
            html += "<p>Deletion on coord: (" + (i+1).toString() + ", " + toCol(j) + ")</p>";
            html += "<p class=\"deletion center\">" + a_value + "</p>";
            deletions++; 
            
          // Current has Value where Master has nothing -> Insertion
          } else if (a_value.equals("")) {
            html += "<p>Insertion on coord: (" + (i+1).toString() + ", " + toCol(j) + ")</p>";
            html += "<p class=\"insertion center\">" + b_value + "</p>";
            insertions++;
            
          // Current and Master both have value but Differ -> Modification 
          }else {
            html += "<p>Modified on coord: (" + (i+1).toString() + ", " + toCol(j) + ")</p>";
            html += "<p class=\"deletion center\">" + a_value + "</p><p class=\"center\">      <===============>      </p><p class=\"insertion center\"> " + b_value + "</p>";
            modifications++; 
          }
        }
      }
    }
    // If there are no Differences -> mark as such 
    if (found_diff) {
      return title + "<p style=\"margin:0;\">Modifcations: " + modifications + " Insertions: " + insertions + " Deletions: " + deletions + "</p>" + html;
    } else {
      return title += "<p style=\"display:inline-block;margin:0;\"> No Difference </p>"
    }
  } 
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }  
}

function validate() {
}

// Helper functions
var Alpha = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
function toCol(num){
  return Alpha[num];
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
    var csvFile = convertRangeToCsvFile_(sheet);
    // create a file in the Docs List with the given name and the csv data
    folder.createFile(fileName, csvFile);
  }
}

function convertRangeToCsvFile_(sheet) {
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
function diff_current_with_daff() {
  // Import Master Sheet
  retrieve_master(); 
   
  // Gett Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Check to see if already on Master
  if (check_branch(current)) {return;}
  
  // Get current sheet on Current spreadsheet, then get get corresponding sheet for 
  var c1 = current.getActiveSheet();
  var sheet_name = c1.getName();
  try {
    var m1 = master_Spreadsheet.getSheetByName(sheet_name);
  } catch (err) {
    Browser.msg("It appears this sheet does not exist or has a different name on Master Copy");
  }
  // Convert booth sheets to CSV to then pass to Daff 
  var c1_csv = convertRangeToCsvFile_(c1);
  var m1_csv = convertRangeToCsvFile_(m1);
  
  // Call Daff Function 
  var html_string = prepend + daff_sheets(c1_csv, m1_csv) + append; 
  var html = HtmlService.createHtmlOutput(html_string).setTitle('Diff Display');
  Browser.msgBox(html_string);
  SpreadsheetApp.getUi().showSidebar(html);
}