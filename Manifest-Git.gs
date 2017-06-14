// Global Variables
var master_Spreadsheet; //Pointer to Master Manifest Copy -> retrieved by retrieve_master() 
// Static HTML used to build sidebard display and full table view
var stylesheet = "<style>p{margin: 0;font-size: 15px;}h3{margin-bottom:0;margin-top:5px;margin-left:0;}.center{text-align:center;}.deletion{color: red;}" +
".insertion{color: green;}.remove{color:red;}.add{color:green;}.modify{font-weight:bold;}table,th,td{border: 1px solid black;}</style>";
var prepend = "<!DOCTYPE html><html><head>"+ stylesheet + "</head><body>";
var append = "</body></html>";

function onOpen() { 
  // Create Menu Item 'Manifest Git' and sub entries
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Manifest Control')
      .addItem('Branch (copy)', 'branch')
      .addSubMenu(ui.createMenu('Diff')
          .addItem('Diff Sheet With Master (Simple)', 'diff_current')
          .addItem('Diff All With Master (Simple)', 'diff_all')
          .addItem('Diff Sheet With Other Version (Simple)', 'diff_with_other')
          .addItem('Diff Sheet With Full Table View (Robust)', 'diff_current_with_daff'))
      .addSubMenu(ui.createMenu('Merge')
          .addItem('Merge with other', 'merge')
          .addItem('Overwrite to Master', 'overwrite_to_master'))
      .addSeparator()
      .addSubMenu(ui.createMenu("Validate Against Ontology")
          .addItem("Validate Sheet", "validate")
          .addItem("Validate All", "validate_all"))
      .addToUi();
  retrieve_master();
  // TODO - ADD automatic Validate upon opening 
};
// Include file -> used in Html file to include stylesheet
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
// Retrieve current master version - may have to alter later - possibly retrieve by Unique Google Sheet ID 
function retrieve_master() {
  try { 
    master_Spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/11j392Y2P2tJ8LUgTGVklejheYmTYTQlfPG8aXekVcbc/edit#gid=995555814");
  } catch(err) {
    Browser.msgBox("Could not retrieve Master Manifest: " + err);
  }
}
// Check to see if current Branch is already Master Manifest - True if so, false otherwise
function check_branch(current) {
  retrieve_master();
  if (current.getId().equals(master_Spreadsheet.getId())) {
      Browser.msgBox("You already appear to be on the master copy");
      return true;
  }
  return false; 
}
// Basic Branch - Copy
function branch() {
  // Import Master Sheet
  retrieve_master();
  
  // Get Active SpreadSheet
  var current = SpreadsheetApp.getActiveSpreadsheet();
  
  // TO DO - Prompt User for TITLE to then prepend to 'Manifest' and date 
  // Copy of Active SpreadSheet #TODO -> Optimize, taking too long 
  var new_ss = current.copy("Manifest - " + Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm:ss"));
  var url = new_ss.getUrl();
  
  // Prompt with link to new copy of Manifest. 
  SpreadsheetApp.setActiveSpreadsheet(new_ss);
  showurl(url);
}


/*   ############### MERGE ################# */ 
// Merge two children to Master
function merge() {
  // Import Master Sheet
  retrieve_master();
  
  // Import Other Spreadsheet
  var other_Spreadsheet = retrieve_other();
  if (other_Spreadsheet == null) {return; };
  
  // Get Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Retrieve Other Sheet
  // Check to see if already on Master
  if (check_branch(current)) {return;}
  
  // TO DO - VISUAL OF MERGE CONFLICTS WITH USER INPUT
  
  
  // Prompt with link to new copy of Manifest. 
  SpreadsheetApp.setActiveSpreadsheet(new_ss);
  showurl(url);
}
// Overwrite current Spreadsheet to Master Copy
function overwrite_to_master() {
  // Import Master Sheet
  retrieve_master();
  
  // Get Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Check to see if already on Master
  if (check_branch(current)) {return;}
  
  // Prompt User for URL to other SpreadSheet
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Script will overwite Master Manifest with current Spreadsheet - Contine?', ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response != ui.Button.YES){return;}
  
  // Clear Every Sheet in Master Sheet
  var master_sheets = master_Spreadsheet.getSheets();
  var last_sheet = master_Spreadsheet.insertSheet("Last Sheet - To Be Deleted");
  for (var i = 0; i < master_sheets.length; i++) {
    master_Spreadsheet.deleteSheet(master_sheets[i]);
  }
  
  // Copy every sheet in Current to Master Spreadsheet
  var current_sheets = current.getSheets();
  for (var i = 0; i < current_sheets.length; i++) {
    current_sheets[i].copyTo(master_Spreadsheet);
  }
  master_Spreadsheet.deleteSheet(last_sheet);
  // Change name from "copy of" to Regulare
  var master_sheets = master_Spreadsheet.getSheets();
  for (var i = 0; i < master_sheets.length; i++) {
    master_sheets[i].setName(current_sheets[i].getName());
  }
}


/*   ############### VALIDATE ################# */ 
function validate() {
  // First Idea - Retrieve 'Ontology Property' Column and insert in list
  var current = SpreadsheetApp.getActive();
  var current_sheet = current.getActiveSheet();
  var current_sheet_values = current_sheet.getDataRange().getValues();
  var ontology_properties = [];
  if (current_sheet_values[2][1].toString() == "Ontology Property") {
    for (var i = 3; i < current_sheet.getDataRange().getHeight(); i++) {
       var val = current_sheet_values[i][1].toString();
       if (val != "") {
         ontology_properties.push(val);
       }
    }
    var property_paths = parse_property(ontology_properties);
  }
}
function parse_property(properties) {
  var parsed_properties = []; // create an empty array
  
  for(var i=0; i < properties.length; i++){
    var current = properties[i];
    var split_string = current.split(":");
    Browser.msgBox(split_string);
    parsed_properties.push(split_string);
  }
  return parsed_properties;
}


/*   ############### DIFF ALL WITH MASTER ################# */ 
function diff_all() {  
  // Import Master Sheet
  retrieve_master(); 
   
  // Get Active Spreadsheet - current
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


/*   ############### DIFF SHEET WITH MASTER ################# */ 
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
  // Browser.msgBox(html_string);
  var html = HtmlService.createHtmlOutput(html_string).setTitle('Diff Display');
  SpreadsheetApp.getUi().showSidebar(html);
}


/*   ############### DIFF WITH OTHER  ################# */ 
function diff_with_other() {
  // Import Master Sheet
  retrieve_master(); 
  
  // Import Other Spreadsheet
  var other_Spreadsheet = retrieve_other();
  if (other_Spreadsheet == null) {return; };
   
  // Gett Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Get current sheet on Current spreadsheet, then get get corresponding sheet for 
  var c1 = current.getActiveSheet();
  var sheet_name = c1.getName();
  try {
    var o1 = other_Spreadsheet.getSheetByName(sheet_name);
  } catch (err) {
    Browser.msg("It appears this sheet does not exist or has a different name on Master Copy");
  }
  
  var html_string = prepend + diff_sheet(o1,c1) + append;
  var html = HtmlService.createHtmlOutput(html_string).setTitle('Diff Display');
  SpreadsheetApp.getUi().showSidebar(html);
}

/*   ############### DIFF SHEET ################# */ 
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
  // Iterate through values of Copy
    for (var i=0; i < values_b.length; i++) {
      // Check to see if within range of Rows of Master Copy = ROW INSERTION
      if (i >= values_a.length) {
        var row_string = "<p>Row Insert at: " + i + "<p class=\"insertion\">";
        for (var j=0; j < values_b[i].length; j++){
           row_string += values_b[i][j].toString() + " | ";
           insertions++
        }
        html = html + row_string + "</p>";
        continue;
      }
      for (var j=0; j < values_b[i].length; j++) {
        // Check to see if within range of cols of Master Copy = COL INSERTION
        if (j >= values_a[i].length) {
          html += insertion(i,j,values_b[i][j].toString());
          insertions++;
        } else {
          var a_value = values_a[i][j].toString();
          var b_value = values_b[i][j].toString();
          if (!(a_value.equals(b_value))) {
            if (!found_diff) {
              found_diff = true;
            } 
            // Current no Longer has Value -> Deletion
            if (b_value.equals("")) {
              html += deletion(i,j,a_value);
              deletions++; 
            // Current has Value where Master has nothing -> Insertion
            } else if (a_value.equals("")) {
              html += insertion(i,j,b_value);
              insertions++;
            // Current and Master both have value but Differ -> Modification 
            }else {
              html += modification(i,j,a_value,b_value);
              modifications++; 
            }
          }
        }
      }
      //Check if Copy missed range of col of Master Copy - COL DELETION
      if (j < values_a[i].length) {
          for (var j2 = j; j2 < values_a[i].length; j2++) {
            html += deletion(i,j,values_a[i][j2].toString());
            deletions++;
          }
      }
    }
    //Check if Copy missed range of rows of Master Copy - ROW DELETION
    if (i < values_b.length){
    }
    
    // If there are no Differences -> mark as such 
    if (found_diff) {
      return title + "<p>Modifcations: " + modifications + " Insertions: " + insertions + " Deletions: " + deletions + "</p>" + html;
    } else {
      return title += "<p style=\"display:inline-block;margin:0;\"> No Difference </p>"
    }
  } 
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }  
}


/*   ############### HELPER FUNCTIONS  ################# */ 
// Helper Function for Deletion - Coordinates and Value
function deletion(i,j,a_value) {
  var html = "<p>Deletion on coord: (" + (i+1).toString() + ", " + toCol(j) + ")</p>";
  html += "<p class=\"deletion center\">" + a_value + "</p>";
  return html;
}
// Helper Function for Insertion - Coordinates and Value
function insertion(i,j,b_value) {
  var html = "<p>Insertion on coord: (" + (i+1).toString() + ", " + toCol(j) + ")</p>";
  html += "<p class=\"insertion center\">" + b_value + "</p>";
  return html;
}
// Helper Function for Modifcation - Coordinates and two Values
function modification(i,j,a_value,b_value) {
  var html = "<p>Modified on coord: (" + (i+1).toString() + ", " + toCol(j) + ")</p>";
  html += "<p class=\"deletion center\">" + a_value + "</p><p class=\"center\">      <===============>      </p><p class=\"insertion center\"> " + b_value + "</p>";
  return html;
}
// Helper functions for Table Coordinates
var Alpha = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
function toCol(num){
  return Alpha[num];
}
function retrieve_other() { 
  // Prompt User for URL to other SpreadSheet
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Diff with other version of Manifest', 'URL to Version:', ui.ButtonSet.OK);

  // Process the user's response.
  var other_url = response.getResponseText()
  
  // Import Other Spreadsheet
  var other_Spreadsheet;
  try { 
    other_Spreadsheet = SpreadsheetApp.openByUrl(other_url);
  } catch(err) {
    Browser.msgBox("Could not retrieve Other Manifest: " + err);
    return null;
  }
  return other_Spreadsheet;
}
// Show Clickable link to User
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

// Convert to CSV @found online 
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
// Robust Diff using Daff Library -> create a HTML Table to display  
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
  var html = HtmlService.createHtmlOutput(html_string).setTitle('Diff Display').setWidth(1000).setHeight(600);
  // Browser.msgBox(html_string);
  SpreadsheetApp.getUi().showModelessDialog(html,"Diff with Table  -  Ignore Null and Empty Cells - Unchanged");
}