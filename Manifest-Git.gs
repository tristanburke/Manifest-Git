/* 
   -*- coding: utf-8 -*-
   Copyright (c) 2017, Syapse Inc. All rights reserved.
   Created: 6/14/17
*/
   
// Global Variables
var master_Spreadsheet; //Pointer to Master Manifest Copy -> retrieved by retrieve_master() 
// Static HTML used to build sidebard display and full table view
var stylesheet = "<style>\
                      p{margin: 0;font-size: 15px;}\
                      h3{margin-bottom:0;margin-top:5px;margin-left:0;}\
                      .center{text-align:center;}\
                      .deletion{color: red;}\
                      .insertion{color: green;}\
                      .remove{color:red;}\
                      .add{color:green;}\
                      .modify{font-weight:bold;}\
                      table,th,td{border: 1px solid black;}\
                  </style>";
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
          .addItem('Merge (Simple)', 'merge_with_other')
          .addItem('Write Sheet to Master (Simple)', 'write_sheet')
          .addItem('Write All to Master (Simple)', 'write_all')
          .addItem('Override Master (Robust)', 'override_master'))
      .addSubMenu(ui.createMenu('Pull')
          .addItem('Pull (Sheet)', 'pull_sheet')
          .addItem('Pull - Force Overwrite (Sheet)', 'pull_override'))
      .addSeparator()
      .addSubMenu(ui.createMenu("Validate Against Ontology")
          .addItem("Validate Sheet", "validate")
          .addItem("Validate All", "validate_all"))
      .addToUi();
  retrieve_master();
  // TODO - ADD automatic Validate upon opening 
};


/*   ############### Branch ################# */ 
// Basic Branch - Copy
function branch() {
  // Import Master Sheet
  retrieve_master();
  
  // Get Active SpreadSheet
  var current = SpreadsheetApp.getActiveSpreadsheet();
  
  // Only Branch from Master
   if (!(current.getId().equals(master_Spreadsheet.getId()))) {
     Browser.msgBox("Due to the simplicity of this versioning system, branches are only allowed from the Master Manifest");
     return;
  }
  // Check for Branches Page -> if not there, create
  /*var master_sheets = master_Spreadsheet.getSheets();
  var branches_sheet = master_sheets[master_sheets.length - 1];
  if (branches_sheet.getName() != "Branches" ) {
     branches_sheet = master_Spreadsheet.insertSheet("Branches");
  } */
  
  // Prompt User for TITLE to then prepend to 'Manifest' and date 
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter Title for new Manifest Copy', 'Title:', ui.ButtonSet.OK);
  var title = response.getResponseText()
  var date = Utilities.formatDate(new Date(), "PST", "yyyy-MM-dd HH:mm:ss");
  
  // Copy of Active SpreadSheet #TODO -> Optimize, taking too long 
  var new_ss = current.copy(title + " (Manifest Copy) - " + date);
  var url = new_ss.getUrl();
  
  // Add new Branch to Branches Sheet on Master
  // branches_sheet.appendRow([Title, Date, url]);
  
  // Prompt with link to new copy of Manifest. 
  SpreadsheetApp.setActiveSpreadsheet(new_ss);
  showurl(url);
}


/*   ############### MERGE ################# */ 
// Merge two children to Master
function merge_with_other() {
  // Import Master Sheet
  retrieve_master();
  
  // Get Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Check to see if already on Master
  if (check_branch(current)) {return;}
  
  // Import Other Spreadsheet
  var other_Spreadsheet = retrieve_other();
  if (other_Spreadsheet == null) {return; };
  
  // Retrieve Corresponding Sheet for Other and Master 
  var c1 = current.getActiveSheet();
  var sheet_name = c1.getName();
  try {
    var m1 = master_Spreadsheet.getSheetByName(sheet_name);
    var o1 = other_Spreadsheet.getSheetByName(sheet_name);
  } catch (err) {
    Browser.msg("It appears this sheet does not exist or has a different name on Master Copy");
    return;
  }
  
  /** TO DO - VISUAL OF MERGE CONFLICTS WITH USER INPUT
   PLAN keep track of Diff's through Coordinates -> Diff both current and other
   with master sheet -> build array of Coord for each. Compare arrays, and present 
   merge conflict if both diff at same Coord -> present basic UI with values and a button to choose 
   one over other (left title vs. right title) -> then build single array with coord from both 
   individual arrays, and chosen coord in both, and write diff values to master 
   **/
  var merge_current_master = diff_sheet(m1, c1, true)[1];
  var merge_other_master = diff_sheet(m1, o1, true)[1];
  
  var diff_conflict = [];
  var diff_master = [];
  
  // Iterate through Current, check for same coordinates 
  for (var i = 0; i < merge_current_master.length; i++) {
    var conflict = false;
    var curr_coord_diff = merge_current_master[i]
    for (var j = 0; j < merge_other_master.length; j++) {
      if (merge_other_master[j] != null) {
        var curr_other_diff = merge_other_master[j]
        if (curr_coord_diff[0] == curr_other_diff[0] && curr_coord_diff[1] == curr_other_diff[1]) {
          // Mark Conflict for current -> Don't add Diff to diff_master
          conflict = true;
          // Store Conlfict in diff_conflict, Format === [i, j, [curr_value, other_value]]
          diff_conflict.push([curr_coord_diff[0], curr_coord_diff[1], [curr_coord_diff[2], curr_other_diff[2]]]);
          // Mark Conflict for other -> Don't add Diff to diff_master, no longer need double reference, marked for quicker looping
          merge_other_master[j] = null;
        }
      }
    }
    if (!conflict) {
      diff_master.push(curr_coord_diff);
    }
  }
  for (var i = 0; i < merge_other_master.length; i++) {
    if (merge_other_master[i] != null) {
      diff_master.push(merge_other_master[i]);
    }
  }
  Browser.msgBox("There are " + diff_conflict.length + " Merge Conflicts.\n Press \"OK\" to continue and pick values");
  var ui = SpreadsheetApp.getUi();
  for (var i = 0; i < diff_conflict.length; i++) {
    var curr_diff = diff_conflict[i];
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert("Press YES to pick " + current.getName() + " value.\n Press NO to pick " + other_Spreadsheet.getName()
    + " value.\n At Coordinate : ("
    +  (curr_diff[0]+1).toString() + ", " + toCol(curr_diff[1]) + ")\n\n" + 
    "YES: " + current.getName() + ":" + curr_diff[2][0] + "\n\n" + 
    "NO: " + other_Spreadsheet.getName() + ":" + curr_diff[2][1] + "\n\n" , 
    ui.ButtonSet.YES_NO_CANCEL);
  
    // Process the user's response.
    if (response == ui.Button.YES){
     diff_master.push([curr_diff[0], curr_diff[1], curr_diff[2][0]]);
    } else if (response == ui.Button.NO) {
     diff_master.push([curr_diff[0], curr_diff[1], curr_diff[2][1]]);
    }  else {
      Browser.msgBox("Merge Aborted. No changes written.");
      return;
    }
  }
  write_diffs(diff_master, m1);
  return;
     
  // Prompt with link to new copy of Manifest. 
  // SpreadsheetApp.setActiveSpreadsheet(new_ss);
  // showurl(url);
}
// Write all Sheets to Master
function write_all() {
  // Import Master Sheet
  retrieve_master();
  
  // Get Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Check to see if already on Master
  if (check_branch(current)) {return;}
  
  // Get Sheets
  var current_sheets = current.getSheets();
  var master_sheets = master_Spreadsheet.getSheets();
  
  var pages_written = ""
  for (var i=0; i < current_sheets.length; i++) {
      var current_sheet = current_sheets[i]
      var curr_name = current_sheet.getName();
      try {
        var master_sheet = master_Spreadsheet.getSheetByName(curr_name);
        var merge_current_master = diff_sheet(master_sheet, current_sheet, true)[1];
        if (merge_current_master.length > 0) {
          write_diffs(merge_current_master, master_sheet);
          pages_written += curr_name + '\n, ';
        }
      } catch (err) {
        Browser.msgBox("Could not find sheet: " + curr_name);
      }
  }
  if (pages_written == "") {
    Browser.msgBox("Done. No Changes - nothing written");
  } else {
    Browser.msgBox("Done. Written to Master. Pages overwritten:\n" + pages_written); 
  }
}
// Write Current Sheet to Master
function write_sheet() {
  // Import Master Sheet
  retrieve_master();
  
  // Get Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Check to see if already on Master
  if (check_branch(current)) {return;}
  
  // Get Sheet and check master also has 
  var c1 = current.getActiveSheet();
  var sheet_name = c1.getName();
  try {
    var m1 = master_Spreadsheet.getSheetByName(sheet_name);
  } catch (err) {
    Browser.msg("It appears this sheet does not exist or has a different name on Master Copy");
    return;
  }
  var merge_current_master = diff_sheet(m1, c1, true)[1];
  write_diffs(merge_current_master, m1);
  Browser.msgBox("Done. Sheet written to Master");  
}

// Write an array of diffs on a sheet to same sheet on Master
function write_diffs(diffs, sheet) {

  for (var i=0; i < diffs.length; i++) {
    var current_cell = sheet.getRange(diffs[i][0]+1, diffs[i][1]+1);
    var current_value = diffs[i][2];
    current_cell.setValue(current_value);
  }
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

/*   ###############   Pull   ################# */ 
function pull_sheet() {
  // Import Master Sheet
  retrieve_master(); 
   
  // Get Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Check to see if already on Master
  if (check_branch(current)) {return;}
  
  // Get current sheet on Current spreadsheet, then get get corresponding sheet from Master 
  var c1 = current.getActiveSheet();
  var sheet_name = c1.getName();
  try {
    var m1 = master_Spreadsheet.getSheetByName(sheet_name);
  } catch (err) {
    Browser.msg("It appears this sheet does not exist or has a different name on Master Copy");
    return;
  }
  
  var diff_conflict = diff_sheet(m1, c1, true)[1];
  var m1_data = m1.getDataRange();
  var m1_values = m1_data.getValues();
  
  Browser.msgBox("There are " + diff_conflict.length + " Differences Between Current and Master.\n Press \"OK\" to continue and pick values");
  var ui = SpreadsheetApp.getUi();
  var diff_master = [];
  for (var index = 0; index < diff_conflict.length; index++) {
    var ui = SpreadsheetApp.getUi();
    var diff = diff_conflict[index];
    var i = diff[0]
    var j = diff[1]
    var curr_val = diff[2]
    var master_val = diff[3]
    var response = ui.alert("Press YES to pick Master Manifest value.\n Press NO to pick " + current.getName()
    + " value.\n At Coordinate : ("
    +  (i+1).toString() + ", " + toCol(j) + ")\n\n" + 
    "YES: Master Copy:  " + master_val + "\n\n" +
    "NO: "  + current.getName() + ":  " + curr_val + "\n\n",
    ui.ButtonSet.YES_NO_CANCEL);
  
    // Process the user's response.
    if (response == ui.Button.NO){
     diff_master.push([i, j, curr_val]);
    } else if (response == ui.Button.YES) {
     diff_master.push([i, j, master_val]);
    }  else {
      Browser.msgBox("Pull Aborted. No changes written.");
      return;
    }
  }
  write_diffs(diff_master, c1);
  Browser.msgBox("Done. Master Sheet pulled to " + current.getName()); 
}
// Override and automatically pull all diffs onto current sheet
function pull_override() {
  // Import Master Sheet
  retrieve_master(); 
   
  // Get Active Spreadsheet - current
  var current = SpreadsheetApp.getActive();
  
  // Check to see if already on Master
  if (check_branch(current)) {return;}
  
  // Get current sheet on Current spreadsheet, then get get corresponding sheet from Master 
  var c1 = current.getActiveSheet();
  var sheet_name = c1.getName();
  try {
    var m1 = master_Spreadsheet.getSheetByName(sheet_name);
  } catch (err) {
    Browser.msg("It appears this sheet does not exist or has a different name on Master Copy");
    return;
  }
  
  var diff_master = diff_sheet(c1, m1, true)[1];
  write_diffs(diff_master, c1);
  Browser.msgBox("Done. Master Sheet pulled to " + current.getName()); 
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
      try {
        var master_sheet = master_Spreadsheet.getSheetByName(curr_name);
        var temp = diff_sheet(master_sheet, current_sheet);
        built_string += temp[0];
      } catch (err) {
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
  
  // Get current sheet on Current spreadsheet, then get get corresponding sheet from Master 
  var c1 = current.getActiveSheet();
  var sheet_name = c1.getName();
  try {
    var m1 = master_Spreadsheet.getSheetByName(sheet_name);
  } catch (err) {
    Browser.msg("It appears this sheet does not exist or has a different name on Master Copy");
    return;
  }
  
  var temp = diff_sheet(m1,c1);
  var html_string = prepend + temp[0] + append;
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
    return;
  }
  
  var temp_String, merge = diff_sheet(c1, o1);
  var html_string = prepend + temp_String + append;
  var html = HtmlService.createHtmlOutput(html_string).setTitle('Diff Display');
  SpreadsheetApp.getUi().showSidebar(html);
}

/*   ############### DIFF SHEET ################# */ 
function diff_sheet(sheet_a, sheet_b, _merge) {
  // implement _merge as optional parameter -> default false, true when using diff for merge
  if (typeof _merge === 'undefined') { _merge = false; }
  
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
  var coord_and_diff = []
  
  var end_row = 0;
  var end_col = 0;
  try {
  // Iterate through values of Copy
    for (var i=0; i < values_b.length; i++) {
      for (var j=0; j < values_b[i].length; j++) {
        // Check to see if within range of Rows of Master Copy 
        if (i >= values_a.length || j >= values_a[i].length) {
             if (values_b[i][j] != "") { 
               if (!found_diff) { found_diff = true;}
               if (_merge) {
                  var temp = [i, j, values_b[i][j].toString(), ""];
                  coord_and_diff.push(temp)
               }
               html += insertion(i,j,values_b[i][j].toString());
               insertions++;
            }
        } else {
          var a_value = values_a[i][j].toString();
          var b_value = values_b[i][j].toString();
          if (!(a_value.equals(b_value))) {
            // Boolean for "No Difference" summary 
            if (!found_diff) { found_diff = true;}
            
            // Information for Merging
            if (_merge) {
              var temp = [i, j, b_value, a_value];
              coord_and_diff.push(temp)
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
      end_row = i;
      end_col = j;
    }
    //Check if Copy missed range of rows of Master Copy - ROW DELETION
    if (end_row < values_a.length || end_col < values_a[0].length){
      for (var i = 0; i < values_a.length; i++) {
        for (var j = 0; j < values_a[i].length; j++) {
          if (i > end_row || j > end_col){
            if (values_a[i][j] != "") {
              if (!found_diff) { found_diff = true;}
              if (_merge) {
                var temp = [i, j, ""];
                coord_and_diff.push(temp)
              }
              html += deletion(i,j,"",values_a[i][j].toString());
              deletions++;
            }
          }
        }
      }
    }
    
    // If there are no Differences -> mark as such 
    if (found_diff) {
      return [title + "<p>Modifcations: " + modifications + " Insertions: " + insertions + " Deletions: " + deletions + "</p>" + html, coord_and_diff];
    } else {
      return [title += "<p style=\"display:inline-block;margin:0;\"> No Difference </p>", coord_and_diff];
    }
  } 
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
    return;
  }  
}


/*   ############### HELPER FUNCTIONS  ################# */
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