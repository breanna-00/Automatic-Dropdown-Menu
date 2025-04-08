/* Note: To apply colors and "chip" display styles to dropdown menus, follow these steps:
1. Add all additional resources to the necessary rows.
2. Click the pencil icon on one of the dropdown menus to edit its appearance.
3. Make the desired changes, then click "Done" and "Apply to all instances."

**Heads up:** Each time you do this, a new data validation rule titled "Value contains one from range" will be created. */

function test(){
  // get current spreadsheet
  var currentSS = SpreadsheetApp.getActiveSpreadsheet();

  // get range of drop down menus
  var range = currentSS.getRange("D3:M70");
  console.log(range.getValues());
  /*for (let i = 3; i < 67; i++){
    for (let n = 0; n < 10; n++){
      // get the value of the last cell that has a dropdown menu
      console.log(currentSS.getRange(range));
    }
  }*/
}

function addDropDown() {
  // get current spreadsheet
  var currentSS = SpreadsheetApp.getActiveSpreadsheet();

  // get range of drop down menus
  var range = currentSS.getRange("D3:M70");

  const columns = ["D","E","F","G","H","I","J","K","L","M"];

  // goes through sheet, one cell at a time (Plan to make more efficent)
  for (let i = 3; i < 67; i++){
    
    for (let n = 0; n < 9; n++){
      // get the value of the last cell that has a dropdown menu
      var col = columns[n]
      var newDropDownRange = col + i;
      var findLastDropDownPos = currentSS.getRange(newDropDownRange);
      var lastDropDownValues = findLastDropDownPos.getValues();

      // checks if the user has chosen an option from a dropdown menu
      if(lastDropDownValues != ''){
        // if dropdown menu has option chosen, then add new dropdown menu to cell adjacent to the right
        nextColumn = columns[n+1]; 
        neighborRange = nextColumn + i;
        
        createDropDown(neighborRange);
      }
    }
  }
}

function createDropDown(range) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sheet2");

  // finds location for the dropdown
  const cell = sheet.getRange(range);

  // adds the options to dropdown and builds dropdown
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(["Director","Director/ASC","Director/ Dr. Pagni","C-Real Narrative","Previous C-REAL Survey","CoBro Service Summaries","AUHSD Data","GEAR UP Team Narrative","Outside Data (Clearinghouse, CSAC, etc)","AUHSD Senior Survey"]).build();
  
  cell.setDataValidation(rule);
}

// Add menu that has button to add new menu (Plan to change this to onEdit())
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Add Dropdown Menu')
      .addItem('Add Dropdown Menu', 'addDropDown')
      .addToUi();
}

function onEdit(e) {
  // gets location of current active cell
  var col = e.range.columnStart;
  var row = e.range.rowStart;

  // gets location next to current active cell
  const columns = ["D","E","F","G","H","I","J","K","L","M"];
  var newCol = columns[col-3] + "" + row;

  // checks that cell is in correct range and that option was chosen from dropdown menu before adding another one next to it
  if(col >= 4 && row >= 3 && e.value.length > 0) {
    createDropDown(newCol);
  }
}
