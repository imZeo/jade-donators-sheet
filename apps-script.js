function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = sheet.getRange("Welcome and FAQ!B3");

  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, "UTC", "yyyy-MM-dd HH:mm")
  range.setValue(formattedDate.toLocaleString());
}

// function onEdit(e) {
//   var sheet = e.source.getActiveSheet();
//   var range = e.range;
//   var col = range.getColumn();
  
//   if (col === 1) { // Column A
//     var searchString = range.getValue();
//     if (searchString) {
//       var searchValue = Browser.inputBox("Add Value", "Enter a value to add to Column B:", Browser.Buttons.OK_CANCEL);
//       if (searchValue !== "cancel") {
//         sheet.getRange(range.getRow(), 2).setValue(searchValue);
//       }
//     }
//   }
// }

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Menu")
    .addItem("Log a Donation", "showInputDialog")
    .addToUi();
}

function showInputDialog() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All Donators");
  
  if (sheet) {
    var searchValueResponse = ui.prompt("Log a Donation", "Enter the name of the guildie (IGN name, e.g., Zeo.1026):", ui.ButtonSet.OK_CANCEL);
    
    if (searchValueResponse.getSelectedButton() === ui.Button.OK) {
      var searchValue = searchValueResponse.getResponseText();
      var data = sheet.getDataRange().getValues();
      var newRow = -1;
      var existingRow = -1;
      var existingResult = "";
      var newResult = "";
      var guildMember = "";
      
      for (var row = 0; row < data.length; row++) {
        if (data[row][0].toUpperCase() === searchValue.toUpperCase()) { // Case-insensitive comparison
          existingRow = row + 1;
          existingResult = sheet.getRange(existingRow, 3).getValue(); // Get existing result in Column C
          guildMember = data[row][0];
          break;
        } else if (data[row][0] === "") {
          newRow = row + 1;
          break;
        }
      }
      
      if (existingRow !== -1) {
        var inputValue = ui.prompt("Log a Donation", "Enter the donation:", ui.ButtonSet.OK_CANCEL);
        if (inputValue.getSelectedButton() === ui.Button.OK) {
          var newValue = Number(inputValue.getResponseText());
          var currentValue = sheet.getRange(existingRow, 2).getValue();
          var oldTotal = currentValue / 10000 + " gold";
          sheet.getRange(existingRow, 2).setValue(currentValue + newValue); // Increment Column B
          
          // Calculate the new result in Column C after updating Column B
          newResult = sheet.getRange(existingRow, 3).getValue();
          var newTotal = (currentValue + newValue) / 10000 + " gold";
          
          var cUpdated = "";
          if (newResult !== existingResult) {
            cUpdated = "\n===========\nCurrent rank: " + existingResult + "\nNew rank: " + newResult;
          }
          
          var summary = "Summary:\n===========\nGuild Member: " + guildMember + "\n===========\nOld total: " + oldTotal + "\nNew total: " + newTotal + cUpdated;
          ui.alert("Donation logged with success. Mo' money mo' problems!\n\n" + summary);
          return;
        } else {
          ui.alert("Operation canceled.");
          return;
        }
      }
      
      if (newRow === -1) {
        newRow = data.length + 1;
      }
      
      sheet.getRange(newRow, 1).setValue(searchValue);
      var inputValue = ui.prompt("Log a Donation", "Enter the donation:", ui.ButtonSet.OK_CANCEL);
      if (inputValue.getSelectedButton() === ui.Button.OK) {
        var newValue = inputValue.getResponseText();
        sheet.getRange(newRow, 2).setValue(Number(newValue)); // Column B
        
        // Add the formula to Column C for the new row
        sheet.getRange(newRow, 3).setFormula('=IF(B' + newRow + ' < 1000000, "Non Donator", IF(B' + newRow + ' < 4000000, "Generous Skritt", IF(B' + newRow + ' < 10000000, "The Pimp", "The High Roller")))');
        
        newResult = sheet.getRange(newRow, 3).getValue(); // Get the output of the new formula
        var newTotal = Number(newValue) / 10000 + " gold";
        
        var summary = "Summary:\n===========\nGuild Member: " + searchValue + "\nNew total: " + newTotal + "\nNew rank: " + newResult;
        ui.alert("Huzzah, the donation and new donator has been logged o/\n\n" + summary);
      } else {
        ui.alert("Operation canceled.");
      }
    }
  } else {
    ui.alert("Sheet not found. Make sure you have a sheet named 'All Donators'.");
  }
}
