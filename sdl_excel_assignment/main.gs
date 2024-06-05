/**
 * The main function of the program. This function is triggered by the `Grade Work` button in the "Instructions" worksheet,
 * which calls up the other functions in `modules.gs`.
 */
function Main() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Instructions");
  var hidden_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hidden Tracker");
  var status = "Completed";
  var hidden_status = "Done";
  var row_num = 8; // first row to fill in starts from 8
  var hidden_row_num = 2;

  SpreadsheetApp.getUi().alert("Assessment grading started.");
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check if your NRICs in the 'Nominal Roll' worksheet is masked correctly.");

    if (Check_Mask_NRIC()) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
    return;
    }
  }
  
  row_num++;
  hidden_row_num++;
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check if the names in the 'Leave List' worksheet are referenced correctly.");

    if (Check_Item_Referenced_Correctly("name", "B2:B75")) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
      return;
    }
  }

  row_num++;
  hidden_row_num++;
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check if the masked NRICs in the 'Leave List' worksheet are referenced correctly.");

    if (Check_Item_Referenced_Correctly("masked NRIC", "C2:C75")) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
    return;
    }
  }
  
  row_num++;
  hidden_row_num++;
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check if you have data validation for the types of leaves applied by the servicemen. This step will take longer, so don't worry if the script seems like it is not running, because it is running!");

    if (Check_Leave_Type_Validation()) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
      return;
    }
  }
  
  row_num++;
  hidden_row_num++;
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check if the leave duration for each leave occurrence is correct.");

    if (Check_Leave_Length()) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
      return;
    }
  }
  
  row_num++;
  hidden_row_num++;
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check if the total amount of leave applied by each serviceman is correct. This can be found in the worksheet 'Leave Total by Personnel'.");

    if (Check_Leave_Total_Table()) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
      return;
    }
  }
  
  row_num++;
  hidden_row_num++;
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check if the OOC highlighting in the worksheet 'Leave Total By Personnel' is done correctly.");

    if (Check_Conditional_Formatting()) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
      return;
    }
  }
  
  row_num++;
  hidden_row_num++;
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check if the total strength is correctly evaluated.");
    
    if (Check_Total_Strength()) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
      return;
    }
  }
  
  row_num++;
  hidden_row_num++;
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check if the number of personnel on leave each day is correct.");

    if (Check_Leave_Personnel()) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
      return;
    }
  }
  
  row_num++;
  hidden_row_num++;
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check if the present strength for each day is correct.");

    if (Check_Present_Strength()) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
      return;
    }
  }
  
  row_num++;
  hidden_row_num++;
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check if the average strength is correctly evaluated.");

    if (Check_Average_Strength()) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
      return;
    }
  }
  
  row_num++;
  hidden_row_num++;
  
  if (sheet.getRange("L" + row_num).getValue() == status) {
    if (hidden_sheet.getRange("B" + hidden_row_num).getValue() == "") {
      SpreadsheetApp.getUi().alert("You think you can fool me?");
      clearCells(sheet, "L" + row_num, "L" + row_num + 11);
      return;
    }
  } else {
    SpreadsheetApp.getUi().alert("I will now check the bar chart is appropriately presented.");

    if (Check_Bar_Chart()) {
      sheet.getRange("L" + row_num).setValue(status);
      hidden_sheet.getRange("B" + hidden_row_num).setValue(hidden_status);
    } else {
      return;
    }
  }
  
  row_num++;
  hidden_row_num++;

  SpreadsheetApp.getUi().alert("Good job completing the assessment! Take a screenshot of this very screen now, and submit your work (ie. this screenshot) to your trainers!");
}
