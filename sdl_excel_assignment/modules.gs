/**
 * Checks if the NRICs in Column F of the table in the worksheet "Nominal Roll"
 * is masked properly.
 * 
 * @return {boolean}  `True` if the NRICs are masked properly, `False` if otherwise.
*/
function Check_Mask_NRIC() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Nominal Roll");

  var original_nric_range = sheet.getRange("E2:E95");
  var masked_nric_range = sheet.getRange("F2:F95");

  var original_values = original_nric_range.getValues();
  var masked_nric_values = masked_nric_range.getValues();
  var formulas = masked_nric_range.getFormulas();

  for (var i = 0; i < masked_nric_values.length; i++) {
    var row_num = i + 2;
    var ans = original_values[i][0].substring(0, 1) + "XXXX" + original_values[i][0].substring(original_values[i][0].length-4 ,original_values[i][0].length);

    if (masked_nric_values[i][0] != ans || formulas[i][0] == "") {
      SpreadsheetApp.getUi().alert("Cell F" + row_num + " has to contain a formula, and mask the NRIC correctly.");
      return false;
    }

  }
  SpreadsheetApp.getUi().alert("All NRICs are masked correctly! Good job :)");
  return true;
}

/**
 * Helper function for `Check_Item_Referenced_Correctly` which simulates `INDEX-MATCH` in Excel.
 * 
 * Can consider using binary search instead of linear search since the 4Ds are in sorted order.
 * O(log n) time is more efficient than O(n) time, but again, small dataset.
 * 
 * @param {string} identifier - the exact value that serves as the key.
 * @param {integer} identifier_col - the column number where the key resides in.
 * @param {integer} target_col - the column number of the table where the desired value is in.
 * 
 * @return {string} the value that we want to get
*/
function Index_Match(data, identifier, identifier_col, target_col) {
  for (var i  = 0; i < data.length; i++) {
    if (data[i][identifier_col] == identifier) {
      return data[i][target_col];
    }
  }
}

/**
 * Checks if the specified item is correctly referenced from the table in the worksheet "Nominal Roll"
 * to the table in the worksheet "Leave List" by generating the answers, then checking it against
 * provided values in the tables.
 * 
 * @param {string} item - the item to be referenced.
 * 
 * @param {string} cell_range - the range of cells in the table in "Leave List" to refer to.
 * 
 * @return {boolean} `True` if the items are referenced properly, `False` if otherwise.
*/
function Check_Item_Referenced_Correctly(item, cell_range){
  var leave_list_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leave List");

  var leave_list_range = leave_list_sheet.getRange(cell_range);
  var leave_list_values = leave_list_range.getValues();
  var formulas = leave_list_range.getFormulas();

  var leave_list_4Ds = leave_list_sheet.getRange("A2:A75").getValues();
  var nom_roll_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Nominal Roll");
  var nom_roll_data = nom_roll_sheet.getRange(2, 1, nom_roll_sheet.getLastRow() - 1, nom_roll_sheet.getLastColumn()).getValues();

  for (var i = 0; i < leave_list_4Ds.length; i++) {
    var row_num = i + 2;

    // generate answers
    var ans = item == "name" ? Index_Match(nom_roll_data, leave_list_4Ds[i][0], 0, 2) : Index_Match(nom_roll_data, leave_list_4Ds[i][0], 0, 5);

    if (ans != leave_list_values[i][0] || formulas[i][0] == "") {
      SpreadsheetApp.getUi().alert("Cell " + cell_range.charAt(0) + row_num + " has to contain a formula, and reference the " + item + " correctly.");
      return false;
    }
  }

  SpreadsheetApp.getUi().alert("All " + item + "s are referenced correctly! Good job :)");
  return true;
}

/**
 * Checks if the cells "D2:D75" in the table in the worksheet "Leave List" all contain 
 * conditional formatting.
 * 
 * @return {boolean} `True` if there is conditional formatting for all the cells, `False` if otherwise.
*/
function Check_Leave_Type_Validation() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leave List");
  var range = sheet.getRange("D2:D75");
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    var row_num = i + 2;
    var validation_formula = "'Leave Types'!A2:A6";
    
    try {
      var t = range.getCell(i+1, 1).getDataValidation().getCriteriaType();

      if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
        var data_validation = range.getCell(i+1, 1).getDataValidation()
        var criteria_values = data_validation.getCriteriaValues()[0];
        var criteria_sheet_name = criteria_values.getSheet().getName();
        var criteria_range_A1 = criteria_values.getA1Notation();

        if ("\'" + criteria_sheet_name + "\'!" + criteria_range_A1 !== validation_formula) {
          SpreadsheetApp.getUi().alert("Cell D" + row_num + " requires data validation from the entire table in Worksheet 'Leave Types'.");
          return false;
        }
      } else {
        SpreadsheetApp.getUi().alert("Cell D" + row_num + " requires data validation from the entire table in Worksheet 'Leave Types'.");
        return false;
      }
    } catch (error) {
      SpreadsheetApp.getUi().alert("Cell D" + row_num + " requires data validation from the entire table in Worksheet 'Leave Types'.");
      return false;
    }
  }
  
  SpreadsheetApp.getUi().alert("All cells in column D have the correct data validation. Good job!");
  return true;
}

/**
 * Checks if each leave length is correctly evaluated in the table in the worksheet "Leave List".
 * 
 * @return {boolean} `True` if all leaves lengths are calculated correctly, `False` if otherwise.
*/
function Check_Leave_Length() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leave List");
  var range = sheet.getRange("A2:G75");
  var values = range.getValues();
  var formulas = sheet.getRange("G2:G75").getFormulas()
  
  for (var i = 0; i < values.length; i++) {
    var row_num = i + 2;

    // evaluate answer
    var leave_type = range.getCell(i+1, 4).getValue();
    var start_date = range.getCell(i+1, 5).getValue();
    var end_date = range.getCell(i+1, 6).getValue();
    var ans = (leave_type === 'MA') ? 0.5 : Math.floor((end_date - start_date) / (24 * 60 * 60 * 1000)) + 1;
    
    if (values[i][6] != ans || formulas[i][0] == "") {
      SpreadsheetApp.getUi().alert("Cell G" + row_num + " must have the correct formula and evaluate to the correct number of days.");
      return false;
    }
  }
  
  SpreadsheetApp.getUi().alert("All leave durations are correctly calculated!");
  return true;
}

/**
 * Calcluates the total days of leave taken by a serviceman identified with his 4D.
 * 
 * @param {string} identifier - the serviceman's 4D.
 * 
 * @return {(integer|float)} The total number of days the serviceman took leave based on the table in "Leave List".
*/
function Calculate_Leave_Total(identifier) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leave List");
  var range = sheet.getRange("A2:G75");
  var values = range.getValues();
  var total = 0

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == identifier) {
      total += values[i][6]
    }
  }

  return total
}

/**
 * Checks if the total number of days of leave for each serviceman is correctly evaluated in the table
 * in "Leave Totals By Personnel".
 * 
 * @return {boolean} `True` if total number of days of leave is correctly evaluated, `False` if otherwise.
*/
function Check_Leave_Total_Table() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leave Totals By Personnel");
  var range = sheet.getRange("A2:C95");
  var values = range.getValues();
  var formulas = sheet.getRange("C2:C95").getFormulas();
  
  for (var i = 0; i < values.length; i++) {
    var row_num = i + 2;
    var ans = Calculate_Leave_Total(values[i][0])
    
    if (values[i][2] != ans || formulas[i][0] == "") {
      SpreadsheetApp.getUi().alert("Cell C" + row_num + " must have the correct formula and evaluate to the correct number of days.");
      return false;
    }
  }
  
  SpreadsheetApp.getUi().alert("All leave totals are correctly calculated!");
  return true;
}

/**
 * Checks if conditional formatting is applied correctly on the table in "Leave Totals By Personnel".
 * 
 * App Script does not support checking current conditional formatting rules through the `ConditionalFormatRule`
 * object, so we have to check the individual values and their respective backgrounds. Funny thing is,
 * only the cells which are conditionally formatted have their backgrounds changed.
 * 
 * @return {boolean} `True` if conditional formatting is applied correctly, `False` if otherwise.
*/
function Check_Conditional_Formatting() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leave Totals By Personnel");
  var range = sheet.getRange("C2:C95");
  var values = range.getValues()

  for (var i = 0; i < values.length; i++) {
    var background = range.getCell(i+1, 1).getBackground()

    if (values[i][0] >= 5 & ['#ffffff', '#d8d8d8', '#d9e2f3'].includes(background) || values[i][0] < 5 & !(['#ffffff', '#d8d8d8', '#d9e2f3'].includes(background))) {
      SpreadsheetApp.getUi().alert("You need to have conditional formatting!");
      return false;
    }
  }

  SpreadsheetApp.getUi().alert("Conditional formatting is correct!");
  return true;
}

/**
 * Checks if the total strength is evaluated correctly through a formula.
 * 
 * @return {boolean} `True` if the total strength is correctly evaluated, `False` if otherwise.
*/
function Check_Total_Strength() {
  var cell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Present Strength").getRange("F1");
  
  if (cell.getValue() != 94 || !cell.getFormula()) {
    SpreadsheetApp.getUi().alert("Total strength needs to be evaluated correctly using a formula.");
    return false;
  }
  
  SpreadsheetApp.getUi().alert("Total Strength is correct!");
  return true;
}

/**
 * Checks if the total number of personnel on leave per day is calculated correctly.
 * 
 * Note: Javascript (which is what App Script is based on) reads "<M>/<D>/<YYYY>" as a DateObject
 *       directly from the sheet. Calculations are done with that in mind.
 * 
 * @return {boolean} `True` if the leave strength per day is correctly evaluated, `False` if otherwise.
*/
function Check_Leave_Personnel() {
  var present_str_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Present Strength");
  var leave_personnel_values = present_str_sheet.getRange("A2:B38").getValues();
  var formulas = present_str_sheet.getRange("B2:B38").getFormulas();

  var leave_list_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leave List");
  var leave_occur_dates = leave_list_sheet.getRange("E2:F75").getValues();

  var dict = {};

  // populate values in dictionary
  for (var i = 0; i < leave_personnel_values.length; i++) {
    dict[leave_personnel_values[i][0]] = 0;
  }

  // calculate leave personnel per day and store in dictionary
  for (var i = 0; i < leave_occur_dates.length; i++){
    var start_date = leave_occur_dates[i][0];
    var end_date = leave_occur_dates[i][1];

    dict[start_date] += 1;

    // `start_date` and `end_date` are interpreted as DateObjects by Javascript
    // so we increase the dates by one through this 
    while (end_date - start_date > 0) {
      start_date.setDate(start_date.getDate() + 1);
      dict[start_date] += 1;
    }
  }

  // check answers
  for (var i = 0; i < leave_personnel_values.length; i++) {
    var row_num = i + 2
    var date = leave_personnel_values[i][0];

    if (dict[date] !== leave_personnel_values[i][1] || formulas[i][0] == "") {
      SpreadsheetApp.getUi().alert("Cell B" + row_num + " must have the correct formula and evaluate to the correct strength.");
      return false;
    }
  }

  SpreadsheetApp.getUi().alert("All number of personnel on leave for each day are correctly calculated!");
  return true;
}

/**
 * Checks if the total number of personnel present per day is calculated correctly.
 * 
 * @return {boolean} `True` if the present strength per day is correctly evaluated, `False` if otherwise.
*/
function Check_Present_Strength() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Present Strength");
  var values = sheet.getRange("B2:C38").getValues();
  var formulas = sheet.getRange("C2:C38").getFormulas();
  const total_str = 94;

  for (var i = 0; i < values.length; i++) {
    var row_num = i + 2;

    if (values[i][1] !== total_str - values[i][0] || formulas[i][0] == "") {
      SpreadsheetApp.getUi().alert("Cell C" + row_num + " must have the correct formula and evaluate to the correct strength.");
      return false;
    }
  }

  SpreadsheetApp.getUi().alert("All number of personnel present for each day are correctly calculated!");
  return true;
}

/**
 * Checks if the average strength is evaluated correctly.
 * 
 * @return {boolean} `True` if the average strength is correctly evaluated, `False` if otherwise.
*/
function Check_Average_Strength() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Present Strength");
  var daily_str = sheet.getRange("C2:C38").getValues();
  var total = 0;

  var cell = sheet.getRange("F2");

  for (var i = 0; i < daily_str.length; i++) {
    total += daily_str[i][0];
  }
  
  if (cell.getValue() !== Math.ceil(total / 37) || !cell.getFormula()) {
    SpreadsheetApp.getUi().alert("Average strength needs to be evaluated correctly using a formula.");
    return false;
  }
  
  SpreadsheetApp.getUi().alert("Average Strength is correct!");
  return true;
}

/**
 * Checks if there is a bar chart, and said chart has all required components.
 * 
 * @return {boolean} `True` if the bar chart is correctly done, `False` if otherwise.
*/
function Check_Bar_Chart() {
  var charts = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Present Strength").getCharts();
  
  if (charts.length == 0) {
    SpreadsheetApp.getUi().alert("Where is your bar chart...?");
    return false;
  } 

  var chart = charts[0];
  var options = chart.getOptions();

  var ranges = chart.getRanges().map(r => `'${r.getSheet().getSheetName()}'!${r.getA1Notation()}`);

  // can't use <element> in ranges, because array syntax works differently i guess
  if (!ranges.includes("'Present Strength'!C2:C38") || !ranges.includes("'Present Strength'!A2:A38")) {
    SpreadsheetApp.getUi().alert("Are you sure you got your data for the bar chart from the correct place? You need to get it from A2:A38 and C2:C38 in the 'Present Strength' worksheet..");
    return false;
  }

  var title = options.get('title');

  if (title == null || title == "") {
    SpreadsheetApp.getUi().alert("We need a title for the bar chart. The title is 'Daily Strength'.");
    return false;
  } else if (title.trim() != "Daily Strength") {
    SpreadsheetApp.getUi().alert("The bar chart title is to be 'Daily Strength'.");
    return false;
  }

  var x_title = options.get('hAxis.title')

  if (x_title == null || x_title == "") {
    SpreadsheetApp.getUi().alert("We need a x-axis title for the bar chart. The title is 'Date'.");
    return false;
  } else if (x_title.trim() != "Date") {
    SpreadsheetApp.getUi().alert("The bar chart x-axis title is to be 'Date'.");
    return false;
  }

  var y_title = options.get('vAxes.0.title')

  if (y_title == null || y_title == "") {
    SpreadsheetApp.getUi().alert("We need a y-axis title for the bar chart. The title is 'Present Strength'.");
    return false;
  } else if (y_title.trim() != "Present Strength") {
    SpreadsheetApp.getUi().alert("The bar chart y-axis title is to be 'Present Strength'.");
    return false;
  }

  SpreadsheetApp.getUi().alert("Nicely done! The bar chart is correct :)");
  return true;
}

/**
 * Clears the values of a specified range.
 * 
 * @param {string} range_start - the start of the range of cells to clear values.
 * @param {string} range_start - the end of the range of cells to clear values.
*/
function clearCells(sheet, range_start, range_end) {
  sheet.getRange(range_start + ":" + range_end).setValue("")
}
