/**
 * Builds an in-memory data table from the input sheet using parameter settings.
 *
 * It uses:
 *   - The "Input Sheet" parameter to determine the input sheet name.
 *   - The "Table Header Row" parameter to get the header row number.
 *   - The "jeellyfish_line_item" group in the parameters to map header columns.
 *
 * For each row in the input sheet (starting at header row+1, until a blank row is encountered),
 * the function builds an object where:
 *   - For groupings with a single mapping, the object gets an entry with key = object_api_name.
 *   - For groupings with multiple mappings, the values are nested under the group key.
 *
 * @returns {Array<Object>} An array of records representing the data table.
 */
function buildDataTableFromParamsDynamic(ss) {
  logMessage(ss,"Starting buildDataTableFromParamsDynamic with " + ss + ".");
  
  // Load parameters.
  var params = loadParametersFromJson(ss);
  logMessage(ss,"Loaded parameters: " + JSON.stringify(params));
  
  Logger.log("Your Parames Are: " + params );

  var inputSheetName = params["Input Sheet"].Name;
  var headerRowNumber = params["Table Header Row"].Name;
  
  var sheet = ss.getSheetByName(inputSheetName);
  if (!sheet) {
    throw new Error("Input sheet '" + inputSheetName + "' not found.");
  }
  logMessage(ss,"Processing input sheet: " + inputSheetName + ", header row: " + headerRowNumber);
  
  // Read the header row.
  var lastColumn = sheet.getLastColumn();
  var headerRow = sheet.getRange(headerRowNumber, 1, 1, lastColumn).getValues()[0];
  
  // Build mapping for each grouping (keys other than "Input Sheet" and "Table Header Row").
  var overallMapping = {}; // key: grouping key, value: array of mapping objects { api: <object_api_name>, col: <columnIndex> }
  for (var key in params) {
    if (key === "Input Sheet" || key === "Table Header Row") {
      continue;
    }
    // For each grouping, normalize value to an array.
    var groupData = params[key];
    var groupArray = Array.isArray(groupData) ? groupData : [groupData];
    overallMapping[key] = [];
    
    for (var i = 0; i < groupArray.length; i++) {
      var item = groupArray[i];
      if (!item.object_label || !item.object_api_name) {
        continue;
      }
      var labelToFind = String(item.object_label).trim();
      var apiName = item.object_api_name;
      var found = false;
      for (var j = 0; j < headerRow.length; j++) {
        if (String(headerRow[j]).trim() === labelToFind) {
          overallMapping[key].push({ api: apiName, col: j });
          found = true;
          break;
        }
      }
      if (!found) {
        logMessage(ss,"Warning: Header label '" + labelToFind + "' for group '" + key + "' not found.");
      }
    }
  }
  
  logMessage(ss,"Final overall mapping: " + JSON.stringify(overallMapping));
  
  // Read data rows from headerRowNumber + 1 until a blank row is encountered.
  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange(headerRowNumber + 1, 1, lastRow - headerRowNumber, lastColumn);
  var dataRows = dataRange.getValues();
  
  var records = [];
  for (var r = 0; r < dataRows.length; r++) {
    var row = dataRows[r];
    // Stop if the first cell is blank (assumes that indicates end-of-data).
    if (row[0] === "" || row[0] === null) {
      logMessage(ss,"Blank row encountered at row " + (headerRowNumber + 1 + r) + ". Stopping data processing.");
      break;
    }
    
    var record = {};
    // For each group in the overall mapping, assign values.
    for (var group in overallMapping) {
      var mappings = overallMapping[group];
      if (mappings.length === 0) continue;
      
      // If only one mapping, add directly under its API field.
      if (mappings.length === 1) {
        record[mappings[0].api] = row[mappings[0].col];
      } else {
        // If multiple mappings, nest the values under the group key.
        record[group] = {};
        for (var k = 0; k < mappings.length; k++) {
          record[group][mappings[k].api] = row[mappings[k].col];
        }
      }
    }
    records.push(record);
    // logMessage(ss,"Processed row " + (headerRowNumber + 1 + r) + ": " + JSON.stringify(record));
  }
  
  logMessage(ss,"Built data table with " + records.length + " records.");
  logMessage(ss,"Final Data Table:\n" + JSON.stringify(records, null, 2));
  flushLogs(ss);
  return records;
}