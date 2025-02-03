/**
 * Loads the parameter JSON strings from the "JF_SCRIPT_PARAMS" sheet and returns a dictionary.
 *
 * Each row in column A should contain a valid JSON string.
 * For example:
 *   {"Input Sheet":{"Name":"Test Opp"}}
 *   {"Table Header Row":{"Name":1}}
 *   {"oli":{"object_label":"Opportunity","object_api_name":"Opportunity__c"}}
 *   {"oli":{"object_label":"Product","object_api_name":"Product__c"}}
 *   {"un":{"object_label":"Customer ID","object_api_name":"Customer_ID__c"}}
 *   {"un":{"object_label":"Customer Name","object_api_name":"Customer_Name__c"}}
 *
 * The function works as follows:
 *   - It first reads all rows from column A.
 *   - It checks whether the first row is a header row by attempting to parse its cell.
 *     If it fails, it assumes that row 1 is a header and starts processing from row 2.
 *   - For each parsed JSON object, it iterates over its top-level keys.
 *     If a key already exists in the resulting dictionary:
 *       - If its value is not already an array, convert it to an array.
 *       - Then push the new value.
 *     Otherwise, store it directly.
 *
 * @returns {Object} A dictionary of parameters.
 */
function loadParametersFromJson(ss) {
  logMessage(ss,"Starting loadParametersFromJson()", ss);
  
  var sheet = ss.getSheetByName("JF_SCRIPT_PARAMS");
  if (!sheet) {
    logMessage(ss,"Sheet 'JF_SCRIPT_PARAMS' not found.");
    throw new Error("Sheet 'JF_SCRIPT_PARAMS' not found.");
  }
  
  var lastRow = sheet.getLastRow();
  logMessage(ss,"Detected last row in JF_SCRIPT_PARAMS: " + lastRow);
  if (lastRow < 1) {
    logMessage(ss,"No data found in 'JF_SCRIPT_PARAMS' sheet.");
    throw new Error("No data found in 'JF_SCRIPT_PARAMS' sheet.");
  }
  
  // Read all rows from column A.
  var data = sheet.getRange(1, 1, lastRow, 1).getValues();
  logMessage(ss,"Raw JSON data: " + JSON.stringify(data));
  
  // Determine whether the first row is a header.
  var startRow = 0;
  try {
    var testCell = String(data[0][0]).trim();
    JSON.parse(testCell);
    // If parsing succeeds, no header row is present.
  } catch (e) {
    // logMessage(ss,"First row did not parse as JSON. Assuming it is a header row and skipping it.");
    startRow = 1;
  }
  
  var params = {};
  
  // Process each row from startRow onward.
  for (var i = startRow; i < data.length; i++) {
    var cellValue = String(data[i][0]).trim();
    if (!cellValue) {
      logMessage(ss,"Skipping blank cell at row " + (i + 1));
      continue;
    }
    
    try {
      var parsed = JSON.parse(cellValue);
      // logMessage(ss,"Parsed JSON at row " + (i + 1) + ": " + JSON.stringify(parsed));
      
      // For each top-level key in the parsed object:
      for (var key in parsed) {
        var value = parsed[key];
        // If the key already exists, aggregate it.
        if (params.hasOwnProperty(key)) {
          if (!Array.isArray(params[key])) {
            // Convert to array if it isn't already.
            params[key] = [ params[key] ];
          }
          params[key].push(value);
          // logMessage(ss,"Appended parameter for key '" + key + "': " + JSON.stringify(value));
        } else {
          // Otherwise, set it directly.
          params[key] = value;
          // logMessage(ss,"Added parameter: '" + key + "' -> " + JSON.stringify(value));
        }
      }
    } catch (e) {
      logMessage(ss,"Error parsing JSON at row " + (i + 1) + ": " + e.message);
      throw new Error("Error parsing JSON at row " + (i + 1) + ": " + e.message);
    }
  }
  
  logMessage(ss,"Finished loadParametersFromJson(). Final parameters: " + JSON.stringify(params));
  flushLogs(ss);
  return params;
}

/**
 * Example usage of loadParametersFromJson().
 */
function exampleLoadParameters(ss) {
  var params = loadParametersFromJson(ss);
  logMessage(ss,"Loaded Parameters:\n" + JSON.stringify(params, null, 2));
  return params;
}
