  /**
   * Converts a single-row range into a grouped JSON string.
   *
   * The range should be formatted as:
   *   GroupName, Key1, Value1, Key2, Value2, ... 
   * Processing stops when a blank key is encountered.
   *
   * Example:
   *   If the row contains:
   *     ["MyGroup", "Name", "John", "Age", "30", "", "ignored"]
   *   The function returns:
   *     {"MyGroup": {"Name": "John", "Age": "30"}}
   *
   * @param {range} inputRange A single-row range containing the group name and key/value pairs.
   * @return {String} A JSON string representing the grouped key/value pairs.
   * @customfunction
   */
  function JSONIFY_groupname_key_value_pairings(inputRange) {
    // Ensure the input range is valid and use the first (and only) row.
    if (!inputRange || inputRange.length < 1) {
      return "Error: Input range must contain at least one row.";
    }
    var row = inputRange[0];
    if (row.length < 1) {
      return "Error: Input row must contain at least one cell for the group name.";
    }
    
    // The first cell is the overall group name.
    var groupName = row[0];
    if (groupName === null || groupName === undefined || groupName.toString().trim() === "") {
      return "Error: Group name (first cell) is blank.";
    }
    
    // Process the remaining cells as key/value pairs.
    var obj = {};
    // Start at index 1 (the second cell)
    for (var i = 1; i < row.length; i += 2) {
      var key = row[i];
      // Stop processing if key is blank.
      if (key === null || key === undefined || key.toString().trim() === "") {
        break;
      }
      // Get the corresponding value (if available).
      var value = (i + 1 < row.length) ? row[i + 1] : null;
      obj[key] = value;
    }
    
    // Build the final object and return its JSON representation.
    var result = {};
    result[groupName] = obj;
    return JSON.stringify(result);
  }
