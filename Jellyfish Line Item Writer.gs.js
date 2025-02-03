/**
 * processOliRecords(ss)
 *
 * End-to-end function to process Opportunity Line Items (OLIs).
 *
 * Steps:
 * 1. Load parameters from the "JF_SCRIPT_PARAMS" sheet.
 * 2. Build the in-memory OLI table from the input sheet using the parameter mappings.
 * 3. Deactivate any existing active OLIs for the Opportunity.
 * 4. Retrieve the highest existing revision number for OLIs and compute the new revision.
 * 5. For each new OLI record, set:
 *    - Opportunity__c to the given Opportunity Id.
 *    - Active__c to true.
 *    - Version_Number__c to (highest revision + 1).
 * 6. Insert the new OLI records into Salesforce.
 *
 * This function leverages:
 *    - exampleLoadParameters() [from your parameter loader code]
 *    - buildOliDataTableFromParams() [from your OLI builder code]
 *    - Library functions from Jellyfish Revops SFDC Integrations:
 *         getHighestRevisionNumber(oppId)
 *         deactivateAllActiveLineItems(oppId)
 *         createLineItems(oppId, lineItems)
 *
 * @returns {Object} An object summarizing the deactivation and insertion results.
 */
function processOliRecords(ss) {
  try {
    // STEP 1: Load parameters from the parameter sheet.
    // This function returns an object of parameters (including Input Sheet, Table Header Row, and oli mappings).
    var params = exampleLoadParameters(ss);
    // logMessage(ss,"Parameters loaded successfully: " + JSON.stringify(params, null, 2));
    
    // STEP 2: Build the in-memory OLI table from the input sheet.
    var oliData = buildDataTableFromParamsDynamic(ss);
    logMessage(ss,"OLI Data Table built successfully. Number of records: " + oliData.length);
    
    var oppId = oliData[0]["jellyfish_line_item__c"]["opportunity_id__c"];
    logMessage(ss,'Found Opportunity ID: ' + oppId)
    // STEP 3: Query and deactivate existing active OLIs.
    // Note: The library functions are referenced via the attached library namespace.
    
    var deactivationResults = deactivateAllActiveLineItems(oppId);
    logMessage(ss,"Deactivation results:\n" + JSON.stringify(deactivationResults, null, 2));
    
    // STEP 4: Retrieve the highest revision number from Salesforce.
    var highestRevision = getHighestRevisionNumber(oppId);

    var newRevision = highestRevision + 1;
    logMessage(ss,"Highest Revision: " + highestRevision + ", New Revision: " + newRevision);
    
    // STEP 5: Augment each new OLI record with necessary fields.
    for (var i = 0; i < oliData.length; i++) {
      var record = oliData[i];
      // Ensure the Opportunity__c field is set.
      record.opportunity_id__c = oppId;
      // Set the Active__c flag to true.
      record.Active__c = true;
      // Set the Version_Number__c field to the new revision number.
      record.Version_Number__c = newRevision;
    }
    logMessage(ss,"New OLI records after augmentation:\n" + JSON.stringify(oliData, null, 2));
    
    // STEP 6: Insert new OLI records into Salesforce.

    formattedOliData = formatOLIs(oliData); 
    logMessage(ss,formattedOliData)

    var insertionResults = createLineItems(oppId, formattedOliData);
    
    logMessage(ss,"Insertion results:\n" + JSON.stringify(insertionResults, null, 2));
    
    // Return a summary of the operations.
    var resultSummary = {
      deactivation: deactivationResults,
      insertion: insertionResults
    };
    logMessage(ss,"Process completed successfully:\n" + JSON.stringify(resultSummary, null, 2));
    return resultSummary;
    
  } catch (e) {
    logMessage(ss,"Error in processOliRecords: " + e.message);
    throw e;
  }
}

/**
 * Flattens the records produced by buildDataTableFromParamsDynamic().
 *
 * Each record from buildDataTableFromParamsDynamic() may have nested groups (for example,
 * under the key "jellyfish_line_item__c") that hold the field mappings. This function merges
 * those nested objects into the top-level record.
 *
 * @param {Array<Object>} oliData - The array of OLI records as built by buildDataTableFromParamsDynamic().
 * @returns {Array<Object>} - An array of flattened OLI records, ready to be sent to createLineItems().
 */
function formatOLIs(oliData) {
  var formatted = [];
  for (var i = 0; i < oliData.length; i++) {
    var record = oliData[i];
    var flatRecord = {};
    
    // Loop through each key in the record.
    for (var key in record) {
      if (record.hasOwnProperty(key)) {
        // If the value is an object (and not null), merge its properties.
        // This assumes that any nested object (e.g. "jellyfish_line_item__c") should be flattened.
        if (typeof record[key] === 'object' && record[key] !== null) {
          for (var nestedKey in record[key]) {
            if (record[key].hasOwnProperty(nestedKey)) {
              flatRecord[nestedKey] = record[key][nestedKey];
            }
          }
        } else {
          // Otherwise, keep the top-level property.
          flatRecord[key] = record[key];
        }
      }
    }
    formatted.push(flatRecord);
  }
  return formatted;
}


/**
 * Example usage: Call this function to run the entire end-to-end operation.
 */
function processJellyfishLineItems(ss) {
  var startTime = new Date(); // Capture the start time
  logMessage(ss,"Starting processJellyfishLineItems at " + startTime.toISOString());

  var results = processOliRecords(ss);

  var endTime = new Date(); // Capture the end time
  var executionTime = (endTime - startTime) / 1000; // Convert to seconds

  logMessage(ss,"Completed processJellyfishLineItems at " + endTime.toISOString());
  logMessage(ss,"Total Execution Time: " + executionTime.toFixed(2) + " seconds");
  logMessage(ss, "Final results:\n" + JSON.stringify(results, null, 2));
  
  flushLogs(ss);
  return results;
}
