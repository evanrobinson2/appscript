/**
 * Creates one or more line items for a given Opportunity.
 * @param {String} oppId Salesforce Opportunity Id (e.g. "006XXXXXXXXXXXX").
 * @param {Array<Object>} lineItems An array of JS objects with all necessary fields.
 *   Example lineItem object:
 *   {
 *     Product__c: '01tXXXXXXXXXXXX',    // required
 *     Price_Book__c: '01sXXXXXXXXXXXX', // required
 *     Quantity__c: 10,
 *     List_Price__c: 500.00,
 *     Sales_Price__c: 450.00,
 *     Active__c: true,
 *     Start_Date__c: '2025-02-01',
 *     End_Date__c: '2025-12-31',
 *     // etc.
 *   }
 * @returns {Array} Array of results from Salesforce for each inserted record.
 */
function createLineItems(oppId, lineItems) {
  const token = getSalesforceAccessToken();
  const instanceUrl = PropertiesService.getScriptProperties().getProperty('SF_INSTANCE_URL');
  const baseUrl = instanceUrl + '/services/data/v58.0/sobjects/jellyfish_line_item__c';

  const results = [];

  lineItems.forEach(item => {
    // Add the opportunity_id__c field if not already included
    item.opportunity_id__c = oppId;
    item.Sales_Discount__c = 100 * item.Sales_Discount__c;
    const options = {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(item),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(baseUrl, options);
    const json = JSON.parse(response.getContentText());
    results.push(json);
  });

  return results;
}

/**
 * Gets the highest revision (Version_Number__c) among line items for the given Opportunity.
 * @param {String} oppId - The Salesforce Opportunity Id (e.g., "006XXXXXXXXXXXX").
 * @returns {Number} The highest version number found, or 0 if none exist.
 */
function getHighestRevisionNumber(oppId) {
  // 1) Obtain token and instance URL
  const token = getSalesforceAccessToken();
  const instanceUrl = PropertiesService.getScriptProperties().getProperty('SF_INSTANCE_URL');

  // 2) Build SOQL query: sort descending by Version_Number__c, grab top 1
  const soql = `
    SELECT Version_Number__c
    FROM jellyfish_line_item__c
    WHERE opportunity_id__c = '${oppId}'
    ORDER BY Version_Number__c DESC
    LIMIT 1
  `;
  const queryUrl = instanceUrl + '/services/data/v58.0/query?q=' + encodeURIComponent(soql);

  // 3) Make the query call
  const options = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(queryUrl, options);
  const body = JSON.parse(response.getContentText());

  // 4) Check if we got any records
  if (body.records && body.records.length > 0) {
    // Return the highest version number
    return body.records[0].Version_Number__c || 0;
  } else {
    // If no records, return 0
    return 0;
  }
}

/**
 * Deactivate all active jellyfish_line_item__c records for a given Opportunity.
 * @param {String} oppId - The Salesforce Opportunity Id (e.g. "006XXXXXXXXXXXX").
 * @returns {Array} Array of results indicating the status for each record updated.
 */
function deactivateAllActiveLineItems(oppId) {
  // 1) Get the token and instance URL
  const token = getSalesforceAccessToken();
  const instanceUrl = PropertiesService.getScriptProperties().getProperty('SF_INSTANCE_URL');

  // 2) Query to find all active line items for the specified Opp
  const soql = `
    SELECT Id
    FROM jellyfish_line_item__c
    WHERE opportunity_id__c = '${oppId}'
    AND Active__c = true
  `;
  const queryUrl = instanceUrl + '/services/data/v58.0/query?q=' + encodeURIComponent(soql);

  // 3) Execute the query
  const queryOptions = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + token
    },
    muteHttpExceptions: true
  };
  const queryResponse = UrlFetchApp.fetch(queryUrl, queryOptions);
  const queryBody = JSON.parse(queryResponse.getContentText());
  
  // If no records or an error occurred, return early
  if (!queryBody.records || !Array.isArray(queryBody.records)) {
    return [{
      status: queryResponse.getResponseCode(),
      error: queryBody
    }];
  }

  // 4) For each active line item, PATCH Active__c = false
  const results = [];
  for (const record of queryBody.records) {
    const lineItemId = record.Id;
    const url = instanceUrl + '/services/data/v58.0/sobjects/jellyfish_line_item__c/' + lineItemId;
    
    const payload = { Active__c: false };
    const patchOptions = {
      method: 'patch',
      headers: {
        Authorization: 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const patchResponse = UrlFetchApp.fetch(url, patchOptions);
    results.push({
      id: lineItemId,
      status: patchResponse.getResponseCode(),
      body: patchResponse.getContentText()
    });
  }

  // 5) Return an array of result objects
  return results;
}


