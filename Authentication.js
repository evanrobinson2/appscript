function getSalesforceAccessToken() {
  const props = PropertiesService.getScriptProperties();
  
  // Retrieve Salesforce credentials from Script Properties
  const instanceUrl   = props.getProperty('SF_INSTANCE_URL');      // e.g., "https://myDomain.my.salesforce.com"
  const clientId      = props.getProperty('SF_CLIENT_ID');         // Consumer Key
  const clientSecret  = props.getProperty('SF_CLIENT_SECRET');     // Consumer Secret

  // Build the token endpoint URL
  const tokenUrl = instanceUrl + 'services/oauth2/token';
  Logger.log('Requesting token from URL: ' + tokenUrl);

  // Payload for username-password flow
  const payload = {
    'grant_type':    'client_credentials',
    'client_id':     clientId,
    'client_secret': clientSecret,
  };

  // Configure the request
  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  // Make the call
  let response;
  try {
    response = UrlFetchApp.fetch(tokenUrl, options);
  } catch (e) {
    Logger.log('Error making token request: ' + e);
    throw new Error('Failed to reach Salesforce token endpoint: ' + e);
  }
  
  // Log the HTTP status code
  Logger.log('Response Code: ' + response.getResponseCode());

  // Parse the JSON response
  const body = response.getContentText();
  Logger.log('Response Body: ' + body);
  
  let json;
  try {
    json = JSON.parse(body);
  } catch (e) {
    Logger.log('JSON parse error: ' + e);
    throw new Error('Invalid JSON response from token request: ' + e);
  }

  // Check for access token
  if (json.access_token) {
    Logger.log('Successfully retrieved access token: ' + json.access_token);
    return json.access_token;
  } else {
    Logger.log('Failed to get access token. Full response: ' + body);
    throw new Error('Failed to get access token: ' + body);
  }
}