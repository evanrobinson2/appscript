/**
 * Logs a message to both Logger and the "JF_SCRIPT_LOG" sheet.
 * If the log sheet does not exist, it is created and headers are added.
 *
 * @param {String} message The message to log.
 */

var DEBUG_MODE = true; // Set to true for debugging

var logBuffer = [];

function logMessage(ss, message) {
  if (DEBUG_MODE) {
    Logger.log(message);
    // Optionally, write to a custom log sheet if needed.
    var logSheet = ss.getSheetByName("JF_SCRIPT_LOG");
    
    if (!logSheet) {
      logSheet = ss.insertSheet("JF_SCRIPT_LOG");
      logSheet.appendRow(["Timestamp", "Message"]);
    }
    
    var timestamp = new Date();
    logSheet.appendRow([timestamp, message]);
  }
  logBuffer.push({
    timestamp: new Date(),
    message: message
  });
}

function flushLogs(ss) {
  if (logBuffer.length > 0) {
    var logSheet = ss.getSheetByName("JF_SCRIPT_LOG");
    if (!logSheet) {
      logSheet = ss.insertSheet("JF_SCRIPT_LOG");
      logSheet.appendRow(["Timestamp", "Message"]);
    }
    
    // Map each log entry to a row, preserving the original timestamp.
    var logRows = logBuffer.map(function(logEntry) {
      return [logEntry.timestamp, logEntry.message];
    });
    
    // Append the rows starting from the next empty row.
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, logRows[0].length)
            .setValues(logRows);
    
    // Clear the buffer after flushing.
    logBuffer = [];
  }
}