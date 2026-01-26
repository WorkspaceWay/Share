/**
 * @OnlyCurrentDoc
 */

const MAX_BATCH_SIZE = 100;
const DELAY_BETWEEN_BATCHES = 20000;

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Validation')
      .addItem('Validate Emails', 'validateEmailsWithZeroBounce')
      .addItem('Check API Credits', 'checkZeroBounceCredits')
      .addToUi();
}

function getZeroBounceApiKey() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('ZEROBOUNCE_API_KEY');
  
  if (!apiKey) {
    var errorMsg = 'ZeroBounce API Key not configured. Please set ZEROBOUNCE_API_KEY in Project Settings > Script Properties.';
    Logger.log('ERROR: ' + errorMsg);
    SpreadsheetApp.getUi().alert('Configuration Error', errorMsg, SpreadsheetApp.getUi().ButtonSet.OK);
    throw new Error(errorMsg);
  }
  
  return apiKey;
}

function validateEmailsWithZeroBounce() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('=== Starting Email Validation Process ===');
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    Logger.log('Working on sheet: ' + sheet.getName());
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No data rows found');
      ui.alert('No Data', 'No data rows found in the sheet.', ui.ButtonSet.OK);
      return;
    }
    
    var emailCol = getColumnIndexByName(sheet, 'Email');
    var statusCol = getColumnIndexByName(sheet, 'Status');
    
    if (emailCol === -1 || statusCol === -1) {
      var errorMsg = 'Email or Status column not found. Please ensure your sheet has columns named "Email" and "Status".';
      Logger.log('ERROR: ' + errorMsg);
      ui.alert('Column Error', errorMsg, ui.ButtonSet.OK);
      return;
    }
    
    Logger.log('Email column: ' + emailCol + ', Status column: ' + statusCol);
    
    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    
    var rowsToValidate = [];
    for (var i = 0; i < data.length; i++) {
      var email = data[i][emailCol - 1];
      var status = data[i][statusCol - 1];
      
      if (email && email.toString().trim() !== '' && 
          (!status || status.toString().trim() === '')) {
        rowsToValidate.push({
          rowIndex: i + 2,
          email: email.toString().trim()
        });
      }
    }
    
    Logger.log('Found ' + rowsToValidate.length + ' emails to validate');
    
    if (rowsToValidate.length === 0) {
      Logger.log('No emails need validation');
      ui.alert('Nothing to Validate', 'No emails need validation. All rows either have a status or no email.', ui.ButtonSet.OK);
      return;
    }
    
    var totalBatches = Math.ceil(rowsToValidate.length / MAX_BATCH_SIZE);
    Logger.log('Will process ' + totalBatches + ' batch(es)');
    
    var response = ui.alert(
      'Confirm Validation',
      'Found ' + rowsToValidate.length + ' email(s) to validate.\n\n' +
      'This will be processed in ' + totalBatches + ' batch(es).\n' +
      (totalBatches > 1 ? 'Total time: approximately ' + Math.ceil((totalBatches - 1) * 20 / 60) + ' minutes.\n\n' : '\n') +
      'Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      Logger.log('User cancelled validation');
      return;
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Validating ' + rowsToValidate.length + ' emails...', 'Processing', -1);
    
    var successCount = 0;
    var errorCount = 0;
    var batchErrors = [];
    
    for (var batchStart = 0; batchStart < rowsToValidate.length; batchStart += MAX_BATCH_SIZE) {
      var batchEnd = Math.min(batchStart + MAX_BATCH_SIZE, rowsToValidate.length);
      var batch = rowsToValidate.slice(batchStart, batchEnd);
      var batchNumber = Math.floor(batchStart / MAX_BATCH_SIZE) + 1;
      
      Logger.log('=== Processing batch ' + batchNumber + ' of ' + totalBatches + 
                 ' (emails ' + (batchStart + 1) + ' to ' + batchEnd + ') ===');
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Processing batch ' + batchNumber + ' of ' + totalBatches + '...',
        'Validating Emails',
        -1
      );
      
      try {
        var results = validateBatch(batch);
        var updated = updateSheetWithResults(sheet, statusCol, batch, results);
        successCount += updated;
        
        SpreadsheetApp.flush();
        
        Logger.log('Batch ' + batchNumber + ' completed: ' + updated + ' rows updated');
        
      } catch (batchError) {
        Logger.log('ERROR in batch ' + batchNumber + ': ' + batchError.toString());
        errorCount += batch.length;
        batchErrors.push('Batch ' + batchNumber + ': ' + batchError.toString());
        
        var continueResponse = ui.alert(
          'Batch Error',
          'Error processing batch ' + batchNumber + ' of ' + totalBatches + ':\n\n' + 
          batchError.toString() + '\n\nContinue with remaining batches?',
          ui.ButtonSet.YES_NO
        );
        
        if (continueResponse !== ui.Button.YES) {
          Logger.log('User cancelled after batch error');
          SpreadsheetApp.getActiveSpreadsheet().toast('Validation cancelled', 'Stopped', 3);
          
          var cancelSummary = 'Validation stopped by user.\n\n' +
                            '✓ Successfully validated: ' + successCount + ' emails\n' +
                            '✗ Failed: ' + errorCount + ' emails\n' +
                            'Remaining: ' + (rowsToValidate.length - batchEnd) + ' emails';
          
          ui.alert('Cancelled', cancelSummary, ui.ButtonSet.OK);
          return;
        }
      }
      
      if (batchEnd < rowsToValidate.length) {
        Logger.log('Waiting 20 seconds before next batch...');
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Waiting 20 seconds before next batch (API rate limit)...',
          'Paused',
          20
        );
        Utilities.sleep(DELAY_BETWEEN_BATCHES);
      }
    }
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Validation complete!', 'Done', 3);
    
    var summary = 'Validation Complete!\n\n' +
                  '✓ Successfully validated: ' + successCount + ' emails\n' +
                  (errorCount > 0 ? '✗ Errors: ' + errorCount + ' emails\n' : '') +
                  '\nCheck the Status column for results.';
    
    if (batchErrors.length > 0) {
      summary += '\n\nErrors occurred:\n' + batchErrors.join('\n');
    }
    
    Logger.log('=== ' + summary + ' ===');
    
    ui.alert(errorCount > 0 ? 'Completed with Errors' : 'Complete', summary, ui.ButtonSet.OK);
    
  } catch (error) {
    var errorMsg = 'Fatal error: ' + error.toString();
    Logger.log('FATAL ERROR: ' + errorMsg);
    Logger.log('Stack trace: ' + error.stack);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Error occurred', 'Failed', 3);
    
    ui.alert(
      'Fatal Error',
      'An error occurred during validation:\n\n' + error.toString() + '\n\nPlease check the execution logs for more details.',
      ui.ButtonSet.OK
    );
  }
}

function getColumnIndexByName(sheet, columnName) {
  try {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var index = headers.indexOf(columnName);
    return index === -1 ? -1 : index + 1;
  } catch (error) {
    Logger.log('ERROR in getColumnIndexByName: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Error Reading Headers',
      'Failed to read column headers:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

function validateBatch(batch) {
  try {
    var apiKey = getZeroBounceApiKey();
    var emails = batch.map(function(item) { return item.email; });
    
    Logger.log('Sending ' + emails.length + ' emails to ZeroBounce API');
    
    var url = 'https://bulkapi.zerobounce.net/v2/validatebatch';
    
    var payload = {
      'api_key': apiKey,
      'email_batch': emails.map(function(email, index) {
        return {
          'email_address': email,
          'ip_address': ''
        };
      })
    };
    
    var options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    Logger.log('API Response Code: ' + responseCode);
    
    if (responseCode === 200) {
      var result = JSON.parse(responseText);
      
      if (result.email_batch) {
        Logger.log('Successfully received ' + result.email_batch.length + ' results');
        return result.email_batch;
      } else if (result.errors) {
        Logger.log('API returned errors: ' + JSON.stringify(result.errors));
        throw new Error('ZeroBounce API errors: ' + JSON.stringify(result.errors));
      } else {
        Logger.log('Unexpected API response: ' + responseText);
        throw new Error('Unexpected API response format');
      }
    } else if (responseCode === 401) {
      throw new Error('Invalid API key. Please check your ZEROBOUNCE_API_KEY in Script Properties.');
    } else if (responseCode === 429) {
      throw new Error('Rate limit exceeded. Please wait and try again later.');
    } else {
      Logger.log('API Error Response: ' + responseText);
      throw new Error('API error (code ' + responseCode + '): ' + responseText);
    }
  } catch (error) {
    Logger.log('Error in validateBatch: ' + error.toString());
    throw error;
  }
}

function updateSheetWithResults(sheet, statusCol, batch, results) {
  try {
    if (!results || results.length === 0) {
      Logger.log('WARNING: No results to update');
      return 0;
    }
    
    var statusMap = {};
    results.forEach(function(result) {
      if (result.address) {
        statusMap[result.address.toLowerCase()] = result.status || 'unknown';
      }
    });
    
    var updateCount = 0;
    var updateErrors = [];
    
    batch.forEach(function(item) {
      var status = statusMap[item.email.toLowerCase()];
      if (status) {
        try {
          sheet.getRange(item.rowIndex, statusCol).setValue(status);
          Logger.log('Updated row ' + item.rowIndex + ' (' + item.email + ') with status: ' + status);
          updateCount++;
        } catch (e) {
          var errorMsg = 'Failed to update row ' + item.rowIndex + ': ' + e.toString();
          Logger.log('ERROR: ' + errorMsg);
          updateErrors.push(errorMsg);
        }
      } else {
        Logger.log('WARNING: No status found for ' + item.email);
      }
    });
    
    if (updateErrors.length > 0) {
      SpreadsheetApp.getUi().alert(
        'Update Errors',
        'Some rows could not be updated:\n\n' + updateErrors.slice(0, 5).join('\n') +
        (updateErrors.length > 5 ? '\n\n... and ' + (updateErrors.length - 5) + ' more errors' : ''),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    return updateCount;
  } catch (error) {
    Logger.log('ERROR in updateSheetWithResults: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Update Error',
      'Failed to update sheet with results:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

function checkZeroBounceCredits() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var apiKey = getZeroBounceApiKey();
    var url = 'https://api.zerobounce.net/v2/getcredits?api_key=' + apiKey;
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Checking credits...', 'Please wait', -1);
    
    var response = UrlFetchApp.fetch(url);
    var result = JSON.parse(response.getContentText());
    
    Logger.log('Remaining Credits: ' + result.Credits);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Credits: ' + result.Credits, 'ZeroBounce', 3);
    ui.alert('ZeroBounce Credits', 'Remaining Credits: ' + result.Credits, ui.ButtonSet.OK);
    
    return result.Credits;
  } catch (error) {
    Logger.log('ERROR checking credits: ' + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast('Error checking credits', 'Failed', 3);
    ui.alert('Error Checking Credits', 'Failed to check credits:\n\n' + error.toString(), ui.ButtonSet.OK);
  }
}
