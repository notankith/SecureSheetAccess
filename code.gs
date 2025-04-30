// Global variables
const ACTIVE_USER_SHEET = "ActiveUser";
const LOG_SHEET = "UserLogs";
const INACTIVITY_TIMEOUT = 3 * 60 * 1000; // 3 minutes
const CREDENTIALS = {
  "user123": "pass123!",
  "admin456": "secure789@",
  "john_doe": "mypassword2025"
};
let timeoutId = null;
const ss = SpreadsheetApp.getActiveSpreadsheet();
const properties = PropertiesService.getScriptProperties();

// Function to initialize the script on spreadsheet open
function onOpen() {
  try {
    Logger.log('onOpen triggered');
    // Reset prompt lock on open
    properties.setProperty('promptLock', 'false');
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('User Login')
      .addItem('Login', 'showLoginPrompt')
      .addItem('Logout', 'logoutUser')
      .addToUi();
    
    protectAllSheets();
    initializeLogSheet();
    showLoginPrompt();
  } catch (e) {
    Logger.log('onOpen error: ' + e);
  }
}

// Function to initialize UserLogs sheet
function initializeLogSheet() {
  try {
    Logger.log('initializeLogSheet called');
    let logSheet = ss.getSheetByName(LOG_SHEET);
    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_SHEET);
      logSheet.getRange('A1:F1').setValues([['Timestamp', 'Event', 'Username', 'Sheet', 'Range', 'Details']]);
      logSheet.getRange('I1:J1').setValues([['Username', 'Login Count']]);
      logSheet.getRange('A1:J1').setFontWeight('bold');
    }
    updateLoginStats(logSheet);
  } catch (e) {
    Logger.log('initializeLogSheet error: ' + e);
  }
}

// Function to show login prompt
function showLoginPrompt() {
  try {
    Logger.log('showLoginPrompt called');
    // Check if prompt is already active
    const promptLock = properties.getProperty('promptLock');
    if (promptLock === 'true') {
      Logger.log('Prompt already active, skipping');
      return;
    }
    
    // Set lock
    properties.setProperty('promptLock', 'true');
    
    const activeUserSheet = ss.getSheetByName(ACTIVE_USER_SHEET);
    const currentUser = activeUserSheet ? activeUserSheet.getRange('B1').getValue() : null;
    if (currentUser) {
      Logger.log('User already logged in: ' + currentUser);
      properties.setProperty('promptLock', 'false');
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    Logger.log('Showing username prompt');
    const usernamePrompt = ui.prompt('User Login', 'Enter Username:', ui.ButtonSet.OK_CANCEL);
    
    if (usernamePrompt.getSelectedButton() !== ui.Button.OK) {
      Logger.log('Login prompt cancelled');
      ui.alert('Error', 'You must log in to use the spreadsheet.', ui.ButtonSet.OK);
      properties.setProperty('promptLock', 'false');
      return;
    }
    
    const username = usernamePrompt.getResponseText().trim();
    if (!username) {
      Logger.log('Empty username entered');
      ui.alert('Error', 'Username cannot be empty.', ui.ButtonSet.OK);
      properties.setProperty('promptLock', 'false');
      return;
    }
    
    Logger.log('Showing password prompt for: ' + username);
    const passwordPrompt = ui.prompt('User Login', `Enter Password for ${username}:`, ui.ButtonSet.OK_CANCEL);
    
    if (passwordPrompt.getSelectedButton() !== ui.Button.OK) {
      Logger.log('Password prompt cancelled');
      ui.alert('Error', 'You must log in to use the spreadsheet.', ui.ButtonSet.OK);
      properties.setProperty('promptLock', 'false');
      return;
    }
    
    const password = passwordPrompt.getResponseText().trim();
    if (!password) {
      Logger.log('Empty password entered');
      ui.alert('Error', 'Password cannot be empty.', ui.ButtonSet.OK);
      properties.setProperty('promptLock', 'false');
      return;
    }
    
    Logger.log('Validating credentials for: ' + username);
    if (validateLogin(username, password)) {
      Logger.log('Credentials valid, logging in: ' + username);
      loginUser(username);
      logEvent(username, 'Login', null, null, null);
      ui.alert('Success', `Logged in as ${username}`, ui.ButtonSet.OK);
    } else {
      Logger.log('Invalid credentials for: ' + username);
      ui.alert('Error', 'Invalid username or password.', ui.ButtonSet.OK);
    }
    
    // Release lock
    properties.setProperty('promptLock', 'false');
  } catch (e) {
    Logger.log('showLoginPrompt error: ' + e);
    properties.setProperty('promptLock', 'false');
    SpreadsheetApp.getUi().alert('Error', 'Login failed: ' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Function to validate login credentials
function validateLogin(username, password) {
  return CREDENTIALS.hasOwnProperty(username) && CREDENTIALS[username] === password;
}

// Function to log in user
function loginUser(username) {
  try {
    Logger.log('loginUser called for: ' + username);
    let activeUserSheet = ss.getSheetByName(ACTIVE_USER_SHEET);
    if (!activeUserSheet) {
      activeUserSheet = ss.insertSheet(ACTIVE_USER_SHEET);
    }
    
    activeUserSheet.getRange('A1:B2').setValues([
      ['Current User', username],
      ['Login Time', new Date()]
    ]);
    
    unprotectAllSheets();
    resetInactivityTimer();
    Logger.log('loginUser completed');
  } catch (e) {
    Logger.log('loginUser error: ' + e);
  }
}

// Function to log out user
function logoutUser() {
  if (!checkLoggedIn()) return;
  
  try {
    Logger.log('logoutUser called');
    const activeUserSheet = ss.getSheetByName(ACTIVE_USER_SHEET);
    const currentUser = activeUserSheet ? activeUserSheet.getRange('B1').getValue() : null;
    
    if (activeUserSheet) {
      activeUserSheet.getRange('B1:B2').clearContent();
    }
    
    protectAllSheets();
    if (timeoutId) {
      ScriptApp.getProjectTriggers().forEach(trigger => {
        if (trigger.getUniqueId() === timeoutId) {
          ScriptApp.deleteTrigger(trigger);
        }
      });
      timeoutId = null;
    }
    
    if (currentUser) {
      logEvent(currentUser, 'Logout', null, null, null);
    }
    
    SpreadsheetApp.getUi().alert('Logged out', 'You have been logged out', SpreadsheetApp.getUi().ButtonSet.OK);
    showLoginPrompt();
    Logger.log('Logout completed');
  } catch (e) {
    Logger.log('logoutUser error: ' + e);
  }
}

// Function to protect all sheets except ActiveUser and UserLogs
function protectAllSheets() {
  try {
    Logger.log('protectAllSheets called');
    const sheets = ss.getSheets();
    const currentUserNoEditorsEmail = Session.getActiveUser().getEmail();
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName !== ACTIVE_USER_SHEET && sheetName !== LOG_SHEET) {
        let protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0] || sheet.protect();
        protection.setDescription('Sheet locked for non-logged-in users')
                 .removeEditors(protection.getEditors());
        if (currentUserNoEditorsEmail) {
          protection.addEditor(currentUserNoEditorsEmail);
        }
      }
    });
    Logger.log('protectAllSheets completed');
  } catch (e) {
    Logger.log('protectAllSheets error: ' + e);
  }
}

// Function to unprotect all sheets
function unprotectAllSheets() {
  try {
    Logger.log('unprotectAllSheets called');
    const sheets = ss.getSheets();
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName !== ACTIVE_USER_SHEET && sheetName !== LOG_SHEET) {
        sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(protection => protection.remove());
      }
    });
    Logger.log('unprotectAllSheets completed');
  } catch (e) {
    Logger.log('unprotectAllSheets error: ' + e);
  }
}

// Function to check if a user is logged in
function checkLoggedIn() {
  try {
    Logger.log('checkLoggedIn called');
    const activeUserSheet = ss.getSheetByName(ACTIVE_USER_SHEET);
    const currentUser = activeUserSheet ? activeUserSheet.getRange('B1').getValue() : null;
    if (!currentUser) {
      SpreadsheetApp.getUi().alert('Error', 'Please log in to perform this action.', SpreadsheetApp.getUi().ButtonSet.OK);
      showLoginPrompt();
      Logger.log('checkLoggedIn: No user logged in');
      return false;
    }
    Logger.log('checkLoggedIn: User logged in: ' + currentUser);
    return true;
  } catch (e) {
    Logger.log('checkLoggedIn error: ' + e);
    return false;
  }
}

// Function to track cell edits
function onEdit(e) {
  try {
    const startTime = Date.now();
    Logger.log('onEdit triggered');
    
    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();
    
    if (sheetName === ACTIVE_USER_SHEET || sheetName === LOG_SHEET) return;
    
    const activeUserSheet = ss.getSheetByName(ACTIVE_USER_SHEET);
    const currentUser = activeUserSheet ? activeUserSheet.getRange('B1').getValue() : null;
    
    if (!currentUser) {
      Logger.log('Unauthorized edit detected, reverting');
      ss.toast('Edit reverted: Please log in to edit.', 'Error', 3);
      if (e.oldValue !== undefined) {
        range.setValue(e.oldValue).clearNote();
      } else {
        range.clear({ contentsOnly: true, skipFilteredRows: true }).clearNote();
      }
      Logger.log('Reversion completed in ' + (Date.now() - startTime) + 'ms');
      showLoginPrompt();
      return;
    }
    
    const note = `Edited by ${currentUser} at ${new Date().toLocaleString()}`;
    range.setNote(note);
    
    const oldValue = e.oldValue || '(empty)';
    const newValue = e.value || range.getValue();
    const details = `Old: ${oldValue}, New: ${newValue}`;
    logEvent(currentUser, 'Edit', sheetName, range.getA1Notation(), details);
    
    resetInactivityTimer();
  } catch (e) {
    Logger.log('onEdit error: ' + e);
  }
}

// Function to reset inactivity timer
function resetInactivityTimer() {
  try {
    Logger.log('resetInactivityTimer called');
    if (timeoutId) {
      ScriptApp.getProjectTriggers().forEach(trigger => {
        if (trigger.getUniqueId() === timeoutId) {
          ScriptApp.deleteTrigger(trigger);
        }
      });
    }
    
    const trigger = ScriptApp.newTrigger('logoutUser')
      .timeBased()
      .after(INACTIVITY_TIMEOUT)
      .create();
    timeoutId = trigger.getUniqueId();
    Logger.log('resetInactivityTimer completed');
  } catch (e) {
    Logger.log('resetInactivityTimer error: ' + e);
  }
}

// Function to log events (login, logout, edit)
function logEvent(username, event, sheetName, range, details) {
  try {
    Logger.log('logEvent called: ' + event + ' for ' + username);
    const logSheet = ss.getSheetByName(LOG_SHEET);
    if (!logSheet) return;
    
    logSheet.appendRow([new Date(), event, username, sheetName || '', range || '', details || '']);
    updateLoginStats(logSheet);
    Logger.log('logEvent completed');
  } catch (e) {
    Logger.log('logEvent error: ' + e);
  }
}

// Function to update login stats
function updateLoginStats(logSheet) {
  try {
    Logger.log('updateLoginStats called');
    const data = logSheet.getRange('A2:C' + logSheet.getLastRow()).getValues();
    const loginCounts = {};
    
    data.forEach(row => {
      if (row[1] === 'Login') {
        const username = row[2];
        loginCounts[username] = (loginCounts[username] || 0) + 1;
      }
    });
    
    logSheet.getRange('I2:J' + logSheet.getLastRow()).clearContent();
    const stats = Object.entries(loginCounts);
    if (stats.length > 0) {
      logSheet.getRange('I2:J' + (stats.length + 1)).setValues(stats);
    }
    Logger.log('updateLoginStats completed');
  } catch (e) {
    Logger.log('updateLoginStats error: ' + e);
  }
}
