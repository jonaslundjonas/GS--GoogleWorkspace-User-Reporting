//Created by Jonas Lund 2023
// The script needs the AdminSDK AdminDirectory service enabled. 
//You need to be a super admin to enable that API in services.

// Function to prompt the user for the primary domain
function promptForDomain() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter the primary domain for the report (e.g., example.com):');
  var button = response.getSelectedButton();
  var domain = response.getResponseText();

  if (button == ui.Button.OK) {
    // Check if the entered domain is not empty
    if (domain.trim() !== '') {
      // Store the domain in the user properties
      PropertiesService.getUserProperties().setProperty('primaryDomain', domain);
      ui.alert('Domain set successfully.');
    } else {
      ui.alert('Domain cannot be empty. Please try again.');
    }
  }
}

// Function to list all users and populate the "userreport" sheet
function listAllUsers() {
  // Get the primary domain from user properties
  var domain = PropertiesService.getUserProperties().getProperty('primaryDomain');

  if (!domain || domain === 'yourdomain.com') {
    // If the domain is not set or set to 'yourdomain.com', prompt the user to change it
    promptForDomain();
    return;
  }

  // Get the currently active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('userreport');

  // Create the 'userreport' sheet if it doesn't exist
  if (!sheet) {
    sheet = spreadsheet.insertSheet('userreport');
  }

  // Clear existing data in the sheet
  sheet.clearContents();

  // Freeze the header row and make it bold
  sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
  sheet.setFrozenRows(1);

  var pageToken;
  var page;
  do {
    // Use the AdminDirectory service to list users
    page = AdminDirectory.Users.list({
      domain: domain, // Use the domain provided by the user
      maxResults: 100,
      pageToken: pageToken,
      fields: 'users(primaryEmail,name(fullName),lastLoginTime,agreedToTerms,creationTime,changePasswordAtNextLogin,isAdmin,isDelegatedAdmin,isEnrolledIn2Sv,suspended)',
    });

    var users = page.users;
    if (users) {
      // Add headers to the sheet before appending user data
      sheet.getRange(1, 1, 1, 10).setValues([['Email', 'Full Name', 'Last Login', 'Agreed to Terms', 'Creation Date', 'Change Password at Next Login', 'isAdmin', 'isDelegatedAdmin', 'isEnrolledIn2Sv', 'Suspended']]);

      for (var i = 0; i < users.length; i++) {
        var user = users[i];
        var changePassword = user.changePasswordAtNextLogin;
        var isAdmin = user.isAdmin ? 'True' : 'False';
        var isDelegatedAdmin = user.isDelegatedAdmin ? 'True' : 'False';
        var isEnrolledIn2Sv = user.isEnrolledIn2Sv ? 'True' : 'False';
        var suspended = user.suspended ? 'True' : 'False';

        // Append user data to the 'userreport' sheet
        sheet.appendRow([user.primaryEmail, user.name.fullName, user.lastLoginTime, user.agreedToTerms, user.creationTime, changePassword, isAdmin, isDelegatedAdmin, isEnrolledIn2Sv, suspended]);
      }
    } else {
      Logger.log('No users found.');
    }
    pageToken = page.nextPageToken; // Retrieve the next page of users, if available
  } while (pageToken);
}

// Function to clear existing data in the 'userreport' sheet
function clearData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('userreport');
  if (sheet) {
    sheet.clearContents();
  }
}

// Function to freeze the header row in the 'userreport' sheet
function freezeHeader() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('userreport');
  if (sheet) {
    sheet.setFrozenRows(1);
  }
}

// Function to make the text in the header row bold in the 'userreport' sheet
function makeHeaderBold() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('userreport');
  if (sheet) {
    var headerRange = sheet.getRange(1, 1, 1, 10); // Adjust the number of columns as needed
    headerRange.setFontWeight('bold');
  }
}

// Function to create a custom menu in the Google Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Create a custom menu called 'Report Script Menu'
  ui.createMenu('Report Script Menu')
    // Add an item 'Run User Report' that triggers the 'listAllUsers' function
    .addItem('Run User Report', 'listAllUsers')
    // Add an item 'Clear Data' that triggers the 'clearData' function
    .addItem('Clear Data', 'clearData')
    // Add an item 'Freeze Header' that triggers the 'freezeHeader' function
    .addItem('Freeze Header', 'freezeHeader')
    // Add an item 'Make Header Bold' that triggers the 'makeHeaderBold' function
    .addItem('Make Header Bold', 'makeHeaderBold')
    // Add an item 'Set Domain' that triggers the 'promptForDomain' function
    .addItem('Set Domain', 'promptForDomain')
    // Add the custom menu to the Google Sheet
    .addToUi();
}
