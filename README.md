SheetGuard
SheetGuard is a Google Apps Script that secures Google Spreadsheets with strict user authentication, fast edit reversion, a detailed audit trail, and automatic logout after inactivity. Designed for shared environments, it ensures only authorized users can edit the spreadsheet, with all actions logged for accountability.
Features

Strict Login Enforcement: Requires username/password login to access the spreadsheet; unauthorized users are blocked.
Fast Edit Reversion: Instantly reverts unauthorized edits with a toast notification.
Enhanced Audit Trail: Logs all actions (login, logout, edits) with timestamps, usernames, sheet names, ranges, and old/new values.
Inactivity Timeout: Automatically logs out users after 3 minutes of inactivity, re-protecting the spreadsheet.
Single Prompt Control: Prevents multiple login dialogs from overlaying, ensuring a smooth user experience.
Shared Environment Support: Works for multiple users (e.g., 7–8) sharing one Google account on a single computer.
Error-Free Operation: Handles edge cases like corrupted sheets, invalid credentials, and script errors gracefully.

Installation

Open Your Spreadsheet:

Create or open a Google Spreadsheet where you want to implement SheetGuard.


Access Apps Script:

Go to Extensions > Apps Script in the Google Sheets menu.


Copy the Script:

Copy the contents of Code.gs from this repository.
Paste it into the Apps Script editor, replacing any existing code.
Save the file as Code.gs.


Configure Credentials:

In Code.gs, locate the CREDENTIALS object:const CREDENTIALS = {
  "user123": "pass123!",
  "admin456": "secure789@",
  "john_doe": "mypassword2025"
};


Replace with your desired usernames and passwords.


Set Up the Trigger:

In the Apps Script editor, click the Triggers icon (clock).
Click + Add Trigger.
Configure:
Function: onOpen
Event source: From spreadsheet
Event type: On open


Save and authorize the script (sign in with your Google account and allow permissions).


Test the System:

Reload the spreadsheet. A login prompt should appear.
Enter valid credentials to unlock the spreadsheet.
Test unauthorized edits (should revert instantly), logout, and inactivity timeout.



Usage

Login: On opening the spreadsheet or after logout, enter a username and password. Only valid credentials (defined in CREDENTIALS) unlock the spreadsheet.
Editing: Authorized users can edit any sheet except ActiveUser and UserLogs. Edits are logged with details (sheet, range, old/new values).
Logout: Use User Login > Logout in the menu to lock the spreadsheet and require re-login.
Audit Trail: The UserLogs sheet tracks all actions:
Columns A–F: Timestamp, Event (Login/Logout/Edit), Username, Sheet, Range, Details (e.g., Old: 123, New: 456).
Columns I–J: Username and login count.


Inactivity: After 3 minutes of no edits, the script logs out, protects sheets, and shows a login prompt.

Troubleshooting

Login Fails or Prompt Doesn’t Work:
Check View > Logs in the Apps Script editor for errors (e.g., showLoginPrompt error: [message]).
Reset the prompt lock:function resetPromptLock() {
  PropertiesService.getScriptProperties().setProperty('promptLock', 'false');
  Logger.log('Prompt lock reset');
}


Run this in the Apps Script editor, then reload the spreadsheet.


Reauthorize permissions in your Google Account (https://myaccount.google.com/security, remove “Google Apps Script,” then reauthorize).


Edits Not Reverting Instantly:
Verify logs for Reversion completed in Xms. If >500ms, check for server-side delays or contact support.


Logout Errors:
Ensure no manual changes to ActiveUser or UserLogs sheets, as they store state.


General Issues:
Test in Chrome (incognito mode) to rule out browser extensions.
Create a new spreadsheet and copy the script to rule out corruption.



Security Notes

Hardcoded Credentials: Passwords in CREDENTIALS are stored in plain text, which is insecure. Consider:
Storing hashed passwords in a hidden sheet.
Implementing an admin UI for credential management.


Shared Account: Designed for multiple users on one Google account. Ensure users don’t share credentials outside the intended group.
Audit Trail: Regularly review UserLogs to monitor access and edits.

Contributing
Contributions are welcome! To contribute:

Fork the repository.
Create a feature branch (git checkout -b feature-name).
Commit changes (git commit -m "Add feature").
Push to the branch (git push origin feature-name).
Open a pull request with a description of changes.

Please include tests and update documentation for new features.
License
This project is licensed under the MIT License. See the LICENSE file for details.
Contact
For issues, suggestions, or questions, open an issue on this repository or contact the maintainer via GitHub.

Built for secure, shared spreadsheet management. Protect your data with SheetGuard!
