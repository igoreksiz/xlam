## Test Procedure for finbox.io Excel Add-In

### Platforms To Test

* Excel 2016 - Windows (64-bit)
* Excel 2016 - Windows (32-bit)
* Excel 2013 - Windows (32-bit)
* Excel 2010 - Windows (32-bit)
* Excel 2007 - Windows (32-bit)
* Excel 2016 - Mac
* Excel 2011 - Mac


### Test Procedures

Unless otherwise specified, all tests assume that the version of the add-in
you are testing is installed, that 'Automatic' calculation is enabled, and
that a premium user is logged in.

Unless otherwise specified, if you are asked to update links when
opening a workbook, choose 'Ignore Links' (or 'Dont Update')

Unless otherwise specified, exit Excel completely at the end of each test
and do not save any changes made to test workbooks.


#### Installation

###### Excel Add-In should succesfully install on a clean system

- [ ] VERIFY that any previous add-in installation is completely removed.
- [ ] Open Excel Add-In workbook.
- [ ] Accept "Enable Macros" prompt.
- [ ] VERIFY user is prompted to approve installation.
- [ ] Accept installation prompt.
- [ ] VERIFY that user is prompted to quit excel.
- [ ] Accept prompt to quit excel.
- [ ] Restart Excel.
- [ ] Open a new workbook.
- [ ] VERIFY that finbox.io ribbon is present.
- [ ] VERIFY that finboxio.xlam is present in add-in folder.
- [ ] VERIFY that FNBX formula is available.

###### Excel Add-In should successfully overwrite an existing installation

 - [ ] VERIFY that a previous add-in is installed (follow documented installation procedure).
 - [ ] Open Excel Add-In workbook.
 - [ ] Accept "Enable Macros" prompt.
 - [ ] VERIFY user is prompted to approve installation.
 - [ ] Accept installation prompt.
 - [ ] VERIFY that user is prompted to quit excel.
 - [ ] Accept prompt to quit excel.
 - [ ] Restart excel.
 - [ ] VERIFY that finbox.io ribbon is present.
 - [ ] VERIFY that finboxio.xlam is present in add-in folder.
 - [ ] VERIFY that FNBX formula is available.
 - [ ] VERIFY that older version was overwritten and only latest install is present.

###### Excel Add-In should not install if user declines installation prompt

 - [ ] Open Excel Add-In workbook.
 - [ ] Accept "Enable Macros" prompt.
 - [ ] VERIFY user is prompted to approve installation.
 - [ ] Decline installation prompt.
 - [ ] VERIFY that no more prompts are displayed.
 - [ ] VERIFY that finboxio.xlam is not present in add-in folder.
 - [ ] Restart excel.
 - [ ] VERIFY that finbox.io ribbon is not present after restart.


#### Authentication

###### Excel Add-In should allow authentication with email/password

 - [ ] Open Excel.
 - [ ] Ensure no user is logged in (if user is logged in, log out and restart excel).
 - [ ] VERIFY that the "Log In" button is available in the finbox.io ribbon.
 - [ ] VERIFY that the "Log Out" button is NOT available in the finbox.io ribbon.
 - [ ] Select "Log In" from finbox.io ribbon.
 - [ ] VERIFY that login form is displayed with email and password fields.
 - [ ] Enter valid premium login credentials.
 - [ ] VERIFY that password entry is obscured.
 - [ ] Click "Login" button.
 - [ ] VERIFY that the login prompt is closed.
 - [ ] VERIFY that the "Log In" button is removed from the finbox.io ribbon.
 - [ ] VERIFY that the "Log Out" button is added to the finbox.io ribbon.
 - [ ] Add a FNBX formula to the workbook using a restricted company.
 - [ ] VERIFY that the formula returns a value.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a message exists with the correct email, api key, and usage tier for the logged-in user.
 - [ ] VERIFY that the user api key is stored on the filesystem alongside the add-on.

###### Excel Add-In login form should include a link to sign up

 - [ ] Open Excel.
 - [ ] Ensure no user is logged in (if user is logged in, log out and restart excel).
 - [ ] Select "Log In" from finbox.io ribbon.
 - [ ] VERIFY that the signup page is linked to from the login form.

###### Excel Add-In should submit the login form on 'Enter'

 - [ ] Open Excel.
 - [ ] Ensure no user is logged in (if user is logged in, log out and restart excel).
 - [ ] Select "Log In" from finbox.io ribbon.
 - [ ] VERIFY that login form is displayed with email and password fields.
 - [ ] Enter valid login credentials.
 - [ ] Press 'Enter'.
 - [ ] VERIFY that the login prompt is closed.
 - [ ] VERIFY that the "Log In" button is removed from the finbox.io ribbon.
 - [ ] VERIFY that the "Log Out" button is added to the finbox.io ribbon.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a message exists with the correct email, api key, and usage tier for the logged-in user.
 - [ ] VERIFY that the user api key is stored on the filesystem alongside the add-on.

###### Excel Add-In should warn on login attempts with unrecognized emails

 - [ ] Open Excel.
 - [ ] Ensure no user is logged in (if user is logged in, log out and restart excel).
 - [ ] Select "Log In" from finbox.io ribbon.
 - [ ] VERIFY that login form is displayed with email and password fields.
 - [ ] Enter login credentials with an invalid email.
 - [ ] Click "Login" button.
 - [ ] VERIFY that a dialog is presented indicating that the credentials were invalid.
 - [ ] Close the login prompt.
 - [ ] VERIFY that the "Log In" button is still available in the finbox.io ribbon.
 - [ ] VERIFY that the "Log Out" button is NOT available in the finbox.io ribbon.
 - [ ] VERIFY that no key is stored on the filesystem alongside the add-on.

###### Excel Add-In should warn on login attempts with unrecognized emails

 - [ ] Open Excel.
 - [ ] Ensure no user is logged in (if user is logged in, log out and restart excel).
 - [ ] Select "Log In" from finbox.io ribbon.
 - [ ] VERIFY that login form is displayed with email and password fields.
 - [ ] Enter login credentials with an incorrect password.
 - [ ] Click "Login" button.
 - [ ] VERIFY that a dialog is presented indicating that the credentials were invalid.
 - [ ] Close the login prompt.
 - [ ] VERIFY that the "Log In" button is still available in the finbox.io ribbon.
 - [ ] VERIFY that the "Log Out" button is NOT available in the finbox.io ribbon.
 - [ ] VERIFY that no key is stored on the filesystem alongside the add-on.

###### Excel Add-In should allow authenticated users to log out

 - [ ] Open Excel.
 - [ ] Ensure that a user is logged in (if no user is logged in, log in and restart excel).
 - [ ] Select "Log Out" from the finbox.io ribbon.
 - [ ] VERIFY that the "Log In" button is added to the finbox.io ribbon.
 - [ ] VERIFY that the "Log Out" button is removed from the finbox.io ribbon.
 - [ ] Add a FNBX(ticker, metric) formula to the workbook using a restricted company. Close any login prompts presented without logging in.
 - [ ] VERIFY that the formula returns an error.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a message indicates that the user was logged out.
 - [ ] VERIFY that a message indicates that the requested company is restricted to pro members.
 - [ ] VERIFY that no key is stored on the filesystem alongside the add-on.

###### Excel Add-In should persist login information across sessions

 - [ ] Open Excel.
 - [ ] Ensure no user is logged in (if user is logged in, log out and restart excel).
 - [ ] Select "Log In" from finbox.io ribbon.
 - [ ] VERIFY that the login form is displayed with email and password fields.
 - [ ] Enter valid premium login credentials.
 - [ ] Click "Login" button.
 - [ ] VERIFY that the login prompt is closed.
 - [ ] Restart Excel.
 - [ ] Open a new workbook.
 - [ ] VERIFY that the "Log In" button is removed from the finbox.io ribbon.
 - [ ] VERIFY that the "Log Out" button is added to the finbox.io ribbon.
 - [ ] Add a FNBX(ticker, metric) formula to the workbook using a restricted company.
 - [ ] VERIFY that the formula returns a value.

###### Excel Add-In should prompt for login on the first use of the FNBX formula

 - [ ] Open Excel.
 - [ ] Ensure no user is logged in (if user is logged in, log out and restart excel).
 - [ ] Open the 'tests/login.xlsx' workbook.
 - [ ] VERIFY that the login form is displayed with email and password fields.
 - [ ] Enter valid premium login credentials.
 - [ ] Press "Enter".
 - [ ] VERIFY that the login prompt is closed.
 - [ ] Restart Excel. Do not save changes to the workbook.
 - [ ] Open a new workbook.
 - [ ] VERIFY that the "Log In" button is removed from the finbox.io ribbon.
 - [ ] VERIFY that the "Log Out" button is added to the finbox.io ribbon.
 - [ ] Click the "Log Out" button.
 - [ ] Enter a FNBX formula into an empty cell.
 - [ ] VERIFY that the login form is displayed with email and password fields.

###### Excel Add-In should prevent anonymous accounts from accessing premium data

 - [ ] Open Excel.
 - [ ] Ensure no user is logged in (if user is logged in, log out and restart excel).
 - [ ] Add a FNBX(ticker, metric) formula to the workbook using a restricted company. Close any login prompts presented without logging in.
 - [ ] VERIFY that the formula returns an error.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a message indicates that the requested company is restricted to pro members.

###### Excel Add-In should prevent free accounts from accessing premium data

 - [ ] Open Excel.
 - [ ] Ensure no user is logged in (if user is logged in, log out and restart excel).
 - [ ] Click the "Log In" button in the finbox.io ribbon.
 - [ ] Enter valid non-premium login credentials.
 - [ ] Add a FNBX(ticker, metric) formula to the workbook using a restricted company.
 - [ ] VERIFY that the formula returns an error.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a message indicates that the requested company is restricted to pro members.


#### Upgrade

###### Excel Add-In should show the upgrade action to non-pro users

 - [ ] Open Excel.
 - [ ] Log in as a non-pro user.
 - [ ] VERIFY that the "Pro" button is available in the finbox.io ribbon.
 - [ ] Click the "Pro" button.
 - [ ] VERIFY that the user is directed to the upgrade page online.
 - [ ] Click the "Log Out" button.
 - [ ] VERIFY that the "Pro" button is NOT available in the finbox.io ribbon.

###### Excel Add-In should hide the upgrade action from pro users

 - [ ] Open Excel.
 - [ ] Log in as a pro user.
 - [ ] VERIFY that the "Pro" button is NOT available in the finbox.io ribbon.

###### Excel Add-In should hide the upgrade action from anonymous users

 - [ ] Open Excel.
 - [ ] Ensure user is logged out.
 - [ ] VERIFY that the "Pro" button is NOT available in the finbox.io ribbon.


#### Watchlist

###### Excel Add-In should provide easy access to the users watchlists

 - [ ] Open Excel.
 - [ ] VERIFY that the "Watchlist" button is available in the finbox.io ribbon.
 - [ ] Click the "Watchlist" button.
 - [ ] VERIFY that the browser is opened to the user's watchlist page.


#### Screener

###### Excel Add-In should provide easy access to the screener

 - [ ] Open Excel.
 - [ ] VERIFY that the "Screener" button is available in the finbox.io ribbon.
 - [ ] Click the "Screener" button.
 - [ ] VERIFY that the browser is opened to the screener page.


#### Templates

###### Excel Add-In should provide easy access to finbox.io template downloads

 - [ ] Open Excel.
 - [ ] VERIFY that the "Templates" button is available in the finbox.io ribbon.
 - [ ] Click the "Templates" button.
 - [ ] VERIFY that the browser is opened to the templates page.


#### Help

###### Excel Add-In should provide easy access to add-in help

 - [ ] Open Excel.
 - [ ] VERIFY that the "Help" button is available in the finbox.io ribbon.
 - [ ] Click the "Help" button.
 - [ ] Verify that the browser is opened to the Add-In Getting Started page.


#### About

###### Excel Add-In should provide information about the installed add-on version

 - [ ] Open Excel.
 - [ ] VERIFY that the "About" button is available in the finbox.io ribbon.
 - [ ] Click the "About" button in the finbox.io ribbon.
 - [ ] Verify that the add-in version is displayed in a dialog window.

###### Excel Add-In should provide information about the installed add-on location

 - [ ] Open Excel.
 - [ ] VERIFY that the "About" button is available in the finbox.io ribbon.
 - [ ] Click the "About" button in the finbox.io ribbon.
 - [ ] Verify that the add-in location is displayed in a dialog window.


#### Relink

###### Excel Add-In should automatically relink external FNBX formulas to the local add-in

 - [ ] Open Excel.
 - [ ] Open the 'tests/relink.xlsx' workbook.
 - [ ] VERIFY that a prompt is presented about updating workbook links.
 - [ ] Select the option to ignore the links.
 - [ ] VERIFY that all external FNBX formula links are replaced by local FNBX formulas.
 - [ ] Close the workbook and DO NOT SAVE CHANGES.


#### Refresh

###### Excel Add-In should allow users to force-update data in a workbook

 - [ ] Open Excel.
 - [ ] Open the 'tests/refresh1.xlsx' workbook.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a message indicates that 1 key was requested.
 - [ ] Close the message log.
 - [ ] Click the "Refresh" button in the finbox.io ribbon.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a second message indicates that 1 key was requested.

###### Excel Add-In refresh should only reload the current workbook

 - [ ] Open Excel.
 - [ ] Open the 'tests/refresh1.xlsx' workbook.
 - [ ] Open the 'tests/refresh2.xlsx' workbook.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a message indicates that 1 key was requested.
 - [ ] VERIFY that a message indicates that 2 keys were requested.
 - [ ] Close the message log.
 - [ ] Select the 'tests/refresh1.xlsx' workbook.
 - [ ] Click the "Refresh" button in the finbox.io ribbon.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a second message indicates that 1 key were requested.
 - [ ] VERIFY that still only one message indicates that 2 keys were requested.
 - [ ] Close the message log.
 - [ ] Select the 'tests/refresh2.xlsx' workbook.
 - [ ] Click the "Refresh" button in the finbox.io ribbon.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a second message indicates that 2 keys were requested.

###### Excel Add-In refresh should not clear the cache of other open workbooks

 - [ ] Open Excel.
 - [ ] Open the 'tests/refresh1.xlsx' workbook.
 - [ ] Open the 'tests/refresh2.xlsx' workbook.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a message indicates that 1 key was requested.
 - [ ] VERIFY that a message indicates that 2 keys were requested.
 - [ ] Close the message log.
 - [ ] Select the 'tests/refresh1.xlsx' workbook.
 - [ ] Click the "Refresh" button in the finbox.io ribbon.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that a second message indicates that 1 key was requested.
 - [ ] Close the message log.
 - [ ] Select the 'tests/refresh2.xlsx' workbook.
 - [ ] Copy and paste one FNBX cell into another empty cell.
 - [ ] Click the "Message Log" button in the finbox.io ribbon.
 - [ ] VERIFY that no additional messages about requesting keys were added.


#### Unlink

###### Excel Add-In should support unlinking FNBX formulas from workbooks

 - [ ] Open Excel.
 - [ ] Open the 'tests/unlink.xlsx' workbook.
 - [ ] Use 'Save As' to save a copy of the workbook somewhere on your system.
 - [ ] Click the "Unlink" button in the finbox.io ribbon.
 - [ ] VERIFY that the user is prompted to continue creating an unlinked version.
 - [ ] Click through the prompt to continue.
 - [ ] VERIFY that the user is prompted to specify a filename for the unlinked workbook.
 - [ ] Save the unlinked workbook somewhere on your system.
 - [ ] Search the workbook for any formula references to 'FNBX'.
 - [ ] VERIFY that no FNBX references are found.

###### Excel Add-In should force user to save before unlinking workbook

 - [ ] Open Excel.
 - [ ] Create a new workbook.
 - [ ] Add a FNBX formula to the workbook.
 - [ ] Click the "Unlink" button in the finbox.io ribbon.
 - [ ] VERIFY that a prompt is displayed indicating that the user must first save the workbook.
 - [ ] Click "OK" to acknowledge the prompt.
 - [ ] VERIFY that the FNBX formula has not been modified.
 - [ ] Save the workbook somewhere on your system.
 - [ ] Click the "Unlink" button in the finbox.io ribbon.
 - [ ] VERIFY that the user is prompted to continue creating an unlinked version.
 - [ ] Click through the prompt to continue.
 - [ ] VERIFY that the user is prompted to specify a filename for the unlinked workbook.
 - [ ] Save the unlinked workbook somewhere on your system.
 - [ ] Search the workbook for any formula references to 'FNBX'.
 - [ ] VERIFY that no FNBX references are found.
 - [ ] Add another FNBX formula to the open workbook.
 - [ ] Click the "Unlink" button in the finbox.io ribbon.
 - [ ] VERIFY that a prompt is displayed indicating that the user must first save the workbook.
 - [ ] Click "OK" to acknowledge the prompt.
 - [ ] VERIFY that the FNBX formula has not been modified.


#### Update

Unless otherwise specified, the following tests assume you are using an
unreleased workbook version, and is newer than the latest released
version available on finbox.io.

###### Excel Add-In should silently automatically check for updates on the first use of the FNBX formula when no updates are available

 - [ ] Open Excel.
 - [ ] Create a new workbook.
 - [ ] Click the "Message Log" button.
 - [ ] VERIFY that no message exists about checking for updates.
 - [ ] Close the message log.
 - [ ] Enter a FNBX formula into an empty cell. Log in if necessary.
 - [ ] VERIFY that no prompt about updates is displayed.
 - [ ] Click the "Message Log" button.
 - [ ] VERIFY that a message exists indicating that no updates are available.

###### Excel Add-In should automatically check for updates only once per session

 - [ ] Open Excel.
 - [ ] Create a new workbook.
 - [ ] Enter a FNBX formula into an empty cell. Log in if necessary.
 - [ ] Click the "Message Log" button.
 - [ ] VERIFY that a message exists indicating that no updates are available.
 - [ ] Close the message log.
 - [ ] Enter another FNBX formula into an empty cell
 - [ ] Click the "Message Log" button.
 - [ ] VERIFY that still only one message exists indicating that no updates are available.

###### Excel Add-In should indicate that no updates are available when manually checked and no updates are available

 - [ ] Open Excel.
 - [ ] Click the "Check for Updates" button in the finbox.io ribbon.
 - [ ] VERIFY that a prompt is displayed indicating that no updates are available.

###### Excel Add-In should automatically notify user when updates are available on the first use of the FNBX formula

 - [ ] Install an older version of the add-on.
 - [ ] Restart excel.
 - [ ] Create a new workbook.
 - [ ] Enter a FNBX formula into an empty cell.
 - [ ] VERIFY that the user is alerted to the availability of an update and given the option to download.
 - [ ] Choose not to download the update.
 - [ ] VERIFY that the formula resolves properly. Log in if necessary.

###### Excel Add-In should automatically guide users through the update process on the first use of the FNBX formula when updates are available

 - [ ] Install an older version of the add-on.
 - [ ] Restart excel.
 - [ ] Create a new workbook.
 - [ ] Enter a FNBX formula into an empty cell.
 - [ ] VERIFY that the user is alerted to the availability of an update and given the option to download.
 - [ ] Choose the option to download the latest installer.
 - [ ] VERIFY that the user is prompted to choose a location to save the latest installer.
 - [ ] Enter a location to save the add-on installer.
 - [ ] VERIFY that a prompt is displayed indicating that the installer has been saved and that the user must unblock and open it to complete the update process.
 - [ ] VERIFY that the formula resolves properly. Log in if necessary.
 - [ ] VERIFY that the add-on installer is indeed saved in the chosen location.
 - [ ] Open the new add-on installer and proceed with the installation.
 - [ ] Click the "About" button in the finbox.io ribbon.
 - [ ] VERIFY that the 'About' dialog indicates that the latest version is installed.
 - [ ] Restart Excel.
 - [ ] Click the "Check for Updates" button.
 - [ ] VERIFY that no further updates are available.

###### Excel Add-In should guide users through the update process when manually checked and updates are available

 - [ ] Install an older version of the add-on.
 - [ ] Restart excel.
 - [ ] Click the "Check for Updates" button in the finbox.io ribbon.
 - [ ] VERIFY that the user is alerted to the availability of an update and given the option to download.
 - [ ] Choose the option to download the latest installer.
 - [ ] VERIFY that the user is prompted to choose a location to save the latest installer.
 - [ ] Enter a location to save the add-on installer.
 - [ ] VERIFY that a prompt is displayed indicating that the installer has been saved and that the user must unblock and open it to complete the update process.
 - [ ] VERIFY that the add-on installer is indeed saved in the chosen location.
 - [ ] Open the new add-on installer and proceed with the installation.
 - [ ] Click the "About" button in the finbox.io ribbon.
 - [ ] VERIFY that the 'About' dialog indicates that the latest version is installed.
 - [ ] Restart Excel.
 - [ ] Click the "Check for Updates" button.
 - [ ] VERIFY that no further updates are available.


#### FNBX Formula

###### FNBX formulas should function properly in Excel

 - [ ] Open the 'tests/fnbx.xlsm' workbook.
 - [ ] VERIFY that all tests in the workbook are passing.
 - [ ] Click the "Refresh" button in the finbox.io ribbon.
 - [ ] VERIFY that all tests are still passing.

###### FNBX formulas should not be affected by opening another workbook

 - [ ] Open the 'tests/fnbx.xlsm' workbook.
 - [ ] VERIFY that all tests in the workbook are passing.
 - [ ] Create a new workbook.
 - [ ] VERIFY that all tests in 'fnbx.xlsm' are still passing.


#### Batching

###### FNBX formulas should be batched to minimize API requests

 - [ ] Open the 'tests/batch.xlsx' workbook.
 - [ ] Click the 'Message Log' button in the finbox.io ribbon.
 - [ ] VERIFY that 3 requests were made, for 126, 30, and 3 keys respectively.
 - [ ] Close the message log.
 - [ ] Select a different sheet in the workbook.
 - [ ] Click the 'Refresh' button in the finbox.io ribbon.
 - [ ] Click the 'Message Log' button in the finbox.io ribbon.
 - [ ] VERIFY that 3 additional requests were made, for 126, 30, and 3 keys respectively.

###### Request batches should include only keys from the current workbook

 - [ ] Open the 'tests/refresh1.xlsx' workbook.
 - [ ] Open the 'tests/refresh2.xlsx' workbook.
 - [ ] Click the 'Message Log' button in the finbox.io ribbon.
 - [ ] VERIFY that there is one recent message indicating that 1 key was requested.
 - [ ] VERIFY that there is one recent message indicating that 2 keys were requested, after the previous message.
 - [ ] Close the message log.
 - [ ] From the 'refresh1.xlsx' workbook, click the 'Refresh' button in the finbox.io ribbon.
 - [ ] Click the 'Message Log' button in the finbox.io ribbon.
 - [ ] VERIFY that there is one more message indicating that 1 key was requested.
 - [ ] Close the message log.
 - [ ] From the 'refresh2.xlsx' workbook, click the 'Refresh' button in the finbox.io ribbon.
 - [ ] Click the 'Message Log' button in the finbox.io ribbon.
 - [ ] VERIFY that there is one more message indicating that 2 keys were requested.


#### Caching

###### Excel Add-In should cache FNBX values

 - [ ] Open the 'tests/cache.xlsx' workbook.
 - [ ] Click the 'Message Log' button in the finbox.io ribbon.
 - [ ] VERIFY that only one message exists indicating that any keys were requested.
 - [ ] VERIFY that this message indicates that 5 keys were requested.
 - [ ] Close the message log.
 - [ ] Copy all used cells in the workbook and paste into an empty sheet of the same workbook.
 - [ ] VERIFY that the pasted values match the original.
 - [ ] Click the 'Message Log' button in the finbox.io ribbon.
 - [ ] VERIFY that still only one message exists indicating that any keys were requested.

###### Excel Add-In should share cached FNBX values across all workbooks

 - [ ] Open the 'tests/cache.xlsx' workbook.
 - [ ] Click the 'Message Log' button in the finbox.io ribbon.
 - [ ] VERIFY that only one message exists indicating that any keys were requested.
 - [ ] VERIFY that this message indicates that 5 keys were requested.
 - [ ] Close the message log.
 - [ ] Copy all used cells in the workbook and paste into an empty sheet of a new workbook.
 - [ ] VERIFY that the pasted values match the original.
 - [ ] Click the 'Message Log' button in the finbox.io ribbon.
 - [ ] VERIFY that still only one message exists indicating that any keys were requested.


#### International Support

For the following tests, you must temporarily change your system language
to one with different decimal and list separators (e.g. Slovak). Restart the
system after changing the language and before beginning these tests.

###### All supported FNBX usage should work with non-English internationalization settings

 - [ ] Open the 'tests/fnbx.xlsm' workbook.
 - [ ] VERIFY that all tests in the workbook are passing.
 - [ ] Click the "Refresh" button in the finbox.io ribbon.
 - [ ] VERIFY that all tests are still passing.

###### Complex batch requests should work with non-English internationalization settings

 - [ ] Open the 'tests/batch-combined.xlsx' workbook.
 - [ ] Click the 'Message Log' button in the finbox.io ribbon.
 - [ ] VERIFY that only one message exists indicating that any keys were requested.
 - [ ] VERIFY that this message indicates that 4 keys were requested.


#### Quota Usage

###### Excel Add-In should warn users when they surpass their data quota

  - [ ] Open a new workbook in Excel.
  - [ ] Enter the formula `=FNBX("x-mock-status", 429)` in an empty cell.
  - [ ] VERIFY that a dialog is presented indicating that the quota limit has been reached.

###### FNBX formula should return #N/A! errors for data requested after quota has been reached
  - [ ] Open a new workbook in Excel.
  - [ ] Enter a FNBX formula with valid ticker/metric arguments in an empty cell.
  - [ ] Enter the formula `=FNBX("x-mock-status", 429)` in an empty cell.
  - [ ] Click through any limit dialogs presented.
  - [ ] Click the 'Refresh' button in the finbox.io ribbon.
  - [ ] VERIFY that all FNBX cells display #N/A errors.

###### Excel Add-In should only warn users about the quota limit once every 5 minutes

  - [ ] Open a new workbook in Excel.
  - [ ] Enter the formula `=FNBX("x-mock-status", 429)` in an empty cell.
  - [ ] Click through any limit dialogs presented.
  - [ ] Enter a new FNBX formula in the workbook within 5 minutes.
  - [ ] VERIFY that no new dialogs are presented upon entry of the second formula.
  - [ ] Wait for 5 minutes.
  - [ ] Enter another FNBX formula in the workbook.
  - [ ] VERIFY that another limit dialog is presented.

###### Excel Add-In should explicitly warn users about the quota limit if they click 'Refresh' while blocked

  - [ ] Open a new workbook in Excel.
  - [ ] Enter the formula `=FNBX("x-mock-status", 429)` in an empty cell.
  - [ ] Click through any limit dialogs presented.
  - [ ] Click the 'Refresh' button in the finbox.io dialog.
  - [ ] VERIFY that another dialog is presented indicating that the quota limit has been reached.

###### Excel Add-In should temporarily block requests after the quota limit has been reached

  - [ ] Open a new workbook in Excel.
  - [ ] Enter the formula `=FNBX("x-mock-status", 429)` in an empty cell.
  - [ ] Click through any limit dialogs presented.
  - [ ] Enter a new FNBX formula in the workbook with real ticker/metric arguments.
  - [ ] Click the 'Message Log' button in the finbox.io ribbon.
  - [ ] VERIFY that no message exists indicating that keys were requested.

###### Excel Add-In should unblock requests when a user logs in

  - [ ] Open a new workbook in Excel.
  - [ ] Enter the formula `=FNBX("x-mock-status", 429)` in an empty cell.
  - [ ] Click through any limit dialogs presented.
  - [ ] Replace the FNBX formula arguments to request a valid ticker/metric.
  - [ ] Click the 'Log Out' button in the finbox.io ribbon.
  - [ ] Click the 'Log In' button in the finbox.io ribbon.
  - [ ] Enter a valid email/password and click 'Login'.
  - [ ] Verify that the FNBX formula is correctly loaded.

###### Excel Add-In should unblock requests after waiting 5 minutes since last 429 error

  - [ ] Open a new workbook in Excel.
  - [ ] Enter the formula `=FNBX("x-mock-status", 429)` in an empty cell.
  - [ ] Click through any limit dialogs presented.
  - [ ] Wait 5 minutes.
  - [ ] Replace the FNBX formula arguments to request a valid ticker/metric.
  - [ ] Verify that the FNBX formula is correctly loaded.

###### Excel Add-In should indicate quota usage

  - [ ] Open the 'tests/quota.xlsx' workbook.
  - [ ] VERIFY that the user's quota usage is shown in the finbox.io ribbon.
  - [ ] Click the 'Refresh' button in the finbox.io ribbon.
  - [ ] VERIFY that the quota usage is updated after the data is refreshed.
  - [ ] Click the 'Quota Usage' button in the finbox.io ribbon.
  - [ ] VERIFY that a dialog is presented indicating the user's quota usage.
