## Test Procedure for finbox.io Excel Add-In

### Platforms To Test

* Excel 2016 - Windows (64-bit)
* Excel 2016 - Windows (32-bit)
* Excel 2013 - Windows (32-bit)
* Excel 2010 - Windows (32-bit)
* Excel 2007 - Windows (32-bit)
* Excel 2016 - Mac
* Excel 2011 - Mac

### Features To Test

#### Installation

##### Excel Add-In should succesfully install on a clean system

1. VERIFY that any previous add-in installation is completely removed.
2. Open Excel Add-In workbook.
3. Accept "Enable Macros" prompt.
4. VERIFY user is prompted to approve installation.
5. Accept installation prompt.
6. VERIFY that user is prompted to quit excel.
7. Accept prompt to quit excel.
8. Restart excel.
9. VERIFY that finbox.io ribbon is present.
10. VERIFY that finboxio.xlam is present in add-in folder.
11. VERIFY that FNBX formula is available.

##### Excel Add-In should successfully overwrite an existing installation

1. VERIFY that a previous add-in is installed (follow documented installation procedure).
2. Open Excel Add-In workbook.
3. Accept "Enable Macros" prompt.
4. VERIFY user is prompted to approve installation.
5. Accept installation prompt.
6. VERIFY that user is prompted to quit excel.
7. Accept prompt to quit excel.
8. Restart excel.
9. VERIFY that finbox.io ribbon is present.
10. VERIFY that finboxio.xlam is present in add-in folder.
11. VERIFY that FNBX formula is available.
12. VERIFY that older version was overwritten and only latest install is present.

##### Excel Add-In should not install if user declines installation prompt

1. Open Excel Add-In workbook.
2. Accept "Enable Macros" prompt.
3. VERIFY user is prompted to approve installation.
4. Decline installation prompt.
5. VERIFY that no more prompts are displayed.
6. VERIFY that finboxio.xlam is not present in add-in folder.
7. Restart excel.
8. VERIFY that finbox.io ribbon is not present after restart.

#### Authentication

##### Excel Add-In should allow authentication with email/password (non-Mac2016)

1. Open Excel.
2. Ensure no user is loggesd in. If user is logged in, log out and restart excel.
3. Select "Log In" from finbox.io menu.
4. VERIFY that login form is displayed with email and password fields.
5. Enter valid login credentials. 
6. VERIFY that password entry is obscured.
7. Click "Login" button.
8. VERIFY that prompt is closed.
9. 
