
1.6.0-beta.13 / 2019-10-23
==========================

  * Fix circular dependency and json conversion

1.6.0-beta.12 / 2019-10-23
==========================

  * Final install check

1.6.0-beta.11 / 2019-10-23
==========================

  * Please work

1.6.0-beta.10 / 2019-10-23
==========================

  * Debug AddIn.Add

1.6.0-beta.9 / 2019-10-23
=========================

  * Fix upgrade sequence

1.6.0-beta.8 / 2019-10-23
=========================

  * Fix add-in cleanup

1.6.0-beta.7 / 2019-10-23
=========================

  * Fix staged updates

1.6.0-beta.6 / 2019-10-23
=========================

  * Fix staged updates

1.6.0-beta.5 / 2019-10-23
=========================

  * Check legacy installer name

1.6.0-beta.4 / 2019-10-23
=========================

  * Debug upgrade

1.6.0-beta.3 / 2019-10-23
=========================

  * Add legacy version sheets

1.6.0-beta.2 / 2019-10-22
=========================

  * Debug upgrade

1.6.0-beta.1 / 2019-10-22
=========================

  * Update Mac form image clip

1.6.0-beta.0 / 2019-10-22
=========================

  * JSON parse improvements and rename finboxio to finbox
  * Update branding & Finbox urls, fix utf encoding issue

1.5.0 / 2019-07-05
==================

  * Remove S&P toggle
  * Patch Mac installation location


1.4.2-beta.0 / 2019-07-05
=========================

  * Remove S&P toggle and patch Mac installation location

1.4.1 / 2019-06-03
==================

  * Release 1.4.1

1.4.1-beta.0 / 2019-06-03
=========================

  * Reverse logic for determining if system libc needs to be used

1.4.0 / 2019-05-09
==================

  * Change default api version to v3 data
  * Update to support 16.16.9 on OSX

1.3.2-beta.2 / 2019-05-09
=========================

  * Update button text for reverting to legacy api

1.3.2-beta.1 / 2019-05-09
=========================

  * Update button text for reverting to legacy api

1.3.2-beta.0 / 2019-05-09
=========================

  * Change default api version to v3 data
  * Update to support 16.16.9 on OSX

1.3.1 / 2019-03-14
==================

  * Fix clib support for new release 16.16.8

1.3.0 / 2019-02-18
==================

  * Fix for compatibility with Excel for Mac v16.16.7
  * Update beta icons for mac compatibility
  * Add option to toggle beta api

1.3.0-beta.1 / 2019-02-15
=========================

  * Fix for compatibility with Excel for Mac v16.16.7

1.3.0-beta.0 / 2019-02-13
=========================

  * Update beta icons for mac compatibility

1.2.1-beta.0 / 2019-02-13
=========================

  * Add option to toggle beta api

1.2.0 / 2019-01-19
==================

  * Fix an integer overflow error for users with increased quota
  * Make null value configurable and default to 0
  * Send structured client header with system info in X-Finboxio-Addon
  * Adds configuration option for finboxioApiUrl
  * Update for compatibility with Excel for Mac v16.21+
  * Fix Mac installation bug

1.1.7-beta.3 / 2019-01-19
=========================

  * Include Mac v16.21 in the compatibility fix

1.1.7-beta.2 / 2019-01-19
=========================

  * Fix integer overflow error with increased quota

1.1.7-beta.1 / 2019-01-18
=========================

  * Add workbook name to API header
  * Make default null value configurable
  * Send structured client header in X-Finboxio-Addin
  * Make api url configurable

1.1.7-beta.0 / 2019-01-18
=========================

  * Update for compatibility with Excel for Mac v16.22
  * Fix Mac installation bug

1.1.6 / 2018-11-17
==================

  * Fixes bug with Mac update that broke VBA.MkDir

1.1.5 / 2018-11-01
==================

  * Update README

1.1.5-beta.1 / 2018-11-01
=========================

  * 1.1.5-beta.0
  * Try with beta prefix

1.1.5-beta.0 / 2018-11-01
=========================

  * Try with beta prefix

1.1.5-4 / 2018-11-01
====================

  * Fix release-notes one more time

1.1.5-3 / 2018-11-01
====================

  * Fix release notes script

1.1.5-2 / 2018-11-01
====================

  * Fix release-notes script

1.1.5-1 / 2018-11-01
====================

  * Add readme with release instructions

1.1.5-0 / 2018-11-01
====================

  * reset version

1.1.5-beta.0 / 2018-11-01
=========================

  * reset version

1.1.4-beta0.0 / 2018-10-29
==========================

  * Prerelease 1.1.4-beta0
  * Merge pull request #9 from finboxio/fix/vba-install-bug
  * Fix installation bug with new version of VBA

1.1.4 / 2018-01-23
==================

  * Support Mac 2016 v16.*

1.1.4-2 / 2018-01-22
====================

  * Mac v16.* compatibility

1.1.4-1 / 2018-01-22
====================

  * safer publish script

1.1.4-0 / 2018-01-22
====================

  * safer publish script

1.1.3-0 / 2018-01-22
====================

  * Support Mac 2016 v16.*

1.1.2 / 2017-09-14
==================

  * Fixes layout of Mac login warning

1.1.1 / 2017-09-14
==================

  * Add warning to Mac2016 login prompt about missing cursor

1.1.0 / 2017-09-14
==================

  * Switches default parse method on windows to iterate to improve performance for large workbooks
  * Adds settings module to functions add-in
  * Adds settings module to functions add-in

1.0.0 / 2017-09-12
==================

  * Improved performance with key batching
  * Better support for nested and table-based formulas
  * Support for international date and number formats
  * Improved installation flow
  * Fully automatic upgrades
  * Login persistence
  * Real-time quota indicator
  * Many other improvements related to stability and maintainability

0.24.5 / 2017-09-12
===================

  * Completed install/upgrade tests for Win2013
  * Fixes link replacement for links with unquoted filenames
  * Completed install/upgrade tests for Win2007 & Win2010
  * Completed install/upgrade tests for Mac2016
  * Finished Mac2011 install/upgrade tests

0.24.4 / 2017-09-12
===================

  * Fix installation cancel on Mac2011

0.24.3 / 2017-09-12
===================

  * Removes debug statement from uninstall procedure

0.24.2 / 2017-09-12
===================

  * Fixes uninstall on Mac2011

0.24.1 / 2017-09-12
===================

  * Improvements to FNBX link-replacer

0.24.0 / 2017-09-12
===================

  * Minor version bump

0.23.8 / 2017-09-12
===================

  * Adds FNBX help context for Windows users

0.23.7 / 2017-09-12
===================

  * Reload finboxio menu completely when opening add-in functions

0.23.6 / 2017-09-12
===================

  * Adds cleanup logic for older Mac2016 add-in versions installed in the wrong place

0.23.5 / 2017-09-11
===================

  * Ensures finboxio.xlam is fully removed before installing
  * Closes install dialog on Mac2016

0.23.4 / 2017-09-11
===================

  * Promote prerelease

0.23.4-0 / 2017-09-11
=====================

  * Fixes download module for Excel 2007

0.23.3 / 2017-09-11
===================

  * Updates test procedures for installation/update

0.23.3-0 / 2017-09-11
=====================

  * Moves installer log messages to default add-in location

0.23.2 / 2017-09-11
===================

  * Fix log access on Mac2011

0.23.1 / 2017-09-11
===================

  * Use native system app to open log file

0.23.0 / 2017-09-11
===================

  * Adds persistent, cross-component log file

0.22.0 / 2017-09-10
===================

  * Minor version bump

0.21.0 / 2017-09-10
===================

  * Minor version bump

0.21.0-49 / 2017-09-10
======================

  * Prevents install of local functions component outside dev

0.21.0-48 / 2017-09-10
======================

  * User-friendly install prompt
  * Installation fixes for Mac
  * Uninstall functions

0.21.0-47 / 2017-09-09
======================

  * Fix Mac2011 bug loading functions add-in
  * Adds  andler to close installer workbook to fix self-closing bug on Mac

0.21.0-46 / 2017-09-08
======================

  * Fixes repeated macro prompt on Mac 2011
  * Fixes missing up-to-date prompt on manual upgrade check

0.21.0-45 / 2017-09-08
======================

  * Fix excel version determination on Mac 2011

0.21.0-44 / 2017-09-08
======================

  * Fixes downloader path bug in Mac 2011

0.21.0-43 / 2017-09-08
======================

  * Wrap dir access in function for Mac 2011 safety

0.21.0-42 / 2017-09-08
======================

  * Adds more event handlers for promoting staged updates
  * Improves upgrade messaging on Mac

0.21.0-41 / 2017-09-08
======================

  * Fixes update bug when add-in manager is not open

0.21.0-40 / 2017-09-08
======================

  * Prevent recursive calls to update staged manager

0.21.0-39 / 2017-09-07
======================

  * Running in circles again

0.21.0-38 / 2017-09-07
======================

  * Running in circles again

0.21.0-37 / 2017-09-07
======================

  * Fixes upgrade bug

0.21.0-36 / 2017-09-07
======================

  * Closes Add-In manager instead of uninstalling during upgrade

0.21.0-35 / 2017-09-07
======================

  * Clarifies error message when add-in components fail to load

0.21.0-34 / 2017-09-07
======================

  * Adds safety check to prevent unloading functions module while checking updates

0.21.0-33 / 2017-09-07
======================

  * Set AutomationSecurity to low before opening add-in workbooks

0.21.0-32 / 2017-09-07
======================

  * Adds message for auto-confirmed upgrades alerting user about possible upcoming macros prompt

0.21.0-31 / 2017-09-07
======================

  * Fixes for upgrade process

0.21.0-30 / 2017-08-29
======================

  * Improves safety of function component upgrade on Mac

0.21.0-29 / 2017-08-29
======================

  * Fix installer bug for Mac
  * Improves safety of function component upgrade on Mac
  * Improves safety of manager component upgrade on Mac

0.21.0-28 / 2017-08-28
======================

  * Fix version tag in upgrade confirmation dialog

0.21.0-27 / 2017-08-28
======================

  * Adds version tag to update confirmation message

0.21.0-26 / 2017-08-28
======================

  * Install fixes for Mac

0.21.0-25 / 2017-08-28
======================

  * Improves error handling and user feedback for update sequence

0.21.0-24 / 2017-08-28
======================

  * Fixes crash on update when no internet connection is available
  * Removes unnecessary debug statements

0.21.0-23 / 2017-08-28
======================

  * Fixes auto-update interval check

0.21.0-22 / 2017-08-28
======================

  * Fixes messaging in add-in installation prompt

0.21.0-21 / 2017-08-28
======================

  * Updates workbook binaries
  * Adds feedback to the manual update action
  * Adds support for prerelease and autoupdate configuration
  * Improves installation logic and error handling
  * Adds consistency to MsgBox parameters

0.21.0-20 / 2017-08-27
======================

  * Bump version for testing

0.21.0-19 / 2017-08-27
======================

  * Reduces interval for update check

0.21.0-18 / 2017-08-27
======================

  * Fixes accidental recursion bug in forced update

0.21.0-17 / 2017-08-27
======================

  * Adds ForceUpdate function for testing

0.21.0-16 / 2017-08-27
======================

  * Bump version for testing

0.21.0-15 / 2017-08-27
======================

  * Update history

0.21.0-14 / 2017-08-27
======================

  * Prevents functions add-in from promoting manager while being updated

0.21.0-13 / 2017-08-27
======================

 * Bump version for testing

0.21.0-12 / 2017-08-27
======================

  * Ensures function add-in is unloaded before applying update

0.21.0-11 / 2017-08-27
======================

  * Bump version for testing

0.21.0-10 / 2017-08-27
======================

  * Bump version for testing

0.21.0-9 / 2017-08-27
=====================

  * Prevents update check if staged updates already exist

0.21.0-8 / 2017-08-27
=====================

  * Adds event handlers for automatic update checks

0.21.0-7 / 2017-08-27
=====================

  * Renames WorkbookLinkReplacer to generic AppEventHandler

0.21.0-6 / 2017-08-27
=====================

  * Adds app event handler to automatically promote staged manager

0.21.0-5 / 2017-08-27
=====================

  * Promote manager updates from the functions add-in

0.21.0-4 / 2017-08-27
=====================

  * Fixes installation for multi-component add-in module

0.21.0-3 / 2017-08-26
=====================

  * Adds support for blocking quota check on load
  * Renames src files to finboxio.functions.xlam

0.21.0-2 / 2017-08-25
=====================

  * Adds safety checks and finboxio.functions.xlam asset to publish script

0.21.0-1 / 2017-08-25
=====================

  * Updates release script to check in finboxio.functions.xlam

0.21.0-0 / 2017-08-25
=====================

  * Adds initial implementation of true auto-update

0.20.0 / 2017-08-24
===================

  * Completed international tests on Mac v1.0
  * Improves fnbx.xlsm test workbook and macros
  * Improves matching algorithm for replacing remote FNBX links
  * Removes manual calculation property from workbook binary
  * Updates international tests in test template
  * Adds completed international tests for v1.0 windows versions
  * Adds check to make sure workbook is not in use before preparing release
  * Adds check to ensure workbook calculation property is set to auto before release
  * Commits updated workbook binary
  * Updates batch/fnbx tests for international environment and fixes some issues with test macros
  * Updates string conversion of list values to use the system list separator character
  * Fixes batch parsing in international environments
  * Fixes Unlink image icon selection in international environments
  * Namespaces use of VBA.IsDate function for consistency

0.19.1 / 2017-08-24
===================

  * Adds dynamic button icon for Unlink action
  * Completed Mac 2016 tests for v1.0
  * Fixes batching problem with table references that include [#This Row], and fixes Excel 2007/2010 ribbon.
  * Removes extended add-in name from dialog box titles
  * Completed tests for Win 2007, v1.0
  * Update Win-2010.md
  * Completed tests for Win 2013 - v1.0

0.19.0 / 2017-08-23
===================

  * Updates fnbx test workbook
  * Adds dynamic menu items on Mac 2011
  * Fixes link replacement on Mac 2011
  * Completed Mac 2011 tests for v1.0

0.18.11 / 2017-08-22
====================

  * Fixes extraneous (and erroneous) check for updates during upgrade process

0.18.10 / 2017-08-22
====================

  * Implements better solution for missing FNBX arguments that does not require requesting unnecessary keys

0.18.9 / 2017-08-21
===================

  * Drops upgrade tests for unsupported auto-download feature

0.18.8 / 2017-08-21
===================

  * Fixes upgrade check when using outdated version
  * Fixes user pro status logged on authentication
  * add test templates for v1.0

0.18.7 / 2017-08-21
===================

  * update release notes with install reference

0.18.6 / 2017-08-20
===================

  * Check for updates on github

0.18.5 / 2017-08-18
===================

  * Fixes release-notes in publish script

0.18.4 / 2017-08-18
===================

  * Updates release notes to include all commits since last release

0.18.3 / 2017-08-18
===================

  * update release script

0.18.2 / 2017-08-18
===================

  * try reverting race condition fix

0.18.1 / 2017-08-18
===================

  * remove relese script debugging

0.18.0 / 2017-08-18
===================

  * try fixing git index.lock race condition

0.17.3 / 2017-08-18
===================

  * more release script debugging
  * Removes debugging in release script
  * debug git lock issue
  * Adds workbook compression after updating release version
  * Adds publish script to package.json

0.17.2 / 2017-08-18
===================

  * Adds publish script for creating releases

0.17.1 / 2017-08-18
===================

  * Adds support for automatic add-in versioning

0.17.0 / 2017-08-18
==================

  * Merge pull request #1 from finboxio/tests
  * major updates
  * Documents test procedure for installation
  * copy code into new workbook, fix duplicate declaration
  * significant refactoring to improve error handling and maintainability
  * significant improvements to mac performance, login persistence, international support, link replacement, and code documentation
