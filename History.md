
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
