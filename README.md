# high-usage-library-stats-alma-api
**Title:**      concurrent Checkouts

**Author:**     Henry Steele, Library Technology Services, Tufts University

**Date:**        June 2018

**Purpose:**

  Create a report of concurrent checkouts that occured on multiple copies of the same volume, based on an exporte Analytics report with

  the criteria below.   Note the required format.   This script finds out how often during the given time periods that multiple copies \
  
  of the same volume were out at the same time, and how often that all copies of the same volume were out at the same time.

  This report assumes the Tufts University rubric for multiple copies, that they will have the same MMS Id and call number, but 
  
  different barcodes The report returns counts for when all copies of a title were out at the same time, but excludes these counts if 
  
  there is only one copy of a title

**Requirements:**

- **--** You need to set up an Alma Analytics API Key.   Details at [https://developers.exlibrisgroup.com/alma/apis/docs/analytics/R0VUIC9hbG1hd3MvdjEvYW5hbHl0aWNzL3JlcG9ydHM=/](https://developers.exlibrisgroup.com/alma/apis/docs/analytics/R0VUIC9hbG1hd3MvdjEvYW5hbHl0aWNzL3JlcG9ydHM=/)

**Input:**

  The Analtyics report should have the following fields.  They can be in any

  order, and you can have additional fields (they&#39;ll be ignored) but the field names

  must be as below.  It should be in Excel format .xlsx format

- fulfilllment table with at least
  - Title
  - MMS Id
  - Permanent Call Number
  - Barcode
    - This is the item barcode, but leave the field name &quot;Barcode&quot;
  - Loan Date
  - Loan Time
  - Return Date
  - Return Time

**Dependencies:  **

  Note that this code is currently configured for Python 2.7, but I&#39;ve noted in

  the dependencies below and in various places in the code how to convert (refactor) this for Python \&gt; 3

    - Python 2.7

       - pandas (this also installs numpy)

       - tkFileDialog

       - xlwt

        - xlsxwriter

        - xlrd

    - Python \> 3.0

       - pandas (this also installs numpy)

       - tkinter

       - xlwt

       - xlsxwriter

** Output:**

   The script will output an Excel workbook of concurrent checkouts counts

   for each volume.

