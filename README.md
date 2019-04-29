# high-usage-library-stats-alma-api

<a rel="license" href="http://creativecommons.org/licenses/by-nc/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by-nc/4.0/88x31.png" /></a><br />This work is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by-nc/4.0/">Creative Commons Attribution-NonCommercial 4.0 International License</a>.

**Title:**      concurrent Checkouts

**Author:**     Henry Steele, Library Technology Services, Tufts University

**Date:**        June 2018

**Purpose:**

Create a report of concurrent checkouts that occured on multiple copies of the same volume, based on an exporte Analytics report with the criteria below.   Note the required format below.  This script finds out how often during the given time periods that multiple copies of the same volume were out at the same time, and how often **all** copies of the same volume were out at the same time.  This report assumes the Tufts University structure for multiple copies, that they will have the same MMS Id and call number, but different barcodes.  The report returns counts for when all copies of a title were out at the same time, but excludes these counts if there is only one copy of a title.

Comments in the code describe in detail how the algorithm that gets these stats works.

**Input:**

The Analtyics report should have the following fields.  They can be in any order, and you can have additional fields (they&#39;ll be ignored) but the field names must be as below.  It should be in Excel format .xlsx format

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

**To Run:**
You can install the dependencies in the requirements file by running the following command before you run the script:
  - pip install -r requiremets

Then just call the python script, without arguments:
  - python concurrentCheckouts.py


**Dependencies:**

Note that this code is currently configured for Python 2.7, but I've noted in the dependencies below and in various places in the code how to convert (refactor) this for Python \> 3


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
      
      - xlrd

**Output:**

   The script will output an Excel workbook of concurrent checkouts counts for each volume.

