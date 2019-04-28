<a rel="license" href="http://creativecommons.org/licenses/by-nc/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by-nc/4.0/88x31.png" /></a><br />This work is licensed under a <a rel="license" href="http://creativecommons.org/licenses/by-nc/4.0/">Creative Commons Attribution-NonCommercial 4.0 International License</a>.

**Title:**      concurrent Checkouts

**Author:**     Henry Steele, Library Technology Services, Tufts University

**Date:**        June 2018

**Purpose:**

        Create a report of concurrent checkouts that occured on multiple

        copies of the same volume, based on an exporte Analytics report with

        the criteria below.   Note the required format.

        This script finds out how often during the given time periods

        that multiple copies of the same volume were out at the same time,

        and how often that all copies of the same volume were out at the same time.

        This report assumes the Tufts University rubric for multiple copies,

        that they will have the same MMS Id and call number, but different barcodes

        The report returns counts for when all copies of a title were out

        at the same time, but excludes these counts if there is only one copy of a title

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

        - use pip or another Python installation utility to install:

            + pandas (this also installs numpy)

            + tkFileDialog

            + xlwt

            + xlsxwriter

            + xlrd

                + you&#39;ll also need to intall xlrd for read\_exce in pandas to work

       - Python \&gt; 3.0

            + pandas (this also installs numpy)

            + tkinter

            + xlwt

            + xlsxwriter

** Output:**

        The script will output an Excel workbook of concurrent checkouts counts

        for each volume.

**Method:**

        Dataframe &quot;a&quot; is a parsed version of the input report from Analytics.

        It contains &#39;Title&#39;, &#39;MMS Id&#39;, &#39;Permanent Call Number, &#39;Barcode&#39;,

        &#39;Loan Datetime&#39;, &#39;Return Datetime&#39;

        This is used to compare loan periods for different items of the same volume

        The logic of the script is to load loan and return times into

        dataframe &quot;c&quot;, where each datetime in which either a loan or return occurred

        is a column in the dataframe, and each copy (barcode) of the same volume

        is a row.  In the cell for each row,column, the script records whether it

        was a loan or a return.

        Dataframe c is rearranged by column name, so that the column names (datetimes)

        for loans and returns are in order.

        With the columns arranged in this way, some loans will span multiple columns,

        i.e., another transacation&#39;s loan or return will have occured in the middle

        of the loan period of this transacation. This is the kind of event the script

        is looking for, because it means loan periods of different copies of the same

        volume overlapped.  To be able to analyze this situation, the script fills

        in the columns between &quot;loan&quot; and &quot;return&quot; for transacations that span multiple

        columns with &quot;on loan.&quot;

        The last step the script takes is to analyze each column (datetime)

        to see how many of the volume&#39;s copies were on loan at that time.

        The important question the output will answer how often all copies of a given volumes

        were out at a given time because that may indicate that the library

        needs to purchase more copies, depending on how often this happened
