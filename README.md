# BulkCall2XLS
A simple python code to digest bulk ffiec data into a single sheet in xls.  Allowing for creation of charts and pivot tables that would otherwise be difficult to process.

To use this script, you must have python installed on your system.

Place the python script into a folder where your single period bulk call report data downloaded from https://cdr.ffiec.gov/public/PWS/DownloadBulkData.aspx

The name of the zip is also hardcoded right now, but Ill figure out how we want to handle this in the future.  You can always just adjust line 17 with the correct date to use a different zip file

Run script (can take as long as several minutes to process)

Open master.xlsx

You can create charts and pivot tables in anticipation of the data.  Use sheet 2 for this.  Ill udpate later to be able to fill data from multiple time periods and just add a sheet for this.
