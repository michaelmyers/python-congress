python-congress is simple Python function that returns an Excel .xlsx file of the U.S. House of Representatives Roll Call Votes  for a given year.

The function accesses the Roll Call XML data provided by the Clerk of the House through their website, it then parses the data and arranges it by Roll Call number and Representative, then dumps it to a Excel file (.xlsx).

Required for Use

    python-congress
    Python 2.7.2
    Excel or OpenOffice
    Active internet connection

Python Dependancies

    urllib2
    datetime
    BeautifulSoup
    openpyxl


FOr more information, see http://michaelmmyers.com/?p=267