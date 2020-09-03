# python_tuto_vba_version
a quick example of the previous python tuto using vba

THIS IS REALLY BASIC JUST TO SHOW FORM GUI INTERACTION, SQLITE CONNECTION AND EXCEL VBA SYNTAX

NO CHECKS OR DATA VALIDATION DONE, THIS IS ONLY AN EXAMPLE AND CAN BE EXTENDED WITH SOME MORE ADDITIONAL WORK

For this to work, you need to install the sqlite odbc driver, be careful to choose the 32bit version if you run office 32bits even on windows x64. Once this done, check you have that showing in the excel data connection odbc query available

Open this file, there will be a button and if you click it, it will open a form and you can save the data entered on this form in the sqlite table on the worksheet...

For version _1_0 Microsoft ActiveX data object library should be added in the reference menu for the sqlite connection to work
For version _1_10 no need for additional reference
