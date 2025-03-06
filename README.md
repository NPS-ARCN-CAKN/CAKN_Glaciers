# CAKN Glaciers Data Management Scripts

Scripts related to the management of glaciers monitoring data in the National Parks of Alaska

Note: You will likely need to run the scripts as administrator.

# Documentation

This script translates the CAKN Glaciers Snow field data deliverable (Excel spreadsheet) records to an SQL script of insert queries suitable for execution in SQL Server Management Studio to insert the records into the CAKN_Glaciers SQL Server database.

**NOTE**: This script does not alter any data in any source file or the CAKN_Glaciers database; it may be run freely without concern of accidental data loss.

Written by NPS\SDMiller, 2024-12-23

# Instructions

1.  Navigate to the PowerShell file using Windows File Explorer

2.  Right click this file name and select 'Edit'. Ensure the script opens in Windows PowerShell ISE. To set PowerShell ISE as the default editor:

<!-- -->

a.  Right click the .ps1 file

b.  Select `Open with`

c.  Select `Choose another app`

d.  Select `Choose an app on your PC`

e.  Select ``` C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe``(not powershell.exe). ```

f.  You should see a split interface with code above and a PowerShell window below.

<!-- -->

3.  Edit the script as follows:

    a.  Change `$SourceFilename`, to your Excel file name.
    b.  Change `$SourceFileDirectory` to the directory where `$SourceFilename` resides.
    c.  Change `$WorksheetIndex` to the worksheet number, if necessary. Default is 1 (the first worksheet).

4.  Run the script in PowerShell by clicking the green triangle in the main toolstrip.

5.  If the script ran successfully there will be a new file with the same name as the source file with a '`.sql`' file extension.

6.  Open the `.sql` file in SQL Server Management Studio and execute it.

7.  If all the queries succeeded then execute `COMMIT` to complete the transaction and write the records to the database.

8.  If there were any errors execute `ROLLBACK` to cancel the transaction. Fix any errors and try again.

**YOU MUST COMMIT OR ROLLBACK**
