# Documentation: This script translates the CAKN Glaciers Stakes field data deliverable (Excel spreadsheet) records to an SQL script of
# insert queries suitable for execution in SQL Server Management Studio to insert the records into the database's Stakes table.

# NOTE: This script does not alter any data in any source file or the CAKN_Glaciers database; it may be run freely without concern of accidental data loss.

# Written by NPS\SDMiller, 2024-12-23

# Instructions: 
# 1. Change $SourceFilename, below' to your Excel file name.
# 2. Change $SourceFileDirectory to the directory where the file above resides.
# 3. Change $WorksheetIndex to the worksheet number. Default assumption is the Stakes worksheet is the first in the worksheets collection.
# 4. Run the script in PowerShell.
# 5. If the script ran successfully there will be a new file with the same name as the source file with a '.sql' file exension.
# 6. Open the .sql file in SSMS and execute it.
# 7. If all the queries succeeded without any errors then execute COMMIT to complete the transaction and write the records to the database.
# 8. If there were any errors or warnings execute ROLLBACK to cancel the transaction. Fix any errors and try again.

# User supplied parameters
$SourceFilename = "Kennicott stakes 2018 vFinal.xlsx"
$SourceFileDirectory = "C:\Temp\zGlaciers\"
$WorksheetIndex = 1 # Set this to the worksheet index (1 assumes Stakes is the first worksheet


# Concatenate the source file name with the path
$SourceFile = $SourceFileDirectory + $SourceFilename

# Create an Excel object and define it as $Workbook
$ExcelObj = New-Object -ComObject Excel.Application
$WorkBook = $ExcelObj.Workbooks.Open($SourceFile)

# Access the first worksheet
$Worksheet = $Workbook.Sheets.Item($WorksheetIndex)
#$Worksheet

# Get the used range of the worksheet
$UsedRange = $Worksheet.UsedRange
$RowCount = $UsedRange.Rows.Count
$ColumnCount = $UsedRange.Columns.Count


function NullIfBlank {
    param (
        [string]$Value,
        [bool]$IsNumeric
    )

    # Trim the input string
    $TrimmedValue = $Value.Trim()

    # Check if the trimmed string is blank
    if ([string]::IsNullOrWhiteSpace($TrimmedValue)) {
        # NULL
        return "NULL"
    } else {
        # Not NULL
        if($IsNumeric){
            return $TrimmedValue
        } else {
            return "'" + $TrimmedValue + "'"
        }
    }

}

# Example usage
#$Stake_Label = NullIfBlank -Value "   " -IsNumeric $true
#$Stake_Label
#$Stake_Label = NullIfBlank -Value "a stake name" -IsNumeric $false
#$Stake_Label
#$Stake_Label = NullIfBlank -Value "2.45" -IsNumeric $true
#$Stake_Label


# Start the SQL script
$Sql = "-- Script to insert glacier stake data into the CAKN_Glaciers.Stakes table.
"
$Sql = $Sql + "-- Source file: " +  $SourceFile + "
"
$Sql = $Sql + "-- " + $env:USERNAME + "
"
$CurrentDate = Get-Date
$Sql = $Sql + "-- " + $CurrentDate + " 
"
# Start the SQL script
$Sql = $Sql + "

-- INSTRUCTIONS: Execute the insert queries below. If all the insertions completed without error then execute COMMIT, otherwise execute ROLLBACK, fix any errors and try again.
-- YOU MUST COMMIT OR ROLLBACK, or the database will be left in an unusable, hanging state.

"
$Sql = $Sql + "BEGIN TRANSACTION -- COMMIT ROLLBACK -- Make sure to COMMIT or ROLLBACK or the database will be left in a hanging state!
"




# Loop through each row in the worksheet and generate in INSERT query for each
for ($row = 2; $row -le $RowCount; $row++) {

$Glacier = NullIfBlank -Value $Worksheet.Cells.Item($row,1).Value() -IsNumeric $false
$Site = NullIfBlank -Value $Worksheet.Cells.Item($row,2).Value() -IsNumeric $false
$Stake_Label = NullIfBlank -Value $Worksheet.Cells.Item($row,3).Value() -IsNumeric $false
$Stake_Name = NullIfBlank -Value $Worksheet.Cells.Item($row,4).Value() -IsNumeric $false
$Date_Time = NullIfBlank -Value $Worksheet.Cells.Item($row,5).Value() -IsNumeric $false
$Latitude = NullIfBlank -Value $Worksheet.Cells.Item($row,6).Value() -IsNumeric $true
$Longitude = NullIfBlank -Value $Worksheet.Cells.Item($row,7).Value() -IsNumeric $true
$HAMSL_m = NullIfBlank -Value $Worksheet.Cells.Item($row,8).Value() -IsNumeric $true
$Coordinates_Note = NullIfBlank -Value $Worksheet.Cells.Item($row,9).Value() -IsNumeric $false
$Found_or_Left = NullIfBlank -Value $Worksheet.Cells.Item($row,10).Value() -IsNumeric $false
$Stake_Length_m = NullIfBlank -Value $Worksheet.Cells.Item($row,11).Value() -IsNumeric $true
$Stake_Exposed_m = NullIfBlank -Value $Worksheet.Cells.Item($row,12).Value() -IsNumeric $true
$Stake_Condition_note = NullIfBlank -Value $Worksheet.Cells.Item($row,13).Value() -IsNumeric $false
$New_or_Existing = NullIfBlank -Value $Worksheet.Cells.Item($row,14).Value() -IsNumeric $false
$Summer_Lowering_m = NullIfBlank -Value $Worksheet.Cells.Item($row,15).Value() -IsNumeric $true
$Lowering_note = NullIfBlank -Value $Worksheet.Cells.Item($row,16).Value() -IsNumeric $false
$Winter_Ablation_SWE_m = NullIfBlank -Value $Worksheet.Cells.Item($row,17).Value() -IsNumeric $true
$Winter_Ablation_Note = NullIfBlank -Value $Worksheet.Cells.Item($row,18).Value() -IsNumeric $false
$Surface_Type = NullIfBlank -Value $Worksheet.Cells.Item($row,19).Value() -IsNumeric $false
$Surface_Below_Seasonal_Snow = NullIfBlank -Value $Worksheet.Cells.Item($row,20).Value() -IsNumeric $false
$Total_Snow_Depth_m = NullIfBlank -Value $Worksheet.Cells.Item($row,21).Value() -IsNumeric $true
$Summer_Accum_m = NullIfBlank -Value $Worksheet.Cells.Item($row,22).Value() -IsNumeric $true
$Seasonal_Snow_Depth_m = NullIfBlank -Value $Worksheet.Cells.Item($row,23).Value() -IsNumeric $true
$Snow_Depth_note = NullIfBlank -Value $Worksheet.Cells.Item($row,24).Value() -IsNumeric $false
$Seasonal_Snow_SWE_m = NullIfBlank -Value $Worksheet.Cells.Item($row,25).Value() -IsNumeric $true
$Snow_SWE_note = NullIfBlank -Value $Worksheet.Cells.Item($row,26).Value() -IsNumeric $false
$Melt_Season_SWE_Change_m = NullIfBlank -Value $Worksheet.Cells.Item($row,27).Value() -IsNumeric $true
$Summer_SWE_Change_note = NullIfBlank -Value $Worksheet.Cells.Item($row,28).Value() -IsNumeric $true
$Annual_Balance_SWE_m = NullIfBlank -Value $Worksheet.Cells.Item($row,29).Value() -IsNumeric $true
$Other_Notes = NullIfBlank -Value $Worksheet.Cells.Item($row,30).Value() -IsNumeric $false


    # Check if $Glacier is not null and not equal to 'Glacier'
    if ($null -ne $Glacier -and $Glacier -ne 'Glacier') {

    
            $Sql = $Sql + "INSERT INTO [dbo].[Stakes]
([Glacier]
,[Site]
,[Stake_Label]
,[Stake_Name]
,[Date_Time]
,[Latitude]
,[Longitude]
,[HAMSL_m]
,[Coordinates_Note]
,[Found_or_Left]
,[Stake_Length_m]
,[Stake_Exposed_m]
,[Stake_Condition_note]
,[New_or_Existing]
,[Summer_Lowering_m]
,[Lowering_note]
,[Winter_Ablation_SWE_m]
,[Winter_Ablation_Note]
,[Surface_Type]
,[Surface_Below_Seasonal_Snow]
,[Total_Snow_Depth_m]
,[Summer_Accum_m]
,[Seasonal_Snow_Depth_m]
,[Snow_Depth_note]
,[Seasonal_Snow_SWE_m]
,[Snow_SWE_note]
,[Melt_Season_SWE_Change_m]
,[Summer_SWE_Change_note]
,[Annual_Balance_SWE_m]
,[Other_Notes]
,[SourceFileName])
VALUES
(" + $Glacier + "
," + $Site + "
," + $Stake_Label + "
," + $Stake_Name + "
," + $Date_Time + "
," + $Latitude + "
," + $Longitude + "
," + $HAMSL_m + "
," + $Coordinates_Note + "
," + $Found_or_Left  + "
," + $Stake_Length_m + "
," + $Stake_Exposed_m + "
," + $Stake_Condition_note + "
," + $New_or_Existing + "
," + $Summer_Lowering_m + "
," + $Lowering_note + "
," + $Winter_Ablation_SWE_m  + "
," + $Winter_Ablation_Note + "
," + $Surface_Type  + "
," + $Surface_Below_Seasonal_Snow  + "
," + $Total_Snow_Depth_m  + "
," + $Summer_Accum_m  + "
," + $Seasonal_Snow_Depth_m  + "
," + $Snow_Depth_note  + "
," + $Seasonal_Snow_SWE_m  + "
," + $Snow_SWE_note  + "
," + $Melt_Season_SWE_Change_m  + "
," + $Summer_SWE_Change_note  + "
," + $Annual_Balance_SWE_m  + "
," + $Other_Notes  + "
,'" + $SourceFilename + "');

"

    } else {
        #Write-Output "The variable \$Glacier is either null or equal to 'Glacier'."
    }
    
}



# Dump out the SQL to a file with the same name as the input file but with a '.sql' extension.
$SqlFile = $SourceFile + ".sql"
$Sql | Out-File -FilePath $SqlFile

# Close the workbook without saving
$Workbook.Close($false)

# Quit the Excel application
$ExcelObj.Quit()

