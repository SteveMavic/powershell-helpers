$excel = New-Object -ComObject Excel.Application

# Specify Excel file name 
$fileName = "example"

# Set the path to the Excel file (relative to the location of the script)
$excelFile = Join-Path $PSScriptRoot "${fileName}.xlsx"

# Open the workbook
$workbook = $excel.Workbooks.Open($excelFile)

# Select 1st sheet
$worksheet = $workbook.Sheets.Item(1)
$data = @()

# Get the number of rows and columns in the worksheet
$rowCount = ($worksheet.UsedRange.Rows).Count
$columnCount = ($worksheet.UsedRange.Columns).Count

# Loop through each row and column, and add each cell's value to the data array (change 2 to 1 to don't skip header row)
for ($i = 2; $i -le $rowCount; $i++)
{
    $row = @{}
    for ($j = 1; $j -le $columnCount; $j++)
    {
        $cellValue = $worksheet.Cells.Item($i, $j).Value2
        $columnName = $worksheet.Cells.Item(1, $j).Value2
        # Get the header name for the current column and clean it up (replace spaces with underscores and make it lowercase)
        $columnName = ($columnName -replace "\s+", "_").ToLower()
        $row.Add($columnName, $cellValue)
    }
    $data += $row
}

# Convert the data array to JSON and output it to a file
$data | ConvertTo-Json -Depth 4 | Out-File -FilePath "${fileName}.json"

# Close the workbook and quit Excel
$workbook.Close()
$excel.Quit()