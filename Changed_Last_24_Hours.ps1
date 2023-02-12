# Define the target directories to check for downloads
$targetDirectories = @(
    "$([Environment]::GetFolderPath('MyComputer'))\",
    "$([Environment]::GetFolderPath('Desktop'))\",
    "$([Environment]::GetFolderPath('InternetCache'))\",
    "$([Environment]::GetFolderPath('History'))\"
)

# Get all files in the target directories that were created in the last 24 hours
$last24hours = Get-ChildItem $targetDirectories -Recurse |
    Where-Object { $_.CreationTime -gt (Get-Date).AddDays(-1) }

# Create a new Excel workbook and worksheet
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Sheets.Item(1)

# Write headers for the columns
$worksheet.Cells.Item(1, 1) = "Full Path"
$worksheet.Cells.Item(1, 2) = "Filename"
$worksheet.Cells.Item(1, 3) = "Size (KB)"
$worksheet.Cells.Item(1, 4) = "Date Created"

# Loop through the list of files and write their properties to the worksheet
$row = 2  # Start at row 2 to skip the header row
foreach ($file in $last24hours) {
    $fullPath = $file.FullName
    $filename = $file.Name
    $sizeKB = [Math]::Round($file.Length / 1KB, 2)
    $dateCreated = $file.CreationTime

    $worksheet.Cells.Item($row, 1) = $fullPath
    $worksheet.Cells.Item($row, 2) = $filename
    $worksheet.Cells.Item($row, 3) = $sizeKB
    $worksheet.Cells.Item($row, 4) = $dateCreated

    $row++
}

# Autofit the columns and save the workbook
$range = $worksheet.UsedRange
$range.EntireColumn.AutoFit() | Out-Null
$desktopPath = [Environment]::GetFolderPath('Desktop')
$workbook.SaveAs("$desktopPath\Downloads - Last 24 Hours.xlsx")
$excel.Quit()