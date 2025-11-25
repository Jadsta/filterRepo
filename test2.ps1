# Path to the Excel file
$excelPath = "C:\path\to\output.xlsx"

# Start Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item(1)
$range = $sheet.UsedRange
$rowCount = $range.Rows.Count

# Read headers
$headers = @()
for ($col = 1; $col -le $range.Columns.Count; $col++) {
    $headers += $range.Cells.Item(1, $col).Text
}

# Read data into array of hashtables
$data = @()
for ($row = 2; $row -le $rowCount; $row++) {
    $entry = @{}
    for ($col = 1; $col -le $headers.Count; $col++) {
        $key = $headers[$col - 1]
        $value = $range.Cells.Item($row, $col).Text
        $entry[$key] = $value
    }
    $data += $entry
}

# Close Excel
$workbook.Close($false)
$excel.Quit()

# Group by environment
$grouped = $data | Group-Object -Property env

# Build YAML
$yamlLines = @()
foreach ($group in $grouped) {
    $yamlLines += "$($group.Name):"
    foreach ($item in $group.Group) {
        $yamlLines += "  - taskName: $($item.taskName)"
        $yamlLines += "    datasetName: $($item.datasetName)"
        $yamlLines += "    taskCmd: $($item.taskCmd)"
        $yamlLines += "    cronExpression: $($item.cronExpression)"
    }
}

# Output to file
$yamlPath = "C:\path\to\reconstructed.yml"
$yamlLines | Set-Content -Path $yamlPath -Encoding UTF8
