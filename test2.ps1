$excelPath = "C:\path\to\output.xlsx"
$yamlPath = "C:\path\to\reconstructed.yml"

# Start Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item(1)
$range = $sheet.UsedRange
$rowCount = $range.Rows.Count
$colCount = $range.Columns.Count

# Read headers
$headers = @()
for ($col = 1; $col -le $colCount; $col++) {
    $headers += $range.Cells.Item(1, $col).Text
}

# Read data row by row
$data = @()
$envOrder = @()
foreach ($row in 2..$rowCount) {
    $entry = @{}
    for ($col = 1; $col -le $colCount; $col++) {
        $key = $headers[$col - 1]
        $value = $range.Cells.Item($row, $col).Text
        $entry[$key] = $value
    }
    $data += $entry

    # Track env order
    $env = $entry["env"]
    if (-not $envOrder.Contains($env)) {
        $envOrder += $env
    }
}

# Close Excel
$workbook.Close($false)
$excel.Quit()

# Build YAML preserving env order
$yamlLines = @()
foreach ($env in $envOrder) {
    $yamlLines += "$env:"
    $tasks = $data | Where-Object { $_["env"] -eq $env }
    foreach ($item in $tasks) {
        $yamlLines += "  - taskName: $($item["taskName"])"
        $yamlLines += "    datasetName: $($item["datasetName"])"
        $yamlLines += "    taskCmd: $($item["taskCmd"])"
        $yamlLines += "    cronExpression: $($item["cronExpression"])"
    }
}

# Write to YAML file
$yamlLines | Set-Content -Path $yamlPath -Encoding UTF8
