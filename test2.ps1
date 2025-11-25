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

# Read headers dynamically
$headers = @()
for ($col = 1; $col -le $colCount; $col++) {
    $headers += $range.Cells.Item(1, $col).Text
}

# Read data row by row
$data = @()
$envOrder = @()
for ($row = 2; $row -le $rowCount; $row++) {
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

# Build YAML dynamically
$yamlLines = @()
foreach ($env in $envOrder) {
    $yamlLines += "$env:"
    $tasks = $data | Where-Object { $_["env"] -eq $env }

    foreach ($task in $tasks) {
        $yamlLines += "  -"
        foreach ($key in $headers) {
            if ($key -ne "env") {
                $yamlLines += "    $key: $($task[$key])"
            }
        }
    }
}

# Write to YAML file
$yamlLines | Set-Content -Path $yamlPath -Encoding UTF8
