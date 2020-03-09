clear
fgfhfghfghfghfgh
$ExcelApp = New-Object -ComObject Excel.Application
$ExcelApp.Visible = $false
$Workbook = $ExcelApp.Workbooks.Open("C:\Users\selyuto\Desktop\Magik\16 каталог.xlsx")
$Worksheet = $Workbook.Worksheets.Item(1)
Function Delete-EmptyCode ($Worksheet, $ExcelApp)
{
    try {$Worksheet.ShowAllData()} catch {"ShowAllData already applied"}
    $LastRow = $Worksheet.Cells.Item($Worksheet.Rows.Count, "A").End(-4162).Row
    $Worksheet.Range("I2:I$($LastRow)").AutoFilter(9, "")
    $FilteredRange = $Worksheet.Range("A2:A$($LastRow)")
    $FilteredRange.Select()
    $ExcelApp.Selection.EntireRow.Delete()
    $Worksheet.ShowAllData()
}
Delete-EmptyCode -Worksheet $Worksheet -ExcelApp $ExcelApp

Function Delete-UnwantedColumns ($Worksheet, $ExcelApp)
{
    $UnwantedColumns = @("КатегорияКМ", "Осн.ШК", "Коммент Маркетинг", "Коммент КМ", "Кодпоставщика", "Поставщик", "IdИД Кат.-даты действия", "Название акции", "ТО план.,шт", "ТО план.,руб.", "Kpi14 Руб", "Kpi14 Шт")
    try {$Worksheet.ShowAllData()} catch {"ShowAllData already applied"}
    $LastColumn = $Worksheet.Cells(1, $Worksheet.Columns.Count).End(-4159).Column
    for ($i = $LastColumn; $i -ge 1; $i--) {
        if ($UnwantedColumns -contains $Worksheet.Rows.Item(1).Cells.Item($i).Value()) {
            $Worksheet.Columns.Item($i).Delete()
        }
    }
}
Delete-UnwantedColumns -Worksheet $Worksheet -ExcelApp $ExcelApp
fghfghfghfgh

Function Change-ColumnDataFormatToNumeric ($Worksheet, $ExcelApp)
{
    try {$Worksheet.ShowAllData()} catch {"ShowAllData already applied"}
    $LastColumn = $Worksheet.Cells(1, $Worksheet.Columns.Count).End(-4159).Column
    for ($i = $LastColumn; $i -ge 1; $i--) {
        if ($Worksheet.Rows.Item(1).Cells.Item($i).Value() -eq "Код товара") {
            $LastRow = $Worksheet.Cells.Item($Worksheet.Rows.Count, $i).End(-4162).Row
            $ColumnLetter = $Worksheet.Cells.Item(1, $i).Address() -replace "\$", "" -replace "\d", ""
            $Worksheet.Range("$($ColumnLetter)2:$($ColumnLetter)$($LastRow)").Select()
            $ExcelApp.Selection.NumberFormat = "0"
            $ExcelApp.Selection.Value2 = $ExcelApp.Selection.Value2
        }
    }
}
Change-ColumnDataFormatToNumeric -Worksheet $Worksheet -ExcelApp $ExcelApp

Function Truncate-ToTwoDecimalPlaces ($Worksheet, $ExcelApp, $Column)
{
fghfhgfghfghfgh
    try {$Worksheet.ShowAllData()} catch {"ShowAllData already applied"}
    $LastRow = $Worksheet.Cells.Item($Worksheet.Rows.Count, "A").End(-4162).Row
    $Worksheet.Range("$($Column)2:$($Column)$($LastRow)").Select()
    $ExcelApp.Selection.NumberFormat = "@"
    $ExcelApp.Selection.Value2 = $ExcelApp.Selection.Value2
    $ProgressCounter = 1
    Foreach ($Cell in $Worksheet.Range("$($Column)2:$($Column)$LastRow").Cells) {
    Write-Progress  "Проверено $ProgressCounter значений из $LastRow в столбце $Column" -Status "Стараюсь как могу"
        [string]$ValueInCell = $Cell.Value() -replace "\.", ","
        if ($ValueInCell -notmatch "^\d+,\d{2}$") {
            if ($ValueInCell -match "^\d+$") {$Cell.Value2 = $ValueInCell + ",00"}
            if ($ValueInCell -match "^\d+,\d$") {$Cell.Value2 = $ValueInCell + "0"}
            if ($ValueInCell -match "^(\d+,\d\d)\d+$") {$Cell.Value2 = $ValueInCell -replace '^(\d+,\d\d)\d+$', '$1'}
        }
    $ProgressCounter += 1
    }
}
Truncate-ToTwoDecimalPlaces -Worksheet $Worksheet -ExcelApp $ExcelApp -Column "L"
Truncate-ToTwoDecimalPlaces -Worksheet $Worksheet -ExcelApp $ExcelApp -Column "K"
