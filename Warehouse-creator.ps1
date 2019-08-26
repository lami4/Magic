$script:PathToFile = "C:\Users\Tsedik\Desktop\Magik\Test.xlsx"
$script:OutputFilePath = "C:\Users\Tsedik\Desktop\Magik\Test.xml"
$ColumnRangeStart = "H"
$ColumnRangeEnd = "N"

Function Create-NewXMLObject ($OutputXmlFile, $InnerCellText, $SelectedColumnRange, $CellCounterValue) {
    $Element = $OutputXmlFile.CreateNode("element","split-sentence",$null)
    $Element.InnerText = $InnerCellText
    $ColumnCounter = 1
    Foreach ($Column in $SelectedColumnRange.Columns) {
        if ($Column.Address($false, $false, 1) -eq "$($ColumnRangeStart):$($ColumnRangeStart)") {continue}
        $ElementAttribute = $OutputXmlFile.CreateAttribute("substring-$ColumnCounter")
        $ElementAttribute.Value = $Column.Cells.Item($CellCounterValue).Value()
        $Element.Attributes.Append($ElementAttribute)
        $ColumnCounter += 1
    }
    $OutputXmlFile.SelectSingleNode("/list-of-split-sentences").AppendChild($Element)
}

$ExcelApp = New-Object -ComObject Excel.Application
$ExcelApp.Visible = $false
$Workbook = $ExcelApp.Workbooks.Open($script:PathToFile)
$Worksheet = $Workbook.Worksheets.Item(1)
#Remove filters
try {$Worksheet.ShowAllData()} catch {"ShowAllData already applied"}
#Create column range
$SelectedRange = $Worksheet.Range("$($ColumnRangeStart):$($ColumnRangeEnd)")
#Find out number of columns in the range
Write-Host $SelectedRange.Columns.Count
Write-Host $SelectedRange.Columns.Item(1).Address($false, $false, 1)
#Finding out number of the last filled raw in the column that contains original sentences
#$LastRow = $Worksheet.Cells.Item($Worksheet.Rows.Count, "$ColumnRangeStart").End(-4162).Row
#FAKE LASTROW
$LastRow = 10
Write-Host $LastRow
#Create an XML file
$OutputXmlFile = New-Object System.Xml.XmlDocument
$OutputXmlFile.CreateXmlDeclaration("1.0","UTF-8",$null)
$OutputXmlFile.AppendChild($OutputXmlFile.CreateXmlDeclaration("1.0","UTF-8",$null))
$InfoForXml = @"
File was generated: $(Get-Date)
"@
$OutputXmlFile.AppendChild($OutputXmlFile.CreateComment($InfoForXml))
$RootElement = $OutputXmlFile.CreateNode("element","list-of-split-sentences",$null)
$OutputXmlFile.AppendChild($RootElement)
#Here add the function
$CellCounter = 0
Foreach ($Cell in $SelectedRange.Columns.Item(1).Cells) {
    $CellCounter += 1
    Write-Host $Cell.Value()
    Create-NewXMLObject -OutputXmlFile $OutputXmlFile -InnerCellText $Cell.Value() -SelectedColumnRange $SelectedRange -CellCounterValue $CellCounter
    if ($CellCounter -eq $LastRow) {break}
}

$OutputXmlFile.Save($script:OutputFilePath)
#
