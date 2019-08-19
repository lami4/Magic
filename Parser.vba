Sub TruncateToTwoDecimalPlaces(ItemCode, BlackPrice, RedPrice, Discount)
    'Declare variables
    Dim PriceColumnOne As String: PriceColumnOne = BlackPrice
    Dim PriceColumnTwo As String: PriceColumnTwo = RedPrice
    Dim DiscountColumn As String: DiscountColumn = Discount
    Dim CodeColumn As String: CodeColumn = ItemCode
    Dim regEx As New RegExp
    Dim strInput As String
    Dim MyRange As Range
    Dim ParsingTrigger As Boolean
    Dim RangeCoordinates As String
    Dim LastRow As String
    ParsingTrigger = False
    'Code
    On Error Resume Next
    ActiveWorkbook.Worksheets(1).ShowAllData
    LastRow = ActiveWorkbook.Worksheets(1).Cells(ActiveWorkbook.Worksheets(1).Rows.Count, "A").End(xlUp).Row
    'DELETING EMPTY CODE ROWS
    RangeCoordinates = CodeColumn + "2:" + CodeColumn + LastRow
    Set MyRange = ActiveSheet.Range(RangeCoordinates)
    MyRange.AutoFilter Field:=9, Criteria1:=""
    RangeCoordinates = "A2:" + "A" + LastRow
    Set MyRange = ActiveSheet.Range(RangeCoordinates)
    MyRange.Select
    Selection.EntireRow.Delete
    ActiveWorkbook.Worksheets(1).ShowAllData
    'DELETING EMPTY CODE ROWS
    'PRICE COLUMN ONE
    RangeCoordinates = PriceColumnOne + "2:" + PriceColumnOne + LastRow
    ActiveWorkbook.Worksheets(1).Range(RangeCoordinates).Select
    Selection.NumberFormat = "@"
    'Set a range for PriceColumnOne and loop through it
    Set MyRange = ActiveSheet.Range(RangeCoordinates)
    For Each Cell In MyRange
        strInput = Cell.Value
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = "^\d+,\d{2}$"
        End With
        'Check if a value is an integer
        If regEx.Test(strInput) Then
            ParsingTrigger = False
        Else
            ParsingTrigger = True
        End If
        'If a value is not an integer, execute this code
        If ParsingTrigger = True Then
            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = "^\d+$"
            End With
            If regEx.Test(strInput) Then
                NewStringValue = strInput + ",00"
                Cell.Value = NewStringValue
            End If
            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = "^\d+,\d$"
            End With
            If regEx.Test(strInput) Then
                NewStringValue = strInput + "0"
                Cell.Value = NewStringValue
            End If
            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = "^(\d+,\d\d)\d+$"
            End With
            If regEx.Test(strInput) Then
                NewStringValue = regEx.Replace(strInput, "$1")
                Cell.Value = NewStringValue
            End If
        End If
    Next
    'PRICE COLUMN TWO
    RangeCoordinates = PriceColumnTwo + "2:" + PriceColumnTwo + LastRow
    ActiveWorkbook.Worksheets(1).Range(RangeCoordinates).Select
    Selection.NumberFormat = "@"
    'Set a range for PriceColumnTwo and loop through it
    Set MyRange = ActiveSheet.Range(RangeCoordinates)
    For Each Cell In MyRange
        strInput = Cell.Value
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = "^\d+,\d{2}$"
        End With
        'Check if a value is an integer
        If regEx.Test(strInput) Then
            ParsingTrigger = False
        Else
            ParsingTrigger = True
        End If
        'If a value is not an integer, execute this code
        If ParsingTrigger = True Then
            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = "^\d+$"
            End With
            If regEx.Test(strInput) Then
                NewStringValue = strInput + ",00"
                Cell.Value = NewStringValue
            End If
            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = "^\d+,\d$"
            End With
            If regEx.Test(strInput) Then
                NewStringValue = strInput + "0"
                Cell.Value = NewStringValue
            End If
            With regEx
                .Global = True
                .MultiLine = True
                .IgnoreCase = False
                .Pattern = "^(\d+,\d\d)\d+$"
            End With
            If regEx.Test(strInput) Then
                NewStringValue = regEx.Replace(strInput, "$1")
                Cell.Value = NewStringValue
            End If
        End If
    Next
    'DISCOUNT COLUMN
    RangeCoordinates = DiscountColumn + "2:" + DiscountColumn + LastRow
    ActiveWorkbook.Worksheets(1).Range(RangeCoordinates).Select
    Selection.NumberFormat = "@"
    'Set a range for PriceColumnTwo and loop through it
    Set MyRange = ActiveSheet.Range(RangeCoordinates)
    Dim CellValue As Double
    For Each Cell In MyRange
    'Write code here
    CellValue = Cell.Value
    RoundedValue = Round(CellValue)
    Cell.Value = RoundedValue
    Next
    'DELETING UNWANTED COLUMNS
    Dim LastColumn As Integer
    Dim BannedColumnNames(1 To 12) As String
    BannedColumnNames(1) = "КатегорияКМ"
    BannedColumnNames(2) = "Осн.ШК"
    BannedColumnNames(3) = "Коммент Маркетинг"
    BannedColumnNames(4) = "Коммент КМ"
    BannedColumnNames(5) = "Кодпоставщика"
    BannedColumnNames(6) = "Поставщик"
    BannedColumnNames(7) = "IdИД Кат.-даты действия"
    BannedColumnNames(8) = "Название акции"
    BannedColumnNames(9) = "ТО план.,шт"
    BannedColumnNames(10) = "ТО план.,руб."
    BannedColumnNames(11) = "Kpi14 Руб"
    BannedColumnNames(12) = "Kpi14 Шт"
    LastColumn = ActiveWorkbook.Worksheets(1).Cells(1, ActiveWorkbook.Worksheets(1).Columns.Count).End(xlToLeft).Column
    'MsgBox (LastColumn)
    For t = LastColumn To 1 Step -1
    If IsInArray(ActiveWorkbook.Worksheets(1).Rows(1).Cells(t).Value, BannedColumnNames) Then
        ActiveWorkbook.Worksheets(1).Columns(t).Delete
    End If
Next t
End Sub
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
