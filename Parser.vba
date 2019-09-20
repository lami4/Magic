Sub CopyDataToAnotherFile(PiterOneFileFullName, PiterOneItemCodeColumn, PiterOneBColumn, PiterTwoFileFullName, PiterTwoItemCodeColumn, PiterTwoBColumn, PiterTwoAddBColumnFlag, PiterTwoDeleteCitiesFlag, PiterTwoDeleteEmptyRowsFlag, PeterTwoCodeColumnDeleteEmptyRows, NormalizePricesAndDiscountsFlag, BlackPriceColumnPiter2, RedPriceColumnPiter2, DiscountColumnPiter2, RemoveRedundantColumnsFlag)
    Dim PiterOneBook As Workbook
    Dim PiterTwoBook As Workbook
    Dim PiterOneBookItemCodeRange As Range
    Dim PiterTwoBookItemCodeRange As Range
    Dim PiterTwoBookItemCodeRangeCoordintates As String
    Dim LastRow As String
    Dim PiterOneBookColumnBCoordinates As String
    Dim PiterOneBookCorrectNameColumnCoordinates As String
    Dim CodeItemValue As String
    Dim CurrentRow As String
    Dim CurrentRowBookTwo As String
    Dim myRange As Range
    Dim regEx As New RegExp
    Dim strInput As String
    Dim ParsingTrigger As Boolean
    ParsingTrigger = False
    'Search function settings
    Dim FirstFound As String
    Dim FoundCell As Range
    Dim SearchRange As Range
    Dim LastCell As Range
    'Search function settings
    Dim PathToPiterOne As String: PathToPiterOne = PiterOneFileFullName '''
    Dim PathToPiterTwo As String: PathToPiterTwo = PiterTwoFileFullName '''
    Dim PriceColumnOne As String: PriceColumnOne = BlackPriceColumnPiter2 ''' BlackPrice
    Dim PriceColumnTwo As String: PriceColumnTwo = RedPriceColumnPiter2 ''' RedPrice
    Dim DiscountColumn As String: DiscountColumn = DiscountColumnPiter2 ''' Discount
    'PiterOneItemCodeColumn = "I" '''
    'PiterOneBColumn = "AA" '''
    'PiterOneCorrectNameColumn = "AB" '''
    'PiterTwoItemCodeColumn = "I" '''
    'PiterTwoBColumn = "AA" '''
    'PeterTwoCodeColumnDeleteEmptyRows = "I" '''
    'PiterOneCorrectNameColumnNumber = Range(PiterOneCorrectNameColumn & 1).Column

    'Open Piter_2
    Set PiterOneBook = Workbooks.Open(PathToPiterOne)
    'MsgBox (PiterOneBook.FullName)
    Set PiterTwoBook = Workbooks.Open(PathToPiterTwo)
    'MsgBox (PiterTwoBook.FullName)
    PiterOneBColumnNumber = Range(PiterOneBColumn & 1).Column
    PiterTwoBColumnNumber = Range(PiterTwoBColumn & 1).Column
    On Error Resume Next
    PiterOneBook.Worksheets(1).ShowAllData
    On Error Resume Next
    PiterTwoBook.Worksheets(1).ShowAllData
    'DELETING CITIES
    If PiterTwoDeleteCitiesFlag = "true" Then
        Dim CitiesRange As Range
        Dim CitiesRangeCoordinates As String
        On Error Resume Next
        PiterTwoBook.Worksheets(1).ShowAllData
        LastRow = PiterTwoBook.Worksheets(1).Cells(PiterTwoBook.Worksheets(1).Rows.Count, "A").End(xlUp).Row
        'MsgBox (LastRow)
        Dim BannedCities(1 To 24) As String
        BannedCities(1) = "Àñòðàõàíü"
        BannedCities(2) = "Âîëãîãðàä"
        BannedCities(3) = "Âîðîíåæ"
        BannedCities(4) = "Åêàòåðèíáóðã"
        BannedCities(5) = "Èâàíîâî"
        BannedCities(6) = "Èðêóòñê"
        BannedCities(7) = "Êðàñíîäàð"
        BannedCities(8) = "Êðàñíîÿðñê"
        BannedCities(9) = "Ëèïåöê"
        BannedCities(10) = "Ìîñêâà"
        BannedCities(11) = "Ìóðìàíñê"
        BannedCities(12) = "Íèæíèé Íîâãîðîä"
        BannedCities(13) = "Íîâîñèáèðñê"
        BannedCities(14) = "Îìñê"
        BannedCities(15) = "Îðåíáóðã"
        BannedCities(16) = "Ðîñòîâ"
        BannedCities(17) = "Ñàðàòîâ"
        BannedCities(18) = "Ñî÷è"
        BannedCities(19) = "Ñòàâðîïîëü"
        BannedCities(20) = "Ñóðãóò"
        BannedCities(21) = "Ñûêòûâêàð"
        BannedCities(22) = "Òîëüÿòòè"
        BannedCities(23) = "Óôà"
        BannedCities(24) = "Òþìåíü"
        For Each City In BannedCities
            CitiesRangeCoordinates = "A2:A" + LastRow
            Set CitiesRange = PiterTwoBook.Worksheets(1).Range(CitiesRangeCoordinates)
            If Not IsError(Application.Match(City, CitiesRange, 0)) Then
                CitiesRange.AutoFilter Field:=1, Criteria1:=City
                CitiesRangeCoordinates = "A2:A" + LastRow
                Set CitiesRange = PiterTwoBook.Worksheets(1).Range(CitiesRangeCoordinates)
                CitiesRange.Select
                Selection.EntireRow.Delete
                PiterTwoBook.Worksheets(1).ShowAllData
            End If
        Next City
    End If
    'DELETING CITIES
    'INSERTING COLUMN "CORRECT NAME"
    If PiterTwoAddBColumnFlag = "true" Then
        LastColumn = PiterTwoBook.Worksheets(1).Cells(1, PiterTwoBook.Worksheets(1).Columns.Count).End(xlToLeft).Column
        LastColumn = LastColumn + 1
        PiterTwoBook.Worksheets(1).Columns(LastColumn).Cells(1).Value = "/Âåðíîå íàèìåíîâàíèå/"
    End If
    'INSERTING COLUMN "CORRECT NAME"
    'DELETING EMPTY CODE ROWS
    If PiterTwoDeleteEmptyRowsFlag = "true" Then
        LastRow = PiterTwoBook.Worksheets(1).Cells(PiterTwoBook.Worksheets(1).Rows.Count, "A").End(xlUp).Row
        RangeCoordinates = PeterTwoCodeColumnDeleteEmptyRows + "2:" + PeterTwoCodeColumnDeleteEmptyRows + LastRow
        Set myRange = PiterTwoBook.Worksheets(1).Range(RangeCoordinates)
        myRange.AutoFilter Field:=9, Criteria1:=""
        RangeCoordinates = "A2:" + "A" + LastRow
        Set myRange = PiterTwoBook.Worksheets(1).Range(RangeCoordinates)
        myRange.Select
        Selection.EntireRow.Delete
        PiterTwoBook.Worksheets(1).ShowAllData
    End If
    'DELETING EMPTY CODE ROWS
    'TRUNCATING PRICES AND ROUNDING DISCOUNTS
    On Error Resume Next
    PiterTwoBook.Worksheets(1).ShowAllData
    LastRow = PiterTwoBook.Worksheets(1).Cells(PiterTwoBook.Worksheets(1).Rows.Count, "A").End(xlUp).Row
    If NormalizePricesAndDiscountsFlag = "true" Then
        'PRICE COLUMN ONE
        RangeCoordinates = PriceColumnOne + "2:" + PriceColumnOne + LastRow
        PiterTwoBook.Worksheets(1).Range(RangeCoordinates).Select
        Selection.NumberFormat = "@"
        'MsgBox (RangeCoordinates)
        'Set a range for PriceColumnOne and loop through it
        Set myRange = PiterTwoBook.Worksheets(1).Range(RangeCoordinates)
        For Each Cell In myRange
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
        PiterTwoBook.Worksheets(1).Range(RangeCoordinates).Select
        Selection.NumberFormat = "@"
        'Set a range for PriceColumnTwo and loop through it
        Set myRange = PiterTwoBook.Worksheets(1).Range(RangeCoordinates)
        For Each Cell In myRange
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
        PiterTwoBook.Worksheets(1).Range(RangeCoordinates).Select
        Selection.NumberFormat = "@"
        'Set a range for PriceColumnTwo and loop through it
        Set myRange = PiterTwoBook.Worksheets(1).Range(RangeCoordinates)
        Dim CellValue As Double
        For Each Cell In myRange
            CellValue = Cell.Value
            RoundedValue = Round(CellValue)
            Cell.Value = RoundedValue
        Next
    End If
    'TRUNCATING PRICES AND ROUNDING DISCOUNTS
    
    'MsgBox (PiterTwoBookItemCodeRangeCoordintates)
    'MsgBox (SearchRange.Cells.Count)
    'Get last row in the Item code column
    On Error Resume Next
    PiterOneBook.Worksheets(1).ShowAllData
    On Error Resume Next
    PiterTwoBook.Worksheets(1).ShowAllData
    'Piter2 search range for Code Item Column
    PiterTwoBookItemCodeRangeCoordintates = PiterTwoItemCodeColumn + ":" + PiterTwoItemCodeColumn
    Set SearchRange = PiterTwoBook.Worksheets(1).Range(PiterTwoBookItemCodeRangeCoordintates)
    Set LastCell = SearchRange.Cells(SearchRange.Cells.Count)
    'Piter2 search range for Code Item Column
    LastRow = PiterOneBook.Worksheets(1).Cells(PiterOneBook.Worksheets(1).Rows.Count, PiterOneItemCodeColumn).End(xlUp).Row
    'Creates coordinates for the Item code column
    PiterOneBookItemCodeRangeCoordintates = PiterOneItemCodeColumn + "2" + ":" + PiterOneItemCodeColumn + LastRow
    'Message with coordinates
    'MsgBox (PiterOneBookItemCodeRangeCoordintates)
    Set PiterOneBookItemCodeRange = PiterOneBook.Worksheets(1).Range(PiterOneBookItemCodeRangeCoordintates)
    For Each CodeItem In PiterOneBookItemCodeRange
        CodeItemValue = CodeItem.Value
        CurrentRow = CodeItem.Row
        If CodeItemValue <> "" Then
            'GET VALUE OF THE CURRENT ITEM CODE
            'MsgBox (CodeItemValue)
            'GET VALUE OF CORRESPONDING CELL IN COLUMN /B/
            If PiterOneBook.Worksheets(1).Columns(PiterOneBColumnNumber).Cells(CurrentRow).Value <> "" Then
                'MsgBox (PiterOneBook.Worksheets(1).Columns(PiterOneBColumnNumber).Cells(CurrentRow).Value)
                'Find the same itemcode in Piter_2 and paste the value to column b
                Set FoundCell = SearchRange.Find(what:=CodeItemValue, after:=LastCell, LookIn:=xlValues)
                If Not FoundCell Is Nothing Then
                    FirstFound = FoundCell.Address
                    Do
                        'MsgBox ("found")
                        PiterTwoBook.Worksheets(1).Columns(PiterTwoBColumnNumber).Cells(FoundCell.Row).Value = PiterOneBook.Worksheets(1).Columns(PiterOneBColumnNumber).Cells(CurrentRow).Value
                        Set FoundCell = SearchRange.FindNext(FoundCell)
                    Loop While (FoundCell.Address <> FirstFound)
                End If
            End If
        End If
    Next CodeItem
'DELETING UNWANTED COLUMNS
    If RemoveRedundantColumnsFlag = "true" Then
        Dim BannedColumnNames(1 To 12) As String
        BannedColumnNames(1) = "ÊàòåãîðèÿÊÌ"
        BannedColumnNames(2) = "Îñí.ØÊ"
        BannedColumnNames(3) = "Êîììåíò Ìàðêåòèíã"
        BannedColumnNames(4) = "Êîììåíò ÊÌ"
        BannedColumnNames(5) = "Êîäïîñòàâùèêà"
        BannedColumnNames(6) = "Ïîñòàâùèê"
        BannedColumnNames(7) = "IdÈÄ Êàò.-äàòû äåéñòâèÿ"
        BannedColumnNames(8) = "Íàçâàíèå àêöèè"
        BannedColumnNames(9) = "ÒÎ ïëàí.,øò"
        BannedColumnNames(10) = "ÒÎ ïëàí.,ðóá."
        BannedColumnNames(11) = "Kpi14 Ðóá"
        BannedColumnNames(12) = "Kpi14 Øò"
        LastColumn = PiterTwoBook.Worksheets(1).Cells(1, PiterTwoBook.Worksheets(1).Columns.Count).End(xlToLeft).Column
        'MsgBox (LastColumn)
        For t = LastColumn To 1 Step -1
            If IsInArray(PiterTwoBook.Worksheets(1).Rows(1).Cells(t).Value, BannedColumnNames) Then
                PiterTwoBook.Worksheets(1).Columns(t).Delete
            End If
        Next t
    End If
'DELETING UNWANTED COLUMNS
End Sub
Sub ParseRegister(ItemCode, BlackPrice, RedPrice, Discount, DeleteEmptyRowsFlag, NormalizePricesAndDiscountsFlag, RemoveRedundantColumnsFlag, AddBColumnFlag, DeleteCitiesFlag)
    'Declare variables
    Dim PriceColumnOne As String: PriceColumnOne = BlackPrice
    Dim PriceColumnTwo As String: PriceColumnTwo = RedPrice
    Dim DiscountColumn As String: DiscountColumn = Discount
    Dim CodeColumn As String: CodeColumn = ItemCode
    Dim regEx As New RegExp
    Dim strInput As String
    Dim myRange As Range
    Dim ParsingTrigger As Boolean
    Dim RangeCoordinates As String
    Dim LastRow As String
    Dim LastColumn As Integer
    ParsingTrigger = False
    'Code
    Dim n As Name
    For Each n In ActiveWorkbook.Names
        n.Delete
    Next
    On Error Resume Next
    ActiveWorkbook.Worksheets(1).ShowAllData
    LastRow = ActiveWorkbook.Worksheets(1).Cells(ActiveWorkbook.Worksheets(1).Rows.Count, "A").End(xlUp).Row
    'TRUNCATING PRICES AND ROUNDING DISCOUNTS
    If NormalizePricesAndDiscountsFlag = "true" Then
        'PRICE COLUMN ONE
        RangeCoordinates = PriceColumnOne + "2:" + PriceColumnOne + LastRow
        ActiveWorkbook.Worksheets(1).Range(RangeCoordinates).Select
        Selection.NumberFormat = "@"
        'Set a range for PriceColumnOne and loop through it
        Set myRange = ActiveSheet.Range(RangeCoordinates)
        For Each Cell In myRange
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
        Set myRange = ActiveSheet.Range(RangeCoordinates)
        For Each Cell In myRange
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
        Set myRange = ActiveSheet.Range(RangeCoordinates)
        Dim CellValue As Double
        For Each Cell In myRange
            CellValue = Cell.Value
            RoundedValue = Round(CellValue)
            Cell.Value = RoundedValue
        Next
    End If
    'TRUNCATING PRICES AND ROUNDING DISCOUNTS
    'DELETING EMPTY CODE ROWS
    If DeleteEmptyRowsFlag = "true" Then
        RangeCoordinates = CodeColumn + "2:" + CodeColumn + LastRow
        Set myRange = ActiveSheet.Range(RangeCoordinates)
        myRange.AutoFilter Field:=9, Criteria1:=""
        RangeCoordinates = "A2:" + "A" + LastRow
        Set myRange = ActiveSheet.Range(RangeCoordinates)
        myRange.Select
        Selection.EntireRow.Delete
        ActiveWorkbook.Worksheets(1).ShowAllData
    End If
    'DELETING EMPTY CODE ROWS
    'DELETING UNWANTED COLUMNS
    If RemoveRedundantColumnsFlag = "true" Then
        Dim BannedColumnNames(1 To 12) As String
        BannedColumnNames(1) = "ÊàòåãîðèÿÊÌ"
        BannedColumnNames(2) = "Îñí.ØÊ"
        BannedColumnNames(3) = "Êîììåíò Ìàðêåòèíã"
        BannedColumnNames(4) = "Êîììåíò ÊÌ"
        BannedColumnNames(5) = "Êîäïîñòàâùèêà"
        BannedColumnNames(6) = "Ïîñòàâùèê"
        BannedColumnNames(7) = "IdÈÄ Êàò.-äàòû äåéñòâèÿ"
        BannedColumnNames(8) = "Íàçâàíèå àêöèè"
        BannedColumnNames(9) = "ÒÎ ïëàí.,øò"
        BannedColumnNames(10) = "ÒÎ ïëàí.,ðóá."
        BannedColumnNames(11) = "Kpi14 Ðóá"
        BannedColumnNames(12) = "Kpi14 Øò"
        LastColumn = ActiveWorkbook.Worksheets(1).Cells(1, ActiveWorkbook.Worksheets(1).Columns.Count).End(xlToLeft).Column
        'MsgBox (LastColumn)
        For t = LastColumn To 1 Step -1
            If IsInArray(ActiveWorkbook.Worksheets(1).Rows(1).Cells(t).Value, BannedColumnNames) Then
                ActiveWorkbook.Worksheets(1).Columns(t).Delete
            End If
        Next t
    End If
    'DELETING CITIES
    If DeleteCitiesFlag = "true" Then
        Dim CitiesRange As Range
        Dim CitiesRangeCoordinates As String
        On Error Resume Next
        ActiveWorkbook.Worksheets(1).ShowAllData
        LastRow = ActiveWorkbook.Worksheets(1).Cells(ActiveWorkbook.Worksheets(1).Rows.Count, "A").End(xlUp).Row
        Dim BannedCities(1 To 24) As String
        BannedCities(1) = "Àñòðàõàíü"
        BannedCities(2) = "Âîëãîãðàä"
        BannedCities(3) = "Âîðîíåæ"
        BannedCities(4) = "Åêàòåðèíáóðã"
        BannedCities(5) = "Èâàíîâî"
        BannedCities(6) = "Èðêóòñê"
        BannedCities(7) = "Êðàñíîäàð"
        BannedCities(8) = "Êðàñíîÿðñê"
        BannedCities(9) = "Ëèïåöê"
        BannedCities(10) = "Ìîñêâà"
        BannedCities(11) = "Ìóðìàíñê"
        BannedCities(12) = "Íèæíèé Íîâãîðîä"
        BannedCities(13) = "Íîâîñèáèðñê"
        BannedCities(14) = "Îìñê"
        BannedCities(15) = "Îðåíáóðã"
        BannedCities(16) = "Ðîñòîâ"
        BannedCities(17) = "Ñàðàòîâ"
        BannedCities(18) = "Ñî÷è"
        BannedCities(19) = "Ñòàâðîïîëü"
        BannedCities(20) = "Ñóðãóò"
        BannedCities(21) = "Ñûêòûâêàð"
        BannedCities(22) = "Òîëüÿòòè"
        BannedCities(23) = "Óôà"
        BannedCities(24) = "Òþìåíü"
        For Each City In BannedCities
            CitiesRangeCoordinates = "A2:A" + LastRow
            Set CitiesRange = ActiveSheet.Range(CitiesRangeCoordinates)
            If Not IsError(Application.Match(City, CitiesRange, 0)) Then
                CitiesRange.AutoFilter Field:=1, Criteria1:=City
                CitiesRangeCoordinates = "A2:A" + LastRow
                Set CitiesRange = ActiveSheet.Range(CitiesRangeCoordinates)
                CitiesRange.Select
                Selection.EntireRow.Delete
                ActiveWorkbook.Worksheets(1).ShowAllData
            End If
        Next City
    End If
    'DELETING CITIES
    'INSERTING COLUMN /B/
    If AddBColumnFlag = "true" Then
        LastColumn = ActiveWorkbook.Worksheets(1).Cells(1, ActiveWorkbook.Worksheets(1).Columns.Count).End(xlToLeft).Column
        LastColumn = LastColumn + 1
        ActiveWorkbook.Worksheets(1).Columns(LastColumn).Cells(1).Value = "/B/"
    End If
    'INSERTING COLUMN /B/
End Sub
Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function
