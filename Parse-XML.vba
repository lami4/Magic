    

Sub ParseXML()
    'Variables
    Dim fnd As String, FirstFound As String
    Dim FoundCell As Range, rng As Range
    Dim myRange As Range, LastCell As Range
    Dim MaxNumberOfSentences As Integer
    Set SearchRange = Range("H:H")
    Set LastSearchCell = SearchRange.Cells(SearchRange.Cells.Count)
    Set oXMLFile = CreateObject("Microsoft.XMLDOM")
    XMLFileName = "C:\Users\selyuto\Desktop\Magik\Test.xml"
    oXMLFile.Load (XMLFileName)
    'Add the required number of columns to populate with sentence parts
    MaxNumberOfSentences = 6
    For i = 1 To MaxNumberOfSentences
    'SearchRange.EntireColumn.Insert
    Columns("I:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
    Next i
    'Loop through sentences in the database
    Set SplitSentences = oXMLFile.SelectNodes("//split-sentence")
    For Each SplitSentence In SplitSentences
        Set SearchForWhat = SearchRange.Find(what:=SplitSentence.Text, after:=LastSearchCell, LookIn:=xlValues)
        If Not SearchForWhat Is Nothing Then
            FirstFound = SearchForWhat.Address
            Do
                MsgBox (SplitSentence.Text)
                MsgBox (SearchForWhat.Address)
                Set SearchForWhat = SearchRange.FindNext(SearchForWhat)
                'MsgBox (SplitSentence.Attributes(1).Value)
            Loop While (SearchForWhat.Address <> FirstFound)
        End If
    Next
    

    
End Sub
