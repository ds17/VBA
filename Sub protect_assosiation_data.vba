Sub protect_assosiation_data()

Application.ScreenUpdating = False
target_win = ThisWorkbook.Name

'————————————————Clear old data——————————————————————————————————————————————————————————————————————
'clear origin m-net
Sheet2.Activate
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

'clear origin cav
Sheet7.Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

'clear cav data
Sheet3.Activate
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

'clear m-net data
Sheet4.Activate
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

Sheet1.Activate

'————————————————improt CSV date——————————————————————————————————————————————————————————————————————
MsgBox ("M_Net File")

csvPath = Application.GetOpenFilename()
csvName = Right(csvPath, 25)
csvSheetName = Left(csvName, 21)

Workbooks.Open Filename:=csvPath
Range("A2").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Windows(target_win).Activate
Sheet2.Activate
    Range("C2").Select
    ActiveSheet.Paste
    Sheet2.Name = csvSheetName
Windows(csvName).Close

'————————————————import LOG data——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
MsgBox ("LOG File")
cavPath = Application.GetOpenFilename()
cavName = Right(cavPath, 24)
cavSheetName = Left(cavName, 20)
'listName = "List_" + cavSheetName

Sheet7.Activate
With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" + cavPath _
        , Destination:=Range("$A$1"))
        '.CommandType = 0
        .Name = cavName
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 936
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
        , 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False

        Sheet7.Name = cavSheetName
End With

'————————————————cpoy origin data to target sheet——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
Sheet6.Activate
With ActiveSheet
'————setup origin cols array————
    oc1 = Cells(3, 4).Value
    oc2 = Cells(4, 4).Value
    oc3 = Cells(5, 4).Value
    oc4 = Cells(6, 4).Value
    oc5 = Cells(7, 4).Value
    oc6 = Cells(8, 4).Value
    oc7 = Cells(9, 4).Value
    oc8 = Cells(10, 4).Value
    oc9 = Cells(11, 4).Value
    oc10 = Cells(12, 4).Value
    oc11 = Cells(13, 4).Value
    oc12 = Cells(14, 4).Value
    oc13 = Cells(15, 4).Value
    oc14 = Cells(16, 4).Value
    oc15 = Cells(17, 4).Value
    oc16 = Cells(18, 4).Value
    oc17 = Cells(19, 4).Value
    oc18 = Cells(20, 4).Value
    oc19 = Cells(21, 4).Value
    oc20 = Cells(22, 4).Value
    oc_arr = Array(oc1, oc2, oc3, oc4, oc5, oc6, oc7, oc8, oc9, oc10, oc11, oc12, oc13, oc14, oc15, oc16, oc17, oc18, oc19, oc20)

'————setup taget cols array————
    tc1 = Cells(3, 5).Value
    tc2 = Cells(4, 5).Value
    tc3 = Cells(5, 5).Value
    tc4 = Cells(6, 5).Value
    tc5 = Cells(7, 5).Value
    tc6 = Cells(8, 5).Value
    tc7 = Cells(9, 5).Value
    tc8 = Cells(10, 5).Value
    tc9 = Cells(11, 5).Value
    tc10 = Cells(12, 5).Value
    tc11 = Cells(13, 5).Value
    tc12 = Cells(14, 5).Value
    tc13 = Cells(15, 5).Value
    tc14 = Cells(16, 5).Value
    tc15 = Cells(17, 5).Value
    tc16 = Cells(18, 5).Value
    tc17 = Cells(19, 5).Value
    tc18 = Cells(20, 5).Value
    tc19 = Cells(21, 5).Value
    tc20 = Cells(22, 5).Value
    tc_arr = Array(tc1, tc2, tc3, tc4, tc5, tc6, tc7, tc8, tc9, tc10, tc11, tc12, tc13, tc14, tc15, tc16, tc17, tc18, tc19, tc20)
End With

'————recursive copy and paste————
For i = 0 To 19
    oc = oc_arr(i)
    tc = tc_arr(i)
    
    If i <= 11 Then
        Sheet2.Activate
        Cells(1, oc).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Sheet4.Activate
        Cells(1, tc).Select
        ActiveSheet.Paste
    Else
        Sheet7.Activate
        Cells(2, oc).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Sheet3.Activate
        Cells(2, tc).Select
        ActiveSheet.Paste
    End If
Next i
'——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————
'——————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————————

Sheet1.Activate
Application.ScreenUpdating = True
End Sub
