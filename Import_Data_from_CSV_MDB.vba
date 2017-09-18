Sub 数据汇总()

Application.ScreenUpdating = False
target_win = ThisWorkbook.Name

'————————————————原数据清空——————————————————————————————————————————————————————————————————————
Sheet2.Activate
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

Sheet3.Activate
    Range("d2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

Sheet4.Activate
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

Sheet1.Activate
    
'————————————————导入CSV数据——————————————————————————————————————————————————————————————————————
MsgBox ("M_Net文件")


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

'————————————————导入CAV数据——————————————————————————————————————————————————————————————————————————————————————————————————————————————————
MsgBox ("CAV文件")
cavPath = Application.GetOpenFilename()
cavName = Right(cavPath, 24)
cavSheetName = Left(cavName, 20)
listName = "表_" + cavSheetName

Sheet4.Activate
With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=cavPath;Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Pa" _
        , _
        "ssword="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Tr" _
        , _
        "ansactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLED" _
        , _
        "B:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Comple" _
        , _
        "x Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validati" _
        , "on=False"), Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdTable
        .CommandText = Array("CurveData")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .SourceDataFile = _
        cavPath
        .ListObject.DisplayName = listName
        .Refresh BackgroundQuery:=False
    End With
Sheet4.ListObjects(1).unlist

'复制粘贴
Sheet4.Activate
    Range("D2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
Sheet3.Activate
    Range("D2").Select
    ActiveSheet.Paste
Sheet3.Name = cavSheetName

Sheet1.Activate
Application.ScreenUpdating = True
End Sub
