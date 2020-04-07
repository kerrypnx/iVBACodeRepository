#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr) 'MS Office 64 Bit
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long) 'MS Office 32 Bit
#End If
Dim fso As Object, oFolder As Object, oFile As Object
Dim StrPt, Arrow As String, iLoop As Integer, wb As Workbook, ws As Worksheet, StrName As String, iName As Integer
Dim iStart As Long, iEnd As Long, StrCode As String, rng As Range, iLoopEnd As Integer
Public Const TaxSheets3StartRowNum = 3
Public Const TaxSheets4StartRowNum = 3
Public Const TaxSheets5StartRowNum = 3
Public Const TaxSheets6StartRowNum = 3
Public Const TaxSheets7StartRowNum = 3
Public Const TaxSheets8StartRowNum = 3
Public Const TaxSheets9StartRowNum = 3
Public Const TaxSheets10StartRowNum = 3
Public Const TaxSheets12StartRowNum = 5
Public Const DelaySheets2StartRowNum = 5
Public Const DelaySheets3StartRowNum = 5
Public Const FilterSheets2StartRowNum = 5
Public Const FilterSheets3StartRowNum = 3
Public Const FilterSheets4StartRowNum = 6
Public Const FilterSheets5StartRowNum = 7
Function getCompanyRangeStartRowIndex(startRowIndex As Integer, currentSheetStartRowIndex As Long) As Long
      getCompanyRangeStartRowIndex = IIf(startRowIndex = currentSheetStartRowIndex, startRowIndex, currentSheetStartRowIndex)
End Function
Function getCompanyRangeLastRowIndex(CompanyRangeLastRowIndex As Long, currentSheetLastRowIndex As Long) As Long
      getCompanyRangeLastRowIndex = IIf(CompanyRangeLastRowIndex = currentSheetLastRowIndex, CompanyRangeLastRowIndex, currentSheetLastRowIndex)
End Function
Public Function isOnlyNum(str As String) As Boolean
    Dim reg As Object
    Set reg = CreateObject("VBScript.Regexp")
            
    Dim is_exist As Boolean
    With reg
        .Global = True
        .Pattern = "^\d+$"
        is_exist = .Test(str)
    End With
    isOnlyNum = is_exist
End Function
Sub TurnEverythingOff()
With Application
    .Calculation = xlCalculationManual
    .EnableEvents = False
    .DisplayAlerts = False
    .ScreenUpdating = False
End With
End Sub

Sub TurnEverythingOn()
With Application
    .Calculation = xlCalculationAutomatic
    .EnableEvents = True
    .DisplayAlerts = True
    .ScreenUpdating = True
End With
End Sub
Function createCalculateTaxSheetStartRowDict() As Dictionary
Dim sheetStartRowNum As Object
    Set sheetStartRowNum = CreateObject("Scripting.Dictionary")
    sheetStartRowNum(3) = TaxSheets3StartRowNum
    sheetStartRowNum(4) = TaxSheets4StartRowNum
    sheetStartRowNum(5) = TaxSheets5StartRowNum
    sheetStartRowNum(6) = TaxSheets6StartRowNum
    sheetStartRowNum(7) = TaxSheets7StartRowNum
    sheetStartRowNum(8) = TaxSheets8StartRowNum
    sheetStartRowNum(9) = TaxSheets9StartRowNum
    sheetStartRowNum(10) = TaxSheets10StartRowNum
    sheetStartRowNum(12) = TaxSheets12StartRowNum
    Set createCalculateTaxSheetStartRowDict = sheetStartRowNum
End Function
Function createDelaySheetStartRowDict() As Dictionary
Dim sheetStartRowNum As Object
    Set sheetStartRowNum = CreateObject("Scripting.Dictionary")
    sheetStartRowNum(2) = DelaySheets2StartRowNum
    sheetStartRowNum(3) = DelaySheets3StartRowNum
    Set createDelaySheetStartRowDict = sheetStartRowNum
End Function
Function createFilterSheetStartRowDict() As Dictionary
Dim sheetStartRowNum As Object
    Set sheetStartRowNum = CreateObject("Scripting.Dictionary")
    sheetStartRowNum(2) = FilterSheets2StartRowNum
    sheetStartRowNum(3) = FilterSheets3StartRowNum
    sheetStartRowNum(4) = FilterSheets4StartRowNum
    sheetStartRowNum(5) = FilterSheets5StartRowNum
    Set createFilterSheetStartRowDict = sheetStartRowNum
End Function
Sub dealWithCalculateTaxSheet3ToSheet6(sheetStartRowNum As Object, StrCode As String)
        For iSht = 3 To 6
            Set ws = wb.Sheets(iSht)
            With ws
                .UsedRange.AutoFilter
                Set rng = ws.Range("B:B").Find(StrCode, , , , , xlNext)
                If Not rng Is Nothing Then
                    iStart = rng.Row
'
                    Set rng = Nothing
                    Set rng = ws.Range("B:B").Find(StrCode, , , , , xlPrevious)
                    iEnd = rng.Row
                    Set rng = Nothing

                    If iStart = sheetStartRowNum(iSht) Then
                        .Rows(iEnd + 1 & ":" & .Rows.Count).Delete Shift:=xlUp
                    Else
                        If iEnd = ws.Rows.Count Then
                            .Rows(sheetStartRowNum(iSht) & ":" & iStart - 1).Delete Shift:=xlUp
                        Else
                            .Rows(iEnd + 1 & ":" & .Rows.Count).Delete Shift:=xlUp
                            .Rows(sheetStartRowNum(iSht) & ":" & iStart - 1).Delete Shift:=xlUp
                        End If
                    End If
                Else
                    If Len(sheetStartRowNum(iSht)) > 0 Then
                        .Rows(sheetStartRowNum(iSht) & ":" & .Rows.Count).Delete Shift:=xlUp
                    Else
                        .Rows(sheetStartRowNum & ":" & .Rows.Count).Delete Shift:=xlUp
                    End If
                End If
            End With
            Set ws = Nothing
    Next iSht
End Sub
Sub dealWithCalculateTaxSheet7ToSheet10(sheetStartRowNum As Object, StrCode As String)
For iSht = 7 To 10
            Set ws = wb.Sheets(iSht)
            With ws
                .UsedRange.AutoFilter
                Set rng = ws.Range("A:A").Find(StrCode, , , , , xlNext)
                If Not rng Is Nothing Then
                    iStart = rng.Row
                    Set rng = Nothing
                    Set rng = ws.Range("A:A").Find(StrCode, , , , , xlPrevious)
                    iEnd = rng.Row
                    Set rng = Nothing

                    If iStart = sheetStartRowNum(iSht) Then
                        .Rows(iEnd + 1 & ":" & .Rows.Count).Delete Shift:=xlUp
                    Else
                        If iEnd = .Range("A100000").End(xlUp).Row Then
                            .Rows(sheetStartRowNum(iSht) & ":" & iStart - 1).Delete Shift:=xlUp
                        Else
                            .Rows(iEnd + 1 & ":" & .Rows.Count).Delete Shift:=xlUp
                            .Rows(sheetStartRowNum(iSht) & ":" & iStart - 1).Delete Shift:=xlUp
                        End If
                    End If
                Else
                    .Rows(sheetStartRowNum(iSht) & ":" & .Rows.Count).Delete Shift:=xlUp
                End If
            End With
            Set ws = Nothing
        Next iSht
End Sub
Sub dealWithCalculateTaxSheet12(sheetStartRowNum As Object, StrCode As String)
Set ws = wb.Sheets(12)
        With ws
            .UsedRange.AutoFilter
            Set rng = ws.Range("A:A").Find(StrCode, , , , , xlNext)
            If Not rng Is Nothing Then
                iStart = rng.Row
                Set rng = Nothing
                Set rng = ws.Range("A:A").Find(StrCode, , , , , xlPrevious)
                iEnd = rng.Row
                Set rng = Nothing
                
                If iStart = sheetStartRowNum(12) Then
                    .Rows(iEnd + 1 & ":" & .Rows.Count).Delete Shift:=xlUp
                Else
                    If iEnd = .Range("A100000").End(xlUp).Row Then
                        .Rows(sheetStartRowNum(12) & ":" & iStart - 1).Delete Shift:=xlUp
                    Else
                        .Rows(iEnd + 1 & ":" & .Rows.Count).Delete Shift:=xlUp
                        .Rows(sheetStartRowNum(12) & ":" & iStart - 1).Delete Shift:=xlUp
                    End If
                End If
            Else
                .Rows(sheetStartRowNum(12) & ":" & .Rows.Count).Delete Shift:=xlUp
            End If
        End With
End Sub
Function getSaveFileName(StrPt As String, StrCode As String, StrName As String) As String
    getSaveFileName = StrPt & "\" & StrCode & "-" & StrName
End Function
Function getFilePath(str) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = str
        If .Show = -1 Then
            getFilePath = .SelectedItems(1)
        Else
            MsgBox "您没有选择任何文件夹！", vbCritical, "Error:"
            getFilePath = ""
            Exit Function
        End If
    End With
End Function
Sub splitSheets(sheetIndex As Integer, companyCode)
Dim sheetStartRowNum As Object
Dim CompanyRange As Range

Set sheetStartRowNum = createDelaySheetStartRowDict()

            Set ws = wb.Sheets(sheetIndex)
            ws.UsedRange.AutoFilter
            Set CompanyRange = ws.Range("A:A").Find(companyCode, , , , , xlNext): CompanyRangeStartRowIndex = CompanyRange.Row: Set CompanyRange = Nothing
            If Len(CompanyRangeStartRowIndex) < 1 Then ws.Rows(sheetStartRowNum(sheetIndex) & ":" & ws.Rows.Count).Delete Shift:=xlUp: Exit Sub
            Set CompanyRange = ws.Range("A:A").Find(companyCode, , , , , xlPrevious): CompanyRangeEndtRowIndex = CompanyRange.Row: Set CompanyRange = Nothing
            CompanyRangeStartRowIndex = getCompanyRangeStartRowIndex(sheetStartRowNum(sheetIndex), CInt(CompanyRangeStartRowIndex))
            CompanyRangeLastRowIndex = getCompanyRangeLastRowIndex(ws.Rows.Count, CInt(CompanyRangeEndtRowIndex))
            Union(ws.Range(sheetStartRowNum(sheetIndex) & ":" & CompanyRangeStartRowIndex), ws.Range(CompanyRangeLastRowIndex & ":" & ws.Rows.Count)).Delete
'            ws.Rows(sheetStartRowNum(sheetIndex) & ":" & CompanyRangeStartRowIndex).Delete Shift:=xlUp
'            ws.Rows(CompanyRangeLastRowIndex & ":" & ws.Rows.Count).Delete Shift:=xlUp
End Sub
Sub delayBilling()
 Dim CompanyRange As Range: Dim sheetStartRowNum, startRowNum As Object
 Dim filePath As String
companyCodeArray = Sheet5.Range("L:L").Value
filePath = getFilePath("请选择递延账单分析文件夹!"): Set fso = CreateObject("Scripting.FileSystemObject"): Set oFolder = fso.getFolder(filePath): Set oFile = oFolder.Files
For Each myfile In oFile



Set wb = Workbooks.Open(myfile): TurnEverythingOff
    For Each companyCode In companyCodeArray
        
        Call splitSheets(2, companyCode): Call splitSheets(3, companyCode)
        wb.SaveCopyAs getSaveFileName(filePath, CStr(companyCode), wb.name)
    Next
Next
TurnEverythingOn: MsgBox "Done!"
End Sub
Sub a(name)
Set wb = Workbooks.Open(name): TurnEverythingOff
    newName = Replace(name, "xlsb", "xlt")
    wb.Close True, Filename:=newName
End Sub
Sub filterBilling()
Dim iSht, startRow As Integer
Dim startRowNum
Set startRowNum = CreateObject("Scripting.Dictionary")
iName = InStr(1, ThisWorkbook.name, ".")
StrName = Left(ThisWorkbook.name, iName - 1)
    MsgBox "请选择筛选预收款文件夹！", vbInformation, "选择文件："
    iLoopEnd = Sheet5.Range("N65536").End(xlUp).Row
With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "请选择文件夹！"
    If .Show = -1 Then
        StrPt = .SelectedItems(1)
    Else
        MsgBox "您没有选择任何文件夹！", vbCritical, "Error:"
        
    End If
End With
Set fso = CreateObject("Scripting.FileSystemObject")
Set oFolder = fso.getFolder(StrPt)
Set oFile = oFolder.Files
companyCodeArray = Sheet5.Range("N" & iLoop).Value
For Each myfile In oFile
For iLoop = 1 To iLoopEnd
    Set wb = Workbooks.Open(myfile): TurnEverythingOff

    iName = InStr(1, wb.name, ".")
    StrName = Left(wb.name, iName - 1)
    If iFlag <> 12 Then
        
        For iSht = 2 To iFlag
            Set ws = wb.Sheets(iSht)
            With ws.UsedRange.AutoFilter
            Set CompanyRange = ws.Range("A:A").Find(companyCode)
                If Not rng Is Nothing Then
                    CompanyRangeStartRowIndex = getCompanyRangeStartRowIndex(sheetStartRowNum(sheetIndex))
                    CompanyRangeLastRowIndex = getCompanyRangeLastRowIndex(ws.Rows.Count)
                    ws.Rows(CompanyRangeStartRowIndex & ":" & CompanyRangeLastRowIndex).Delete Shift:=xlUp
                Else
                    ws.Rows(startRowNum(iSht) & ":" & ws.Rows.Count).Delete Shift:=xlUp
                End If
            Set ws = Nothing
        Next iSht
    End If
    wb.Close True, StrPt & "\" & StrCode & "-" & StrName
    Set ws = Nothing
    Set wb = Nothing
Arrow = "-" + Arrow

Application.StatusBar = "Progress: " & Arrow & Format(iLoop / iLoopEnd, "0%")
Next iLoop
Next
TurnEverythingOn
MsgBox "Done!"
End Sub
Sub countBilling()

Dim sheetStartRowNum As Object
Dim StrCode
Dim filePath As String

LastRowIndex = Sheet5.Range("P65536").End(xlUp).Row
companyCodeArray = Sheet5.Range("P:P").Value
Set sheetStartRowNum = createCalculateTaxSheetStartRowDict()
filePath = getFilePath
If filePath = "" Then Exit Sub
Set fso = CreateObject("Scripting.FileSystemObject")
Set oFile = fso.getFolder(filePath).Files
For Each file In oFile
    For Each companyCode In companyCodeArray
        Set wb = Workbooks.Open(file): TurnEverythingOff
        dealWithCalculateTaxSheet3ToSheet6 sheetStartRowNum, (companyCode)
        dealWithCalculateTaxSheet7ToSheet10 sheetStartRowNum, (companyCode)
        dealWithCalculateTaxSheet12 sheetStartRowNum, (companyCode)
        wb.Close True, getSaveFileName(filePath, companyCode, Left(wb.name, InStr(1, wb.name, ".") - 1))
        Set ws = Nothing
        Set wb = Nothing
        Application.StatusBar = "Progress: " & "-" + Arrow & Format(RowIndex / LastRowIndex, "0%")
    Next
Next
TurnEverythingOn
MsgBox "Done!"

End Sub


Sub main()

Union(Range("1:2"), Range("5:6")).Delete


End Sub




