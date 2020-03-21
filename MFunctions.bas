Attribute VB_Name = "MFunctions"
Option Explicit

Sub CreateShortcut()

    Application.OnKey "+^{U}"
    Application.OnKey "+^{U}", "Matching"
    Application.OnKey "+^{M}", "CtrlShiftM"
    
End Sub

Sub CtrlShiftM()
    
    If ActiveSheet.Name Like prefix & "*" Then
        Range("FRVesrion").Value = IIf(Range("FRVesrion").Value = "Y", "N", "Y")
    Else
        Call InsertSamples
    End If
    
End Sub
Function bWksExist(sWksName As String, Optional wbk As Workbook) As Boolean

    bWksExist = False
    
    If wbk Is Nothing Then Set wbk = ThisWorkbook
    
    On Error Resume Next
    
    bWksExist = wbk.Sheets(sWksName).Index > 0
    
End Function

Function AppSettings(b As Boolean, cal As XlCalculation)

    With Application
        .ScreenUpdating = b
        .DisplayAlerts = b
        .EnableEvents = b
        .Calculation = cal
    End With
    
End Function

Function AddWksEnd(wbk As Workbook, Optional wks As Worksheet) As Worksheet

    If wks Is Nothing Then
    
        Set AddWksEnd = wbk.Sheets.Add(After:=wbk.Sheets(wbk.Sheets.Count))
        
    Else
    
        wks.Copy before:=ThisWorkbook.Sheets(1)
        Set AddWksEnd = wbk.Sheets(1)
        AddWksEnd.Move After:=wbk.Sheets(wbk.Sheets.Count)
        
    End If
    
End Function


Function FindRow(sWhat As String, rngWhere As Range) As Integer
    
    FindRow = 0
    
    On Error Resume Next
    
    FindRow = rngWhere.Find(What:=sWhat, LookIn:=xlValues, LookAt:=xlPart, _
                                    SearchOrder:=xlByRows).Row
    
End Function

Function dFindALL(sWhat As String, dNumber As Double, rWhere As Range, ByRef sIfMatched As String) As Dictionary

    Dim sNewWhat As String: sNewWhat = sWhat
    Dim iRow As Integer
    Dim sFoundItem As String
    Dim i As Integer
    
    Set dFindALL = New Dictionary
    
    iRow = FindRow(sNewWhat, rWhere)
    
    If iRow = 0 Then
        sNewWhat = Replace(LCase(sNewWhat), "0", "o")
        iRow = FindRow(sNewWhat, rWhere)
    End If
    
    If iRow = 0 Then
        sNewWhat = Replace(LCase(sWhat), "0", "*")
        iRow = FindRow(sNewWhat, rWhere)
    End If
    
    If iRow = 0 Then Exit Function
    
    sFoundItem = rWhere.Find(What:=sNewWhat, LookIn:=xlValues, LookAt:=xlPart, _
                                    SearchOrder:=xlByRows).Address
                                    
    Do
    
        rWhere.Parent.Range(sFoundItem).Interior.ColorIndex = 6
        dFindALL.Add Key:=sFoundItem, Item:=0
        
        For i = 1 To rWhere.Columns.Count
            If rWhere.Parent.Cells(Range(sFoundItem).Row, i).Value = dNumber Then
                
                dFindALL.Remove (sFoundItem)
                dFindALL.Add Key:=sFoundItem, Item:=Cells(Range(sFoundItem).Row, i).Address
                rWhere.Parent.Cells(Range(sFoundItem).Row, i).Interior.ColorIndex = 4

                If sIfMatched = "" Then sIfMatched = sFoundItem
                
            End If
        Next i
        
        On Error Resume Next
        
            sFoundItem = rWhere.FindNext.Address
        
        On Error GoTo 0
    Loop While Not dFindALL.Exists(sFoundItem)
    
End Function


Function FilePicker(capt As String) As Workbook
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = capt
        .InitialFileName = ThisWorkbook.Path & "\"
        .AllowMultiSelect = False
        If .Show <> -1 Then GoTo No_file
        Set FilePicker = Workbooks.Open(Dir(.SelectedItems(1)))
    End With
    
    Exit Function
No_file:
    MsgBox "No file chosen!", , "File open error!"
End Function

