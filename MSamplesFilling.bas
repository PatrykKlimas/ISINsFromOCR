Attribute VB_Name = "MSamplesFilling"
Option Explicit

Sub InsertRows(wks As Worksheet, i_samples As Integer)
    
    If i_samples = 2 Then Exit Sub
    
    If i_samples = 1 Then
        wks.Rows("6:6").Delete
        Exit Sub
    End If
    
    Dim i_rowsToAdd As Integer
    
    i_rowsToAdd = i_samples - 3
    
    With wks
        .Rows("6:" & 6 + i_rowsToAdd).Insert Shift:=xlDown
        .Range("B6:B" & i_samples + 4).Formula = "=R[-1]C+1"
        .Range("N6:N" & i_samples + 4).Formula = "=RC[-5]-RC[-1]"
    End With
End Sub

Sub InsertSamples()
       
    Call AppSettings(False, xlCalculationManual)
    
    Dim iSamplesNumber As Integer
    
    If ActiveSheet.Range("B1").Value <> "Confirm investments, including commitments, margin and collateral" Then _
            GoTo ExitMacro
    
    On Error GoTo ExitMacro
        iSamplesNumber = InputBox("Podaj liczbe sampli: ", "Liczba sampli")
    On Error Resume Next
    
    Call InsertRows(ActiveSheet, iSamplesNumber)
    
    
ExitMacro:
    Call AppSettings(False, xlCalculationAutomatic)
End Sub


Sub SamplesCreateFromRowData()

    Dim wks_source As Worksheet, wks_Template As Worksheet
    Dim arr() As Variant
    Dim sColToFileter As String
    Dim rFundName As Range
    Dim iSamplesNumber As Integer, iLastRowInSource As Integer, i As Integer
    
    On Error GoTo EndMacro
        Set wks_source = FilePicker("Wybierz plik z samplami").Sheets(1)
    On Error Resume Next
        wks_source.ShowAllData
    On Error GoTo 0
    
    arr = ThisWorkbook.Sheets("Info").Range("D3:E7").Value
    sColToFileter = Mapping.Range("K6").Value & ":" & Mapping.Range("K6").Value
    
    Call MainMacro
    
    Call AppSettings(False, xlCalculationManual)
    
    'Ostani wiersz w kolumnie z funduszami
    iLastRowInSource = wks_source.Range(Mapping.Range("K6").Value & _
                wks_source.Range("A1").SpecialCells(xlCellTypeLastCell).Row + 10).End(xlUp).Row
    
    Set rFundName = Mapping.Range("B3")
    
    While rFundName.Value <> ""
    
        wks_source.Range("1:1").AutoFilter Field:=Range(sColToFileter).Column, Criteria1:=rFundName.Value
        
        iSamplesNumber = WorksheetFunction.CountA(wks_source.Range(sColToFileter).SpecialCells(xlCellTypeVisible)) - 1
        
        If iSamplesNumber = 0 Then GoTo NextIteration
        
        Set wks_Template = ThisWorkbook.Sheets(rFundName.Value)
        
        Call InsertRows(wks_Template, iSamplesNumber)
        
        For i = 1 To 5
                        
            wks_Template.Range(arr(i, 1) & "5:" & arr(i, 1) & 4 + iSamplesNumber).Value = _
                    wks_source.Range(arr(i, 2) & "2:" & arr(i, 2) & iLastRowInSource).SpecialCells(xlCellTypeVisible).Value
        
        Next i
        
NextIteration:
        Set rFundName = rFundName.Offset(1, 0)
        
        On Error Resume Next
            wks_source.ShowAllData
        On Error GoTo 0
        
    Wend
    
    wks_source.Parent.Close
EndMacro:
    Call AppSettings(True, xlCalculationAutomatic)
    
End Sub
