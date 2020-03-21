Attribute VB_Name = "MFinalize"
Option Explicit

Const sNotFoundComment As String = "ISIN number not found within depositary confirmation."

Sub FinalizeRequest()

    Dim bCal As XlCalculation
    Dim wResults As Workbook
    Dim wTemplate As Workbook
    Dim wks As Worksheet
    Dim sToMove() As String
    
    bCal = Application.Calculation
    
    Call AppSettings(False, xlCalculationManual)
    
    Set wTemplate = ThisWorkbook

    If wTemplate.Sheets.Count < 3 Then
        MsgBox "There is no sheet to move into SDC Results file.", vbInformation, "Information"
        GoTo ExitMacro
    End If
    
    For Each wks In wTemplate.Sheets
        If wks.Name <> "Template" And wks.Name <> "Lista Funduszy" And wks.Name <> "Info" Then
            On Error Resume Next
            
                ReDim Preserve sToMove(UBound(sToMove) + 1)
                If Err.Number > 0 Then ReDim Preserve sToMove(0)
                
            On Error GoTo 0
            
            sToMove(UBound(sToMove)) = "" & wks.Name & ""
            
        End If
    Next wks
    
    Set wResults = Workbooks.Add
    
    wTemplate.Sheets(sToMove).Copy After:=wResults.Sheets(1)
     
    wResults.Sheets(1).Delete
    
    For Each wks In wResults.Sheets
    
        If Not (wks.Name Like prefix & "*") Then
        
            Call TemplateEditing(wks)
            
        Else
            
            wks.UsedRange.ClearComments
            
        End If
    Next wks
    
    
    'gdyby plik juz istnial to zostawi otwary i zakonczy prace
    On Error GoTo ExitMacro
    wResults.SaveAs ThisWorkbook.Path & "\" & Replace(ThisWorkbook.Name, ".xlsm", "") & " SDC Results.xlsx"
    wResults.Close
    
ExitMacro:
    Call AppSettings(True, bCal)
End Sub

Sub TemplateEditing(wks As Worksheet)
    Dim iFirstRow As Integer: iFirstRow = 5
    Dim iLastRow As Integer, i As Integer


    With wks
    
        'Gdyby klient chcial cos wiecej to zmienimy tutaj
        .Range("S:W").Delete
        
        iLastRow = .Range("L1000000").End(xlUp).Row
        
        If iFirstRow > iLastRow Then Exit Sub
        
        For i = iFirstRow To iLastRow
        
            If .Range("L" & i).Value = "N" And .Range("M" & i).Value = "" Then
            
                With .Range("M" & i & ":O" & i)
                    .Value = "o/s"
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Interior.Color = RGB(252, 228, 214)
                    
                        With .Font
                            .Name = "Georgia"
                            .Size = 9
                            .Italic = False
                            .Bold = True
                            .Underline = False
                        End With

                End With
                
                .Range("P" & i).Value = sNotFoundComment
                
            ElseIf .Range("P" & i).Value = "" Then
            
                .Range("P" & i).Value = "'-"
                
            End If
        Next i
        
    End With
    
End Sub
