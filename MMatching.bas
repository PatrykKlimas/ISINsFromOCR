Attribute VB_Name = "MMatching"
Option Explicit

Const sISINCol As String = "H"
Const sQuanCol As String = "I"
Const sToFill As String = "M"

'Pierwsza czesc to walidacja danych

Sub Matching()
    
    Dim bCal As XlCalculation
    Dim wksTemplate As Worksheet
    Dim wksOCR As Worksheet
    
    Set wksTemplate = ActiveSheet

    bCal = Application.Calculation
    
    Call MFunctions.AppSettings(False, xlCalculationManual)
    
    If Not (MFunctions.bWksExist(prefix & wksTemplate.Name, wksTemplate.Parent)) Then
        
        MsgBox "Worksheet """ & prefix & wksTemplate.Name & """ does not exist!", vbCritical, "Error!"
        GoTo EndMacro
    
    End If
    
    Set wksOCR = wksTemplate.Parent.Sheets(prefix & wksTemplate.Name)
    
    If wksOCR.UsedRange.Cells.Count < 2 Then
    
        MsgBox "Please verify data in the worksheet """ & prefix & wksTemplate.Name & """.", vbInformation, "Information"
        GoTo EndMacro
    
    End If
    
    If WorksheetFunction.CountA(wksTemplate.Range(sISINCol & ":" & sISINCol)) < 3 Then
    
        MsgBox "No ISIN presented in the column H!", vbInformation, "Information"
        GoTo EndMacro
    End If
    
    Call MatchingLoop(wksTemplate, wksOCR)
    
EndMacro:
    Call MFunctions.AppSettings(True, bCal)
End Sub

Sub MatchingLoop(wksTemplate As Worksheet, wksOCR As Worksheet)
    
    Dim rRowTemplate As Integer: rRowTemplate = 5
    Dim iRow As Integer, i As Integer, iNumberOfMatches
    Dim sISIN As String, sNewISIN As String
    Dim dQuantity As Double
    Dim dISINFound As New Dictionary
    Dim sFounded As String, SMatchedQuantity As String
    
    'Petla po wszystkich ISINach
    While wksTemplate.Range(sISINCol & rRowTemplate).Value2 <> ""
        
        'sprawdzenie czy nie jest juz wypelnione
        If wksTemplate.Range("L" & rRowTemplate).Value2 <> "" Then GoTo NextIteration
        
        sISIN = wksTemplate.Range(sISINCol & rRowTemplate).Value
        dQuantity = wksTemplate.Range(sQuanCol & rRowTemplate).Value
        sNewISIN = sISIN
        
        'wyszukiwanie ISINow i kwot
        iRow = FindRow(sISIN, wksOCR.UsedRange)
        
        If iRow = 0 Then
            sNewISIN = Replace(LCase(sISIN), "0", "o")
            iRow = FindRow(sNewISIN, wksOCR.UsedRange)
        End If
        
        If iRow = 0 Then
            sNewISIN = Replace(LCase(sISIN), "0", "*")
            iRow = FindRow(sNewISIN, wksOCR.UsedRange)
        End If
        
        'Gdy nie znajdziemy ISIN'a to idziemy do czesci wypelniajacej Template
        If iRow = 0 Then GoTo FillTemplate
        
        wksTemplate.Range("V" & rRowTemplate).Value = UCase(sNewISIN)
        
        If sNewISIN <> sISIN Then wksTemplate.Range("V" & rRowTemplate).Interior.ColorIndex = 6
        
        sFounded = wksOCR.UsedRange.Find(What:=sNewISIN, LookIn:=xlValues, LookAt:=xlPart, _
                                        SearchOrder:=xlByRows).Address
                                        
        Do
        
            wksOCR.Range(sFounded).Interior.ColorIndex = 6
            dISINFound.Add Key:=sFounded, Item:=0
            
            'petla odpowiedzialna za znalezienie kwoty w wierszu z ISINem
            For i = 1 To wksOCR.UsedRange.Column + wksOCR.UsedRange.Columns.Count
            
                If wksOCR.Cells(Range(sFounded).Row, i).Value = dQuantity Then
                    
                    dISINFound(sFounded) = Cells(Range(sFounded).Row, i).Address
                    
                    wksOCR.Cells(Range(sFounded).Row, i).Interior.ColorIndex = 4
                    
                    'iNumberOfMatches zlicza liczbe ilosci w danym wierszu
                    If SMatchedQuantity = "" Then
                    
                        SMatchedQuantity = sFounded
                        iNumberOfMatches = 1
                        
                    ElseIf Range(SMatchedQuantity).Row = Range(sFounded).Row Then
                        
                        iNumberOfMatches = iNumberOfMatches + 1
                        
                    End If
                    
                End If
            Next i
            
            On Error Resume Next
                sFounded = wksOCR.UsedRange.FindNext.Address
            On Error GoTo 0
            
        Loop While Not dISINFound.Exists(sFounded)
    
FillTemplate:

        'czesc odpowiedzialna za wypelnienie templatu
        With wksTemplate
        
            .Range("S" & rRowTemplate).Value = dISINFound.Count
            
            If dISINFound.Count > 0 Then
            
                .Range("L" & rRowTemplate).Value = "Y"
                
                If SMatchedQuantity <> "" Then
                
                    'dodanie hyperlina - jesli w wierszu sa 2 kwoty to do wiersza
                    .Hyperlinks.Add Anchor:=Range("O" & rRowTemplate), Address:="", SubAddress:= _
                                            "'" & wksOCR.Name & "'!" & IIf(iNumberOfMatches > 1, Split(dISINFound(SMatchedQuantity), "$")(2) & ":" & Split(dISINFound(SMatchedQuantity), "$")(2), _
                                                dISINFound(SMatchedQuantity)), TextToDisplay:="Link"
                                            
                    .Range(sToFill & rRowTemplate).Value = "='" & wksOCR.Name & "'!" & dISINFound(SMatchedQuantity)
                    .Range("T" & rRowTemplate).Value = Range(SMatchedQuantity).Row
                    .Range("U" & rRowTemplate).Value = Range(dISINFound(SMatchedQuantity)).Column
                    .Range("W" & rRowTemplate).Value = iNumberOfMatches
                    .Range(sToFill & rRowTemplate).Interior.Color = RGB(252, 228, 214)
                    
                    
                Else
    
                    .Hyperlinks.Add Anchor:=Range("O" & rRowTemplate), Address:="", SubAddress:= _
                                            "'" & wksOCR.Name & "'!" & Split(dISINFound.Keys(0), "$")(2) & ":" & Split(dISINFound.Keys(0), "$")(2), TextToDisplay:="Link"
                            
                    .Range(sToFill & rRowTemplate).Interior.ColorIndex = 6
                    .Range("T" & rRowTemplate).Value = Range(dISINFound.Keys(0)).Row
                    
                End If
            Else
            
                .Range("L" & rRowTemplate).Value = "N"
                .Range(sToFill & rRowTemplate).Interior.ColorIndex = 6
                          
            End If
            
        End With
        
NextIteration:
        
        'czyszczenie zmiennych przed kolejnym ISINem
        rRowTemplate = rRowTemplate + 1
        dISINFound.RemoveAll
        sFounded = ""
        SMatchedQuantity = ""
        iNumberOfMatches = 0
    Wend
    
End Sub


'makro do sprawdzenia adresu
Sub ShowAddress()

    MsgBox Selection.Address, , "Address"
    
End Sub

