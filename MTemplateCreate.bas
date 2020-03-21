Attribute VB_Name = "MTemplateCreate"
Option Explicit

Public Const prefix As String = "OCR_"

Sub MainMacro()

    Dim wksListing As Worksheet
    Dim wksTemplate As Worksheet

    Call MFunctions.AppSettings(False, xlCalculationManual)
    
    
    If Not (MFunctions.bWksExist("Lista Funduszy")) Then
        MsgBox "Worksheet ""Lista Funduszy"" does not exist!", vbCritical, "Error"
        GoTo EndMacro
    End If
    
    If Not (MFunctions.bWksExist("Template")) Then
        MsgBox "Worksheet ""Template"" does not exist!", vbCritical, "Error"
        GoTo EndMacro
    End If
    
    Set wksListing = ThisWorkbook.Sheets("Lista Funduszy")
    Set wksTemplate = ThisWorkbook.Sheets("Template")
    
    If wksListing.Range("FirstFundNr").Value2 = "" Then
        MsgBox "No fund has been entered or cell B3 is empty!", vbInformation, "Information"
        GoTo EndMacro
    End If
    
    Call TemplateCreator(wksListing, wksTemplate)
    
    wksListing.Activate
    
EndMacro:
    Call MFunctions.AppSettings(True, xlCalculationAutomatic)
   
End Sub

Sub TemplateCreator(wksListing As Worksheet, ByRef wksTemplate As Worksheet)

    Dim dFunds As New Dictionary
    Dim rFundNr As Range
    Dim wksFundTemplate As Worksheet
    Dim vkey As Variant
    
    Set rFundNr = wksListing.Range("FirstFundNr")
    
    While rFundNr.Value2 <> ""
    
        If dFunds.Exists(rFundNr.Value) Then
        
            MsgBox "Please check duplicates values in column B.", vbInformation, "Recheck"
            Exit Sub
            
        End If
        
        If Not MFunctions.bWksExist(rFundNr.Value, wksTemplate.Parent) Then
        
            dFunds.Add rFundNr.Value, rFundNr.Offset(0, -1).Value
            
        End If
        
        Set rFundNr = rFundNr.Offset(1, 0)
   
    Wend
    
    
    For Each vkey In dFunds.Keys
    
        Set wksFundTemplate = MFunctions.AddWksEnd(wksTemplate.Parent, wksTemplate)

        wksFundTemplate.Name = CStr(vkey)
        
        Set wksFundTemplate = MFunctions.AddWksEnd(wksTemplate.Parent)
        
        wksFundTemplate.Name = prefix & CStr(vkey)
        
    Next vkey
    
End Sub
