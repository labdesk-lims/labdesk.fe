Attribute VB_Name = "ManageReport"
'################################################################################################
' This module initializes all reports and translates all fields
'################################################################################################

Option Compare Database
Option Explicit

Public Function TranslateReport(ByRef rfrm As Report, ByVal language As String) As Boolean
'-------------------------------------------------------------------------------
' Function:          TranslateReport
' Date:              2022 September
' Purpose:           Translate title and labels in reports
' In:
' -> rfrm:           Reference to report
' -> language:       Language to translate (e.g. de, en)
' Out:               Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim ctrl As control
    Dim s As String
    
    rfrm.Caption = GetTranslation(rfrm.name, "caption_", language)
    
    For Each ctrl In rfrm.Controls
        'Translate labels
        If TypeName(ctrl) = "Label" Then
            s = GetTranslation(rfrm.name, ctrl.name, language)
            If s <> "" Then rfrm.Controls(ctrl.name).Caption = s
        End If
    Next ctrl
    
    TranslateReport = True
    
Exit_Function:
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function
