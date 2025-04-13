Attribute VB_Name = "ManageForm"
'################################################################################################
' This module initializes all forms and sets the permissions of all controls in form. According
' to the access vocabulary 'READ' is used instead of 'CRUD' notation to handle permissions.
'################################################################################################

Option Compare Database
Option Explicit

' Enumeration to handle the corporate identity coloration
Public Enum CI
    BackColor_TextBox = 15263976
    BackColor_ComboBox = 15263976
    BackColor_TabPage = 14474460
    ForeColor_TabPage = 0
    SubForm_BorderStyle = 1
End Enum

' Get the table name from a GUID
Public Function GetTableNameFromGuid(ByVal tmpGuid As String) As String
    GetTableNameFromGuid = "tmp_" & tmpGuid
End Function

' Get the name of a table copy from a GUID
Public Function GetCpyTableNameFromGuid(ByVal tmpGuid As String) As String
    GetCpyTableNameFromGuid = "tmp_" & tmpGuid & "_cpy"
End Function

Public Function FormExist(strName As String) As Boolean
'-------------------------------------------------------------------------------
' Function:          FormExist
' Date:              2023 December
' Purpose:           Check if form is existent
' In:
' -> strName:        Name of form
' Out:               T/F
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim frm As Access.AccessObject

    For Each frm In Application.CurrentProject.AllForms
        If strName = frm.name Then
            FormExist = True
            Exit For    'We know it exist so let leave, no point continuing
        End If
    Next frm

Exit_Function:
    Exit Function
Catch_Error:
    AddErrorLog Err.Number, "ManageForm.FormExists: " & Err.description
    Resume Exit_Function
End Function

Function IsFormOpen(sFrmName As String) As Boolean
'-------------------------------------------------------------------------------
' Function:          IsFormOpen
' Date:              2022 February
' Purpose:           Check if form is open
' In:
' -> sFrmName:       Name of the form
' Out:               T/F
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
 
    IsFormOpen = Application.CurrentProject.AllForms(sFrmName).IsLoaded
 
Exit_Function:
    Exit Function
Catch_Error:
    AddErrorLog Err.Number, "ManageForm.IsFormOpen: " & Err.description
    Resume Exit_Function
End Function

Public Function TranslateForm(ByRef rfrm As Form, ByVal language As String) As Boolean
'-------------------------------------------------------------------------------
' Function:          TranslateForm
' Date:              2021 October
' Purpose:           Translate title and labels in form
' In:
' -> rfrm:           Reference to form
' -> language:       Language to translate (e.g. de, en)
' Out:               Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim ctrl As control
    Dim i As Integer
    Dim s As String
    
    rfrm.caption = GetTranslation(rfrm.name, "caption_", language)
    
    For Each ctrl In rfrm.Controls
        'Translate labels
        If TypeName(ctrl) = "Label" Then
            'labels
            s = GetTranslation(rfrm.name, ctrl.name, language)
            If s <> "" Then rfrm.Controls(ctrl.name).caption = s
            'tooltips
            s = GetTranslation(rfrm.name, ctrl.name & "tooltip", language)
            If s <> "" And s <> ctrl.name & "tooltip" Then rfrm.Controls(ctrl.name).ControlTipText = s
        End If
        
        'Translate buttons
        If TypeName(ctrl) = "CommandButton" Then
            s = GetTranslation(rfrm.name, ctrl.name, language)
            If s <> "" Then rfrm.Controls(ctrl.name).caption = s
        End If
        
        'Translate register tabs
        If TypeName(ctrl) = "TabControl" Then
            While i < rfrm.Controls(ctrl.name).Pages.count
                s = GetTranslation(rfrm.name, rfrm.Controls(ctrl.name).Pages(i).name, language)
                If s <> "" Then rfrm.Controls(ctrl.name).Pages(i).caption = s
                i = i + 1
            Wend
            i = 0
        End If
    Next ctrl
    
    TranslateForm = True
    
Exit_Function:
    Exit Function
Catch_Error:
    AddErrorLog Err.Number, "ManageForm.TranslateForm: " & Err.description
    Resume Exit_Function
End Function

Public Function CustomizeForm(ByRef rfrm As Form) As Boolean
'-------------------------------------------------------------------------------
' Function:          CustomizeForm
' Date:              2021 October
' Purpose:           Customize the form design
' In:
' -> rfrm:           Reference to form
' Out:               Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim ctrl As control
    Dim i As Integer
    
    'Config form design
    rfrm.RecordSelectors = False
    rfrm.NavigationButtons = False
    
    'Config control design
    For Each ctrl In rfrm.Controls
        'Customize register tabs
        If TypeName(ctrl) = "TabControl" Then
            While i < rfrm.Controls(ctrl.name).Pages.count
                rfrm.Controls(ctrl.name).BorderStyle = 0 'Transparent
                rfrm.Controls(ctrl.name).ForeColor = CI.ForeColor_TabPage
                rfrm.Controls(ctrl.name).BackColor = CI.BackColor_TabPage
                rfrm.Controls(ctrl.name).style = 0
                i = i + 1
            Wend
        End If
        
        'Customize Labels
        If TypeName(ctrl) = "Label" Then
            rfrm.Controls(ctrl.name).FontItalic = True
        End If
        
        'Customize TextBox
        If TypeName(ctrl) = "TextBox" Then
            rfrm.Controls(ctrl.name).BackColor = CI.BackColor_TextBox
            rfrm.Controls(ctrl.name).BorderStyle = 0
        End If
        
        'Customize ComboBox
        If TypeName(ctrl) = "ComboBox" Then
            rfrm.Controls(ctrl.name).BackColor = CI.BackColor_ComboBox
        End If
        
        'Customize SubForms
        If TypeName(ctrl) = "SubForm" Then
            rfrm.Controls(ctrl.name).BorderStyle = CI.SubForm_BorderStyle
        End If
    Next ctrl
    
    CustomizeForm = True
    
Exit_Function:
    Exit Function
Catch_Error:
    AddErrorLog Err.Number, "ManageForm.CustomizeForm: " & Err.description
    Resume Exit_Function
End Function

Public Function ConfigAsyncForm(ByRef rfrm As Form, ByVal Table As String, ByVal tmpGuid As String, Optional navButton As Boolean, Optional readOnly As Boolean, Optional whereClause As String) As Boolean
'-------------------------------------------------------------------------------
' Function:         ConfigAsyncForm
' Date:             2022 January
' Purpose:          Configure an asychronous form
' In:
' -> rfrm:          Form which will be configured
' -> table:         Name of the table the form is linked with
' -> tmpGuid:       GUID of the temporary table for async data handling
' -> navButton:     Show navigation buttons on form (T/F)
' -> readOnly:      Set the form to read_only
' -> whereClause:   Where clause to filter recordset
' Out:              Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    'Check and set form permissions
    If Not GetPermission(rfrm.name).Read Then Err.Raise vbObjectError + 513, , GetTranslation("mdlManageForm", "msgbox_permission_denied", GetDbSetting("language"))
    If Not GetPermission(rfrm.name).Update Or readOnly Then rfrm.AllowEdits = False Else rfrm.AllowEdits = True
    If Not GetPermission(rfrm.name).Create Or (readOnly And Not IsFormView(rfrm)) Then rfrm.AllowAdditions = False Else rfrm.AllowAdditions = True
    If Not GetPermission(rfrm.name).Delete Or readOnly Then rfrm.AllowDeletions = False Else rfrm.AllowDeletions = True
    
    'Customize design
    CustomizeForm rfrm

    'Disable identity column if one exists
    If GetIdentityColumn(Table) <> "" Then rfrm.Controls(GetIdentityColumn(Table)).enabled = False
    
    'Translate all controls
    TranslateForm rfrm.Form, GetDbSetting("language")
    If Not GetPermission(rfrm.name).Update Or readOnly Then rfrm.caption = rfrm.caption & " " & GetTranslation("form", "label_read_only", DbConnect.GetDbSetting("language"))
    
    'Prepare temporary table for actual session and set it as object source
    If TableExists(GetTableNameFromGuid(tmpGuid)) Then CurrentDb.TableDefs.Delete GetTableNameFromGuid(tmpGuid)
    CreateTable Table, GetTableNameFromGuid(tmpGuid)
    rfrm.RecordSource = GetTableNameFromGuid(tmpGuid)
    
    'Copy table for async data handling
    If whereClause <> "" Then CopyTable Table, GetTableNameFromGuid(tmpGuid), whereClause
    If TableExists(GetCpyTableNameFromGuid(tmpGuid)) Then CurrentDb.TableDefs.Delete GetCpyTableNameFromGuid(tmpGuid)
    CopyTable GetTableNameFromGuid(tmpGuid), GetCpyTableNameFromGuid(tmpGuid)
    
    rfrm.Requery

    ConfigAsyncForm = True
    
Exit_Function:
    Exit Function
Catch_Error:
    AddErrorLog Err.Number, "ManageForm.ConfigAsynForm: " & Err.description
    Resume Exit_Function
End Function

Public Function ConfigAsyncSubForm(ByRef rfrm As Form, ByVal Table As String, ByVal tmpGuid As String, Optional ByVal navButton As Boolean, Optional ByVal readOnly As Boolean, Optional ByVal whereClause As String) As Boolean
'-------------------------------------------------------------------------------
' Function:         ConfigAsyncSubForm
' Date:             2022 January
' Purpose:          Configure an asychronous sub form
' In:
' -> rfrm:          Form which will be configured
' -> table:         Name of the table the form is linked with
' -> tmpGuid:       GUID of the temporary table for async data handling
' -> navButton:     Show navigation buttons on form (T/F)
' -> readOnly:      Set the form to read_only
' -> whereClause:   Where clause to filter recordset
' Out:              Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    
    'Check and set form permissions
    If Not GetPermission(rfrm.name).Read Then
        rfrm.visible = False
        rfrm.AllowAdditions = False
        Err.Raise vbObjectError + 513, , GetTranslation("mdlManageForm", "msgbox_permission_denied", GetDbSetting("language"))
    End If
    If Not GetPermission(rfrm.name).Update Or readOnly Then rfrm.AllowEdits = False Else rfrm.AllowEdits = True
    If Not GetPermission(rfrm.name).Create Or readOnly Then rfrm.AllowAdditions = False Else rfrm.AllowAdditions = True
    If Not GetPermission(rfrm.name).Delete Or readOnly Then rfrm.AllowDeletions = False Else rfrm.AllowDeletions = True
    
    'Customize design
    CustomizeForm rfrm

    'Disable identity column
    rfrm.Controls(GetIdentityColumn(Table)).enabled = False

    'Translate all controls
    TranslateForm rfrm.Form, GetDbSetting("language")
    
    'Prepare temporary table for actual session and set it as object source
    If TableExists(GetTableNameFromGuid(tmpGuid)) Then CurrentDb.TableDefs.Delete GetTableNameFromGuid(tmpGuid)
    CreateTable Table, GetTableNameFromGuid(tmpGuid)
    rfrm.RecordSource = GetTableNameFromGuid(tmpGuid)
    
    'Copy table for async data handling
    If whereClause <> "" Then CopyTable Table, GetTableNameFromGuid(tmpGuid), whereClause
    If TableExists(GetCpyTableNameFromGuid(tmpGuid)) Then CurrentDb.TableDefs.Delete GetCpyTableNameFromGuid(tmpGuid)
    CopyTable GetTableNameFromGuid(tmpGuid), GetCpyTableNameFromGuid(tmpGuid)
    rfrm.Requery
    
    ConfigAsyncSubForm = True

Exit_Function:
    Exit Function
Catch_Error:
    AddErrorLog Err.Number, "ManageForm.ConfigAsyncSubForm: " & Err.description
    Resume Exit_Function
End Function

Public Function ConfigSyncForm(ByRef rfrm As Form, ByVal Table As String, Optional ByVal navButton As Boolean, Optional ByVal readOnly As Boolean) As Boolean
'-------------------------------------------------------------------------------
' Function:         ConfigSyncForm
' Date:             2022 January
' Purpose:          Configure a sychronous form
' In:
' -> rfrm:          Form which will be configured
' -> table:         Name of the table the form is linked with
' -> tmpGuid:       GUID of the temporary table for async data handling
' -> navButton:     Show navigation buttons on form (T/F)
' -> readOnly:      Set the form to read_only
' Out:              Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim ctrl As control
    
    'Check and set form permissions
    If Not GetPermission(rfrm.name).Read Then Err.Raise vbObjectError + 513, , GetTranslation("mdlManageForm", "msgbox_permission_denied", GetDbSetting("language"))
    If Not GetPermission(rfrm.name).Update Or readOnly Then rfrm.AllowEdits = False Else rfrm.AllowEdits = True
    If Not GetPermission(rfrm.name).Create Or readOnly Then rfrm.AllowAdditions = False Else rfrm.AllowAdditions = True
    If Not GetPermission(rfrm.name).Delete Or readOnly Then rfrm.AllowDeletions = False Else rfrm.AllowDeletions = True
    
    'Set all labels to read only if edit permission is not set
    For Each ctrl In rfrm.Controls
        'Translate labels
        If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "ComboBox" Or TypeName(ctrl) = "CustomControl" Or TypeName(ctrl) = "ListBox" Or TypeName(ctrl) = "ObjecFrame" Then
            If Not GetPermission(rfrm.name).Update Or readOnly Then
                rfrm.Controls(ctrl.name).Locked = True
            Else
                rfrm.Controls(ctrl.name).Locked = False
            End If
        End If
    Next ctrl
    
    'Customize design
    CustomizeForm rfrm
    
    'Translate all controls
    TranslateForm rfrm.Form, GetDbSetting("language")
    
    'Attach a read only hint if applies
    If Not GetPermission(rfrm.name).Update Or readOnly Then rfrm.caption = rfrm.caption & " " & GetTranslation("form", "label_read_only", GetDbSetting("language"))
    
    ConfigSyncForm = True
    
Exit_Function:
    Exit Function
Catch_Error:
    AddErrorLog Err.Number, "ManageForm.ConfigSyncForm: " & Err.description
    Resume Exit_Function
End Function

Public Sub applyMDIMode()
'-------------------------------------------------------------------------------
' Function:  UseMDIMode
' Date:      2022 November
' Purpose:   Set the multi document interface mode (MDI)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error

    If CurrentDb.Properties("UseMDIMode").value Then
        CurrentDb.Properties("UseMDIMode").value = 0
        MsgBox GetTranslation("system", "MDIMode-Off", GetDbSetting("language")), vbInformation, GetTranslation("msgbox", "vbInformation", GetDbSetting("language"))
    Else
        CurrentDb.Properties("UseMDIMode").value = 1
        MsgBox GetTranslation("system", "MDIMode-On", GetDbSetting("language")), vbInformation, GetTranslation("msgbox", "vbInformation", GetDbSetting("language"))
    End If
    
Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageForm.UseMDIMode: " & Err.description
    Resume Exit_Function
End Sub
