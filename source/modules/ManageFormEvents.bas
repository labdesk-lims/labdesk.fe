Attribute VB_Name = "ManageFormEvents"
'################################################################################################
' This module manages events from forms and will be called from there
'################################################################################################

Option Compare Database
Option Explicit

' -----------------------------------------------------------------------------------------------
' Asynchronous Form Events are configured from here on
' -----------------------------------------------------------------------------------------------

Public Function AsyncFormGetChanged(ByRef rfrm As Form) As Boolean
'-------------------------------------------------------------------------------
' Function:         AsyncFormGetChanged
' Date:             2022 March
' Purpose:          Check if data is changed on form
' In:               Form of choice to be checked
' -> rfrm:
' Out:              Changed (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim control As control
    Dim subFormChanged As Boolean
    
    ' Checks for changes on subforms as well
    For Each control In rfrm.Controls
        If control.ControlType = acSubform Then
            subFormChanged = subFormChanged Or rfrm.Form(control.name).Form.GetChanged(rfrm.Form(control.name).Form)
        End If
    Next
    
    AsyncFormGetChanged = rfrm.FormChanged Or subFormChanged

Exit_Function:
    Exit Function
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.AsyncFormGetChanged: " & Err.description
    Resume Exit_Function
End Function

Public Sub AsyncFormInit(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         AsyncFormInit
' Date:             2022 March
' Purpose:          Initialize an asynchronous form
' In:
' -> rfrm:          Form of choice to be initialized
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim readOnly As Boolean
    Dim ctrl As control

    ' Hide form during adjustment are made
    rfrm.visible = False
    rfrm.tmpGuid = CreateGuid()
    
    'Skip table flag check if new record is created
    If IsNull(rfrm.ID) Then GoTo Skip_Check
    
    ' Check if requested record is already in use by another user
    If Not IsNull(GetTableFlag(rfrm.DataTable, rfrm.ID)) Then
        If MsgBox(GetTranslation("msgbox", "tableflag_set", GetDbSetting("language")), vbYesNo, GetTranslation("msgbox", "vbInformation", GetDbSetting("language")) & " - " & GetTableFlag(rfrm.DataTable, rfrm.ID)) = vbYes Then
            RemoveTableFlag rfrm.DataTable, rfrm.ID
        End If
    End If
    
Skip_Check:
    
    ' Check read only condidtion
    If Not GetPermission(rfrm.name).Update Or rfrm.Form_Init_Individual_Permission() Or Not IsNull(GetTableFlag(rfrm.DataTable, Nz(rfrm.ID, 0))) Then
        readOnly = True
    Else
        ' Block record as in use by actual user
        If Not IsNull(rfrm.ID) Then SetTableFlag GetUserName, rfrm.DataTable, rfrm.ID
    End If
    
    ' Config main form
    If ConfigAsyncForm(rfrm.Form, rfrm.DataTable, rfrm.tmpGuid, , readOnly, rfrm.whereClause) = False Then
        DoCmd.Close acForm, rfrm.name
        Exit Sub
    End If

    ' Config subforms if ID is not null
    If Not IsNull(rfrm.ID) Then
        For Each ctrl In rfrm.Controls
            If ctrl.ControlType = acSubform Then
                rfrm.Form(ctrl.name).Form.Form_Init readOnly, "WHERE " & rfrm.DataTable & " = " & rfrm.ID
            End If
        Next
    Else
        For Each ctrl In rfrm.Controls
            If ctrl.ControlType = acSubform Then
                rfrm.Form(ctrl.name).Form.visible = False
            End If
        Next
    End If

    ' Adjust to best fit width
    For Each ctrl In rfrm.Controls
        If (ctrl.ControlType = acTextBox Or ctrl.ControlType = acComboBox) Then
            ctrl.ColumnWidth = -2
        End If
    Next ctrl
    
    ' Adjust to best fit hight
    rfrm.RowHeight = 300
    
    ' Init form with individual settings
    rfrm.Form_Init_Individual
    
    rfrm.FormChanged = False
    
    ' Show form after adjustments are made
    rfrm.visible = True

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.AsyncFormInit: " & Err.description
    Resume Exit_Function
End Sub

Public Sub AsyncFormOkClick(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         AsyncFormOkClick
' Date:             2022 March
' Purpose:          Save data of table if ok was clicked
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    If rfrm.FormChanged Then
        rfrm.Requery
        UpdateTable rfrm.DataTable, GetTableNameFromGuid(rfrm.tmpGuid)
        rfrm.FormChanged = False
    End If

    ' Close form
    rfrm.ForceClose = True
    DoCmd.Close acForm, rfrm.name, acSaveNo

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.AsyncFOrmOkClick: " & Err.description
    Resume Exit_Function
End Sub

Public Sub AsyncFormCancelClick(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         AsyncFormCancelClick
' Date:             2022 March
' Purpose:          Undo all changes if cancel is clicked
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim selectedRecord As Long
    Dim ctrl As control
    
    'Store selected record
    selectedRecord = rfrm.SelTop
    
    ' Undo all changes on main form
    If rfrm.FormChanged Then
        CopyTable rfrm.DataTable, GetTableNameFromGuid(rfrm.tmpGuid), rfrm.whereClause
        If TableExists(GetCpyTableNameFromGuid(rfrm.tmpGuid)) Then CurrentDb.TableDefs.Delete GetCpyTableNameFromGuid(rfrm.tmpGuid)
        CopyTable GetTableNameFromGuid(rfrm.tmpGuid), GetCpyTableNameFromGuid(rfrm.tmpGuid)
        rfrm.Undo
        rfrm.Requery
        rfrm.FormChanged = False
    End If

    'Undo all changes in subforms
    For Each ctrl In rfrm.Controls
        If ctrl.ControlType = acSubform Then
            rfrm.Form(ctrl.name).Form.FormChanged = False
        End If
    Next
    
    'Jump back to selected record
    DoCmd.Close acForm, rfrm.name, acSaveNo

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.AsyncFormCancelClick: " & Err.description
    Resume Exit_Function
End Sub

Public Sub AsyncFormUnload(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         AsyncFormUnload
' Date:             2022 March
' Purpose:          Handle the unload event of a form
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim ctrl As control
    Dim myForm As Variant
    Dim subFormChanged As Boolean
    Dim subFormValidity As Boolean
    
    ' Check for changes on subforms
    For Each ctrl In rfrm.Controls
        If ctrl.ControlType = acSubform And Not IsNull(rfrm.ID) Then
            subFormChanged = subFormChanged Or rfrm.Form(ctrl.name).Form.FormChanged
        End If
    Next
    
    ' Check validity of subForms
    subFormValidity = True
    For Each ctrl In rfrm.Controls
        If ctrl.ControlType = acSubform And Not IsNull(rfrm.ID) Then
            subFormValidity = subFormValidity And rfrm.Form(ctrl.name).Form.Form_Check_Validity
        End If
    Next
    
    ' Check validity of main form
    If rfrm.Form_Check_Validity = False Or subFormValidity = False Then
        If MsgBox(GetTranslation("msgbox", "entries_not_valid_close_anyway", GetDbSetting("language")), vbYesNo, GetTranslation("msgbox", "vbYesNo", GetDbSetting("language"))) = vbYes Then
            For Each ctrl In rfrm.Controls
                If ctrl.ControlType = acSubform Then
                    rfrm.Form(ctrl.name).Form.FormChanged = False
                End If
            Next
            ' Remove flag to unblock record
            If Not IsNull(rfrm.ID) Then RemoveTableFlag rfrm.DataTable, rfrm.ID
            Exit Sub
        Else
            DoCmd.CancelEvent
            Exit Sub
        End If
    End If
    
    ' Check if record was taken over by antoher user and prevent update if applies
    If Not IsNull(rfrm.ID) Then
        If GetTableFlag(rfrm.DataTable, rfrm.ID) <> GetUserName And (rfrm.FormChanged Or subFormChanged And Not rfrm.ForceClose) Then
            MsgBox GetTranslation("msgbox", "tableflag_readonly", GetDbSetting("language")), vbInformation, GetTranslation("msgbox", "vbInformation", GetDbSetting("language")) & " - " & GetTableFlag(rfrm.DataTable, rfrm.ID)
            GoTo Skip_update
        End If
    End If
    
    ' Save all changes for main form and subforms
    If rfrm.FormChanged Or subFormChanged And Not rfrm.ForceClose Then
        If MsgBox(GetTranslation("msgbox", "save_record", GetDbSetting("language")), vbYesNo, GetTranslation("msgbox", "vbYesNo", GetDbSetting("language"))) = vbYes Then
            If rfrm.FormChanged Then UpdateTable rfrm.DataTable, GetTableNameFromGuid(rfrm.tmpGuid)
        Else
            For Each ctrl In rfrm.Controls
                If ctrl.ControlType = acSubform And Not IsNull(rfrm.ID) Then
                    rfrm.Form(ctrl.name).Form.FormChanged = False
                End If
            Next
            rfrm.FormChanged = False
        End If
    End If

    ' Call individual form unload routines
    rfrm.Form_Unload_Individual
    
    ' Remove flag to unblock record
    If Not IsNull(rfrm.ID) Then RemoveTableFlag rfrm.DataTable, rfrm.ID
    
Skip_update:
    
    ' Requery the listing form with the latest version of all records
    For Each myForm In Split(rfrm.RefreshForm, ",")
        If CurrentProject.AllForms(myForm).IsLoaded Then Forms(myForm).Requery
    Next
    
Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.AsyncFormUnload: " & Err.description
    Resume Exit_Function
End Sub

Public Sub AsyncFormSave(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         AsyncFormSave
' Date:             2022 March
' Purpose:          Check for changes and save them
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim ctrl As control
    Dim myForm As Variant
    Dim subFormChanged As Boolean
    Dim subFormValidity As Boolean

    ' Check for changes on subforms
    For Each ctrl In rfrm.Controls
        If ctrl.ControlType = acSubform And Not IsNull(rfrm.ID) Then
            subFormChanged = subFormChanged Or rfrm.Form(ctrl.name).Form.FormChanged
        End If
    Next
    
     ' Check validity of subForms
    subFormValidity = True
    For Each ctrl In rfrm.Controls
        If ctrl.ControlType = acSubform And Not IsNull(rfrm.ID) Then
            subFormValidity = subFormValidity And rfrm.Form(ctrl.name).Form.Form_Check_Validity
        End If
    Next
    
    ' Check validity of main form
    If rfrm.Form_Check_Validity = False Or subFormValidity = False Then
        MsgBox GetTranslation("msgbox", "invalid_form_entries", GetDbSetting("language")), vbInformation, GetTranslation("msgbox", "vbInformation", GetDbSetting("language"))
        Exit Sub
    End If
    
    ' Check if record was taken over by antoher user and prevent update if applies
    If Not IsNull(rfrm.ID) Then
        If GetTableFlag(rfrm.DataTable, rfrm.ID) <> GetUserName And (rfrm.FormChanged Or subFormChanged And Not rfrm.ForceClose) Then
            MsgBox GetTranslation("msgbox", "tableflag_readonly", GetDbSetting("language")), vbInformation, GetTranslation("msgbox", "vbInformation", GetDbSetting("language")) & " - " & GetTableFlag(rfrm.DataTable, rfrm.ID)
            Exit Sub
        End If
    End If
    
    ' Save all changes for main form and subforms
    If rfrm.FormChanged Or subFormChanged And Not rfrm.ForceClose Then
        UpdateTable rfrm.DataTable, GetTableNameFromGuid(rfrm.tmpGuid)
        For Each ctrl In rfrm.Controls
                If ctrl.ControlType = acSubform And Not IsNull(rfrm.ID) Then
                    AsyncSubFormSaveData rfrm.Form(ctrl.name).Form
                End If
            Next
    End If
    
    ' Remove flag to unblock record
    If Not IsNull(rfrm.ID) Then RemoveTableFlag rfrm.DataTable, rfrm.ID

    ' Call individual form unload routines
    rfrm.Form_Unload_Individual

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.AsyncFormSave: " & Err.description
    Resume Exit_Function
End Sub

Public Sub AsyncFormAuditClick(rfrm As Form)
    On Error GoTo Catch_Error
    ' Open the audit log form
    If Not IsNull(rfrm.ID) Then DoCmd.OpenForm "audit", acNormal, , "table_name = '" & rfrm.DataTable & "' AND table_id = " & rfrm.ID, acFormReadOnly, acWindowNormal

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.AsyncFormAuditClick: " & Err.description
    Resume Exit_Function
End Sub

' -----------------------------------------------------------------------------------------------
' Asynchronous Sub Form Events are configured from here on
' -----------------------------------------------------------------------------------------------

Public Sub AsyncSubFormInit(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         AsyncSubFormInit
' Date:             2022 March
' Purpose:          Initialize an asynchronous form
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim ctrl As control
    Dim cs As ColumnStyle
    
    rfrm.tmpGuid = CreateGuid()
    
    ' Config Main Form
    ConfigAsyncSubForm rfrm.Form, rfrm.DataTable, rfrm.tmpGuid, False, rfrm.readOnly, rfrm.whereClause
    
    ' Adjust to best fit width
    For Each ctrl In rfrm.Controls
        If (ctrl.ControlType = acTextBox Or ctrl.ControlType = acComboBox) Then
            ctrl.ColumnWidth = -2
            'Customize columns
            cs = GetColumnStyle(rfrm, ctrl.name)
            If cs.Width <> -3 Then
                ctrl.ColumnWidth = cs.Width
                ctrl.ColumnOrder = cs.order
            End If
        End If
    Next ctrl
    
    ' Adjust to best fit hight
    rfrm.RowHeight = 300
    
    ' Individual initialisations
    rfrm.Form_Init_Individual
    
    rfrm.FormChanged = False

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.AsyncSubFormInit: " & Err.description
    Resume Exit_Function
End Sub

Public Sub AsyncSubFormSaveData(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         AsyncSubFormSaveData
' Date:             2022 March
' Purpose:          Save data of an asynchronous form
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    If rfrm.FormChanged Then
        rfrm.Requery
        UpdateTable rfrm.DataTable, GetTableNameFromGuid(rfrm.tmpGuid)
        rfrm.FormChanged = False
    End If

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.AsyncSubFormSaveData: " & Err.description
    Resume Exit_Function
End Sub

Public Sub AsyncSubFormUndoChanges(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         AsyncSubFormUndoChanges
' Date:             2022 March
' Purpose:          Undo the changes of the form
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
        If rfrm.FormChanged Then
        CopyTable rfrm.DataTable, GetTableNameFromGuid(rfrm.tmpGuid)
        If TableExists(GetCpyTableNameFromGuid(rfrm.tmpGuid)) Then CurrentDb.TableDefs.Delete GetCpyTableNameFromGuid(rfrm.tmpGuid)
        CopyTable GetTableNameFromGuid(rfrm.tmpGuid), GetCpyTableNameFromGuid(rfrm.tmpGuid)
        rfrm.Undo
        rfrm.Requery
        rfrm.FormChanged = False
    End If

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.AsyncSubFormUndoChanges: " & Err.description
    Resume Exit_Function
End Sub

Public Sub AsyncSubFormClose(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         AsyncSubFormClose
' Date:             2022 March
' Purpose:          Save changes if applies
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim ctrl As control
    
    If rfrm.FormChanged Then UpdateTable rfrm.DataTable, GetTableNameFromGuid(rfrm.tmpGuid)
    
    For Each ctrl In rfrm.Controls
        If (ctrl.ControlType = acTextBox Or ctrl.ControlType = acComboBox) Then
            ' Save column settings
            SetColumnStyle rfrm, ctrl.name, ctrl.ColumnWidth, ctrl.ColumnHidden, ctrl.ColumnOrder
        End If
    Next ctrl
    
Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.AsyncSubFormClose: " & Err.description
    Resume Exit_Function
End Sub

' -----------------------------------------------------------------------------------------------
' Synchronous Form Events are configured from here on
' -----------------------------------------------------------------------------------------------

Public Sub SyncFormOpen(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         SyncFormOpen
' Date:             2022 March
' Purpose:          Routines during opening a synchronous form
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim ctrl As control
    Dim cs As ColumnStyle
    
    ' Hide form during form adjustment are made
    On Error GoTo Skip_hide
        rfrm.visible = False
Skip_hide:
    
     ' Configure a synchronous form
    If ConfigSyncForm(rfrm.Form, rfrm.name) = False Then
        DoCmd.Close acForm, rfrm.name
        Exit Sub
    End If
    
    ' The ID column is used as hyperlink to open the responsible form to edit records
    rfrm.ID.IsHyperlink = True
    
    ' Adjust to best fit width
    For Each ctrl In rfrm.Controls
        If (ctrl.ControlType = acTextBox Or ctrl.ControlType = acComboBox) Then
            ctrl.ColumnWidth = -2
            'Customize columns
            cs = GetColumnStyle(rfrm, ctrl.name)
            If cs.Width <> -3 Then
                ctrl.ColumnWidth = cs.Width
                ctrl.ColumnOrder = cs.order
            End If
        End If
    Next ctrl
    
    ' Adjust to best fit hight
    rfrm.RowHeight = 300
    
    ' Individual initialisation routines
    rfrm.Form_Init_Individual
    
    ' Init filter setting
    rfrm.filter = Nz(DbProcedures.GetFilterSetting(rfrm.name), "")
    rfrm.FilterOn = Not IsNull(DbProcedures.GetFilterSetting(rfrm.name))
    
    ' Show form after adjustments are made
    On Error GoTo Skip_show
        rfrm.visible = True
Skip_show:

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.SyncFormOpen: " & Err.description
    Resume Exit_Function
End Sub

Public Sub SyncFormClose(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         SyncFormClose
' Date:             2022 March
' Purpose:          Routines during closing a synchronous form
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim ctrl As control
    
    For Each ctrl In rfrm.Controls
        If (ctrl.ControlType = acTextBox Or ctrl.ControlType = acComboBox) Then
            ' Save column settings
            SetColumnStyle rfrm, ctrl.name, ctrl.ColumnWidth, ctrl.ColumnHidden, ctrl.ColumnOrder
        End If
    Next ctrl
    
Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.SyncFormClose: " & Err.description
    Resume Exit_Function
End Sub

Public Sub SyncFormIdClick(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         SyncFormIdClick
' Date:             2022 March
' Purpose:          Routines after an id click event is raised
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    ' If form is already open then exit sub to avoid data loss
    If IsFormOpen(rfrm.EditForm) Then
        MsgBox GetTranslation("msgbox", "edit_form_already_open", GetDbSetting("language")), vbExclamation, GetTranslation("msgbox", "vbExclamation", GetDbSetting("language"))
        Exit Sub
    End If

    ' Open the form to edit a record. Add new record if ID is null.
    If IsNull(rfrm.ID) Then
        DoCmd.OpenForm rfrm.EditForm, acNormal, , , acFormAdd, acWindowNormal
        Forms(rfrm.EditForm).Form_Init
    Else
        DoCmd.OpenForm rfrm.EditForm, acNormal, , "id = " & rfrm.ID, acFormEdit
        Forms(rfrm.EditForm).Form_Init "WHERE id = " & rfrm.ID
    End If

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.SyncFormIdClick: " & Err.description
    Resume Exit_Function
End Sub

Public Sub SyncFormAddRecord(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         SyncFormAddRecord
' Date:             2022 March
' Purpose:          Routines when a record should be added
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
        ' If form is already open then exit sub to avoid data loss
    If IsFormOpen(rfrm.EditForm) Then
        MsgBox GetTranslation("msgbox", "edit_form_already_open", GetDbSetting("language")), vbExclamation, GetTranslation("msgbox", "vbExclamation", GetDbSetting("language"))
        Exit Sub
    End If
    
    If Not GetPermission(rfrm.name).Create Then
        MsgBox GetTranslation("msgbox", "add_record_not_supported", GetDbSetting("language")), vbExclamation, GetTranslation("msgbox", "vbExclamation", GetDbSetting("language"))
    Else
        DoCmd.OpenForm rfrm.EditForm, acNormal, , , acFormAdd, acWindowNormal
        Forms(rfrm.EditForm).Form_Init
    End If

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.SyncFormAddRecord: " & Err.description
    Resume Exit_Function
End Sub

Public Sub SyncFormDuplicateRecord(ByVal table As String, ByVal ID As Long)
     ' This sub is called by the context menu, add any code of relevance
     DuplicateRecord table, ID
End Sub

Public Sub SyncFormSaveFilter(rfrm As Form)
    ' Will call the form to save the filter settings
    'If rfrm.FilterOn And Not IsNull(rfrm.filter) Then DoCmd.OpenForm "filter_save_dlg", acNormal, , , acFormAdd, acDialog, rfrm.name
    DoCmd.OpenForm "filter_save_dlg", acNormal, , , acFormAdd, acDialog, rfrm.name
End Sub

Public Sub SyncFormApplyFilter(rfrm As Form)
    ' Will call the form to apply the filter settings
    DoCmd.OpenForm "filter_apply_dlg", acNormal, , "form = 'role' AND userid = '" & GetUserName() & "'", acFormEdit, acDialog, rfrm.name
End Sub

Public Sub SyncFormApplyClmnStd(rfrm As Form)
    ' Will call set all columns to auto width
    Dim ctrl As control
     For Each ctrl In rfrm.Controls
        If (ctrl.ControlType = acTextBox Or ctrl.ControlType = acComboBox) Then
            ' Reset to text size
            ctrl.ColumnWidth = -2
            ' Reset to auto size (needs to be done twice!)
            ctrl.ColumnWidth = -2
        End If
    Next ctrl
End Sub

' -----------------------------------------------------------------------------------------------
' Synchronous SubForm Events are configured from here on
' -----------------------------------------------------------------------------------------------

Public Sub SyncSubFormOpen(ByRef rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         SyncSubFormOpen
' Date:             2022 March
' Purpose:          Routines when a form is opened
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim ctrl As control
    Dim cs As ColumnStyle
    
    ' Hide form during form adjustment are made
    rfrm.visible = False
    
     ' Configure a synchronous form
    If ConfigSyncForm(rfrm.Form, rfrm.name) = False Then
        DoCmd.Close acForm, rfrm.name
        Exit Sub
    End If
    
    ' The ID column is used as hyperlink to open the responsible form to edit records
    rfrm.ID.IsHyperlink = True
    
    ' Adjust to best fit width
    For Each ctrl In rfrm.Controls
        If (ctrl.ControlType = acTextBox Or ctrl.ControlType = acComboBox) Then
            ctrl.ColumnWidth = -2
            'Customize columns
            cs = GetColumnStyle(rfrm, ctrl.name)
            If cs.Width <> -3 Then
                ctrl.ColumnWidth = cs.Width
                ctrl.ColumnOrder = cs.order
            End If
        End If
    Next ctrl
    
    ' Adjust to best fit hight
    rfrm.RowHeight = 300
    
    ' Individual initialisation routines
    rfrm.Form_Init_Individual
    
    ' Show form after adjustments are made
    rfrm.visible = True

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.SyncSubFormOpen: " & Err.description
    Resume Exit_Function
End Sub

Public Sub SyncSubFormIdClick(rfrm As Form)
'-------------------------------------------------------------------------------
' Function:         SyncSubFormIdClick
' Date:             2022 March
' Purpose:          Routines after an id click event is raised
' In:
' -> rfrm:          Form of choice
' Out:              -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    ' If form is already open then exit sub to avoid data loss
    If IsFormOpen(rfrm.EditForm) Then
        MsgBox GetTranslation("msgbox", "edit_form_already_open", GetDbSetting("language")), vbExclamation, GetTranslation("msgbox", "vbExclamation", GetDbSetting("language"))
        Exit Sub
    End If

    ' Open the form to edit a record. Add new record if ID is null.
    If IsNull(rfrm.ID) Then
        DoCmd.OpenForm rfrm.EditForm, acNormal, , , acFormAdd, acWindowNormal
        Forms(rfrm.EditForm).Form_Init
    Else
        DoCmd.OpenForm rfrm.EditForm, acNormal, , "id = " & rfrm.ID, acFormEdit
        Forms(rfrm.EditForm).Form_Init "WHERE id = " & rfrm.ID
    End If

Exit_Function:
    Exit Sub
Catch_Error:
    AddErrorLog Err.Number, "ManageFormEvents.SyncSubFormIdClick: " & Err.description
    Resume Exit_Function
End Sub

Public Sub SyncSubFormSaveFilter(rfrm As Form)
    ' Open form to save filter settings
    If rfrm.FilterOn And Not IsNull(rfrm.filter) Then DoCmd.OpenForm "filter_save_dlg", acNormal, , , acFormAdd, acDialog, rfrm.name
End Sub

Public Sub SyncSubFormApplyFilter(rfrm As Form)
    ' Open form to apply filter settings
    DoCmd.OpenForm "filter_apply_dlg", acNormal, , "form = 'role' AND userid = '" & GetUserName() & "'", acFormEdit, acDialog, rfrm.name
End Sub

