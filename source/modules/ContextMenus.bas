Attribute VB_Name = "ContextMenus"
'################################################################################################
' This module will register context menus used in forms. Initialization will be done in module
' AutoExec.
'################################################################################################

Option Compare Database
Option Explicit

Public Sub InitContextMenus()
    Dim bar As CommandBar
    
    'Reset all commandbars
    For Each bar In Application.CommandBars
        If (bar.BuiltIn = False) And (bar.visible = False) Then
            bar.Delete
        End If
    Next bar
    
    'Create all commandbars
    CreateContextMenu "menu_std", True, False, True, False, False, False, False
    CreateContextMenu "menu_dpl", True, False, True, True, False, False, False
    CreateContextMenu "menu_rqt", True, True, True, False, True, False, False
    CreateContextMenu "menu_flt", False, False, True, False, False, False, False
    CreateContextMenu "menu_fle", False, False, False, False, False, True, False
    CreateContextMenu "menu_sub", False, False, False, False, False, False, True
End Sub

Private Sub CreateContextMenu(ByVal name As String, ByVal Create As Boolean, ByVal createsub As Boolean, ByVal standard As Boolean, ByVal duplicate As Boolean, ByVal documents As Boolean, ByVal files As Boolean, ByVal subform As Boolean)
    Dim cmbBar As CommandBar
    Dim cmbBtn_CreateNew, cmbBtn_Edit As CommandBarButton
    
On Error GoTo Skip
    ' Delete context menu if already exists to register the new one
    CommandBars(name).Delete
Skip:

    Set cmbBar = CommandBars.Add(name, msoBarPopup)

    If Create Then
        ' Button to add record
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.FaceId = 192
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_add_record", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "add_record"
        cmbBtn_CreateNew.OnAction = "fnCall"
    End If
    
    If createsub Then
        ' Button to add sub request
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.FaceId = 188
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_add_subrequest", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "add_subrequest"
        cmbBtn_CreateNew.OnAction = "fnCall"
    End If
    
    If standard Then
        ' Button to apply a filter
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.BeginGroup = True
        cmbBtn_CreateNew.FaceId = 601
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_apply_filter", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "apply_filter"
        cmbBtn_CreateNew.OnAction = "fnCall"
        
        ' Button to create a filter
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.FaceId = 602
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_save_filter", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "save_filter"
        cmbBtn_CreateNew.OnAction = "fnCall"
        
        'Button to select columns
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.BeginGroup = True
        cmbBtn_CreateNew.FaceId = 8
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_clmn_dlg", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "clmn_dlg"
        cmbBtn_CreateNew.OnAction = "fnCall"
        
        'Button to reset column size
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.FaceId = 6
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_clmn_std", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "clmn_std"
        cmbBtn_CreateNew.OnAction = "fnCall"
        
        'Audit trail
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.BeginGroup = True
        cmbBtn_CreateNew.FaceId = 29
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_audit_trail", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "audit_trail"
        cmbBtn_CreateNew.OnAction = "fnCall"
        
        'Refresh form
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.BeginGroup = True
        cmbBtn_CreateNew.FaceId = 37
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_refresh", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "refresh"
        cmbBtn_CreateNew.OnAction = "fnCall"
    End If
    
    If duplicate Then
        'Duplicate form
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.BeginGroup = True
        cmbBtn_CreateNew.FaceId = 19
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_duplicate", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "duplicate_record"
        cmbBtn_CreateNew.OnAction = "fnCall"
    End If
    
    If documents Then
        'Button to show preliminary report
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.BeginGroup = True
        cmbBtn_CreateNew.FaceId = 195
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_print_report", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "show_report"
        cmbBtn_CreateNew.OnAction = "fnCall"
        
        'Button to show worksheet
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.FaceId = 144
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_print_worksheet", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "show_worksheet"
        cmbBtn_CreateNew.OnAction = "fnCall"
        
        'Button to show sticker
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.FaceId = 191
        cmbBtn_CreateNew.Caption = GetTranslation("mdlContextMenu", "contextmnu_" & name & "_print_sticker", GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "show_sticker"
        cmbBtn_CreateNew.OnAction = "fnCall"
    End If
    
    If files Then
        ' Button to down- or upload files
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.BeginGroup = True
        cmbBtn_CreateNew.FaceId = 1632
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_upload", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "upload_file"
        cmbBtn_CreateNew.OnAction = "fnCall"
        
        ' Button to download files
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.FaceId = 1631
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_download", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "download_file"
        cmbBtn_CreateNew.OnAction = "fnCall"
        
        ' Button to open files
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.FaceId = 18
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_open", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "open_file"
        cmbBtn_CreateNew.OnAction = "fnCall"
    End If
    
    If subform Then
        'Audit trail
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.BeginGroup = True
        cmbBtn_CreateNew.FaceId = 29
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_audit_trail", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "audit_trail_subform"
        cmbBtn_CreateNew.OnAction = "fnCall"
        
        'Button to select columns
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.BeginGroup = True
        cmbBtn_CreateNew.FaceId = 8
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_clmn_dlg", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "clmn_dlg"
        cmbBtn_CreateNew.OnAction = "fnCall"
        
        'Refresh form
        Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
        cmbBtn_CreateNew.BeginGroup = True
        cmbBtn_CreateNew.FaceId = 37
        cmbBtn_CreateNew.Caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_refresh", DbConnect.GetDbSetting("language"))
        cmbBtn_CreateNew.Parameter = "refresh"
        cmbBtn_CreateNew.OnAction = "fnCall"
    End If
End Sub

Public Sub fnCall()
'----------------------------------------------------------------------------------------
'Function:      fnCall
' Date:         2022 January
' Purpose:      Calls the procedure in the form named in the parameter strAction
' In:           -
' Out:          -
'----------------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim frmCurrentForm As Form
    Dim strAction As String
    
    strAction = CommandBars.ActionControl.Parameter
    
    Set frmCurrentForm = Screen.ActiveForm
    
    Select Case strAction
        Case "add_record"
            Forms(frmCurrentForm.name).AddRecord
        
        Case "activate_user"
            Forms(frmCurrentForm.name).ActivateUser
        
        Case "deactivate_user"
            Forms(frmCurrentForm.name).DeactivateUser
            
        Case "duplicate_record"
            Forms(frmCurrentForm.name).DuplicateRecord
        
        Case "add_subrequest"
            Forms(frmCurrentForm.name).AddSubRequest
            
        Case "save_filter"
            Forms(frmCurrentForm.name).SaveFilter
            
        Case "apply_filter"
            Forms(frmCurrentForm.name).ApplyFilter
        
        Case "clmn_std"
            Forms(frmCurrentForm.name).ApplyClmnStd
        
        Case "upload_file"
            Forms(frmCurrentForm.name).Form("attachment_sbf").Form.UploadFile
        
        Case "download_file"
            Forms(frmCurrentForm.name).Form("attachment_sbf").Form.DownloadFile
        
        Case "open_file"
            Forms(frmCurrentForm.name).Form("attachment_sbf").Form.OpenFile
            
        Case "clmn_dlg"
            DoCmd.RunCommand acCmdUnhideColumns

        Case "show_report"
            Forms(frmCurrentForm.name).ShowReport
        
        Case "show_sticker"
            Forms(frmCurrentForm.name).ShowSticker

        Case "show_worksheet"
            Forms(frmCurrentForm.name).ShowWorksheet
        
        Case "show_invoice"
            Forms(frmCurrentForm.name).ShowInvoice
        
        Case "audit_trail"
            If Not IsNull(Forms(frmCurrentForm.name).ID) Then DoCmd.OpenForm "_AuditTrail", acNormal, , , acFormReadOnly, acWindowNormal, Forms(frmCurrentForm.name).LinkedTable & ", " & Forms(frmCurrentForm.name).ID
            
        Case "audit_trail_subform"
            If Not IsNull(Forms(frmCurrentForm.name).Form(Screen.ActiveControl.Parent.name).Form.ID) Then DoCmd.OpenForm "_AuditTrail", acNormal, , , acFormReadOnly, acWindowNormal, Forms(frmCurrentForm.name).Form(Screen.ActiveControl.Parent.name).Form.DataTable & ", " & Forms(frmCurrentForm.name).Form(Screen.ActiveControl.Parent.name).Form.ID
            
        Case "refresh"
            Forms(frmCurrentForm.name).Requery
    End Select

Exit_Function:
    Exit Sub
Catch_Error:
    MsgBox "Error (mdlContextMenu - " & CommandBars.ActionControl.Parameter & "): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Sub


