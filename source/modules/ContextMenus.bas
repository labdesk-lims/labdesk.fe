Attribute VB_Name = "ContextMenus"
'################################################################################################
' This module will register context menus used in forms. Initialization will be done in module
' AutoExec.
'################################################################################################

Option Compare Database
Option Explicit

Private Function ContextMenuExist(ByVal name As String) As Boolean
    Dim cmbBar As CommandBar
    
    For Each cmbBar In Application.CommandBars
        If cmbBar.name = name Then ContextMenuExist = True
    Next
End Function

Private Sub ContextMenuReset()
    Dim bar As CommandBar
    
    'Reset all custom commandbars
    For Each bar In Application.CommandBars
        If (bar.BuiltIn = False) And (bar.visible = False) Then
            bar.Delete
        End If
    Next bar
End Sub

Private Function ContextMenuAdd(ByVal name As String, ByVal parameter As String, ByVal faceId As Integer, Optional ByVal beginGroup As Boolean) As Boolean
    Dim cmbBar As CommandBar
    Dim cmbBtn_CreateNew, cmbBtn_Edit As CommandBarButton
    
    If ContextMenuExist(name) Then
        Set cmbBar = CommandBars(name)
    Else
        Set cmbBar = CommandBars.Add(name, msoBarPopup)
    End If
    
    Set cmbBtn_CreateNew = cmbBar.Controls.Add(msoControlButton)
    
    If beginGroup Then cmbBtn_CreateNew.beginGroup = True
    cmbBtn_CreateNew.faceId = faceId
    cmbBtn_CreateNew.caption = DbProcedures.GetTranslation("mdlContextMenu", "contextmnu_" & name & "_" & parameter, DbConnect.GetDbSetting("language"))
    cmbBtn_CreateNew.parameter = parameter
    cmbBtn_CreateNew.OnAction = "fnCall"
End Function

Public Sub ContextMenuInit()
    ContextMenuReset
    
    'Create menu_std entries
    ContextMenuAdd "menu_std", "add_record", 192
    ContextMenuAdd "menu_std", "apply_filter", 601, True
    ContextMenuAdd "menu_std", "save_filter", 602
    ContextMenuAdd "menu_std", "clmn_dlg", 8, True
    ContextMenuAdd "menu_std", "clmn_std", 6
    ContextMenuAdd "menu_std", "audit_trail", 29, True
    ContextMenuAdd "menu_std", "refresh", 37, True
    
    'Create menu_dpl entries
    ContextMenuAdd "menu_dpl", "add_record", 192
    ContextMenuAdd "menu_dpl", "duplicate_record", 19
    ContextMenuAdd "menu_dpl", "apply_filter", 601, True
    ContextMenuAdd "menu_dpl", "save_filter", 602
    ContextMenuAdd "menu_dpl", "clmn_dlg", 8, True
    ContextMenuAdd "menu_dpl", "clmn_std", 6
    ContextMenuAdd "menu_dpl", "audit_trail", 29, True
    ContextMenuAdd "menu_dpl", "refresh", 37, True
    
    'Create menu_rqt entries
    ContextMenuAdd "menu_rqt", "add_record", 192
    ContextMenuAdd "menu_rqt", "add_subrequest", 188
    ContextMenuAdd "menu_rqt", "apply_filter", 601, True
    ContextMenuAdd "menu_rqt", "save_filter", 602
    ContextMenuAdd "menu_rqt", "clmn_dlg", 8, True
    ContextMenuAdd "menu_rqt", "clmn_std", 6
    ContextMenuAdd "menu_rqt", "show_report", 195, True
    ContextMenuAdd "menu_rqt", "show_worksheet", 144
    ContextMenuAdd "menu_rqt", "show_sticker", 191
    ContextMenuAdd "menu_rqt", "audit_trail", 29, True
    ContextMenuAdd "menu_rqt", "refresh", 37, True
    
    'Create menu_flt entries
    ContextMenuAdd "menu_flt", "apply_filter", 601
    ContextMenuAdd "menu_flt", "save_filter", 602
    ContextMenuAdd "menu_flt", "clmn_dlg", 8, True
    ContextMenuAdd "menu_flt", "clmn_std", 6
    ContextMenuAdd "menu_flt", "audit_trail", 29, True
    ContextMenuAdd "menu_flt", "refresh", 37, True
    
    'Create menu_fle entries
    ContextMenuAdd "menu_fle", "open_file", 18
    ContextMenuAdd "menu_fle", "upload_file", 1632
    ContextMenuAdd "menu_fle", "download_file", 1631
    
    'Create menu_sub entries
    ContextMenuAdd "menu_sub", "audit_trail_subform", 29
    ContextMenuAdd "menu_sub", "clmn_dlg", 8
    
    'Create menu_usr entires
    ContextMenuAdd "menu_usr", "add_record", 192
    ContextMenuAdd "menu_usr", "activate_user", 343, True
    ContextMenuAdd "menu_usr", "deactivate_user", 342
    ContextMenuAdd "menu_usr", "apply_filter", 601, True
    ContextMenuAdd "menu_usr", "save_filter", 602
    ContextMenuAdd "menu_usr", "clmn_dlg", 8, True
    ContextMenuAdd "menu_usr", "clmn_std", 6
    ContextMenuAdd "menu_usr", "audit_trail", 29, True
    ContextMenuAdd "menu_usr", "refresh", 37, True
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
    
    strAction = CommandBars.ActionControl.parameter
    
    Set frmCurrentForm = Screen.ActiveForm
    
    Select Case strAction
        Case "add_record"
            Forms(frmCurrentForm.name).AddRecord
            
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
        
        Case "clmn_dlg"
            DoCmd.RunCommand acCmdUnhideColumns
        
        Case "upload_file"
            Forms(frmCurrentForm.name).Form("attachment_sbf").Form.UploadFile
        
        Case "download_file"
            Forms(frmCurrentForm.name).Form("attachment_sbf").Form.DownloadFile
        
        Case "open_file"
            Forms(frmCurrentForm.name).Form("attachment_sbf").Form.OpenFile

        Case "show_report"
            Forms(frmCurrentForm.name).ShowReport
        
        Case "show_sticker"
            Forms(frmCurrentForm.name).ShowSticker

        Case "show_worksheet"
            Forms(frmCurrentForm.name).ShowWorksheet
        
        Case "show_invoice"
            Forms(frmCurrentForm.name).ShowInvoice
        
        Case "audit_trail"
            If Not isnull(Forms(frmCurrentForm.name).ID) Then DoCmd.OpenForm "_AuditTrail", acNormal, , , acFormReadOnly, acWindowNormal, Forms(frmCurrentForm.name).LinkedTable & ", " & Forms(frmCurrentForm.name).ID
            
        Case "audit_trail_subform"
            If Not isnull(Forms(frmCurrentForm.name).Form(Screen.ActiveControl.Parent.name).Form.ID) Then DoCmd.OpenForm "_AuditTrail", acNormal, , , acFormReadOnly, acWindowNormal, Forms(frmCurrentForm.name).Form(Screen.ActiveControl.Parent.name).Form.DataTable & ", " & Forms(frmCurrentForm.name).Form(Screen.ActiveControl.Parent.name).Form.ID
            
        Case "refresh"
            Forms(frmCurrentForm.name).Requery
            
        Case "activate_user"
            Forms(frmCurrentForm.name).ActivateUser
        
         Case "deactivate_user"
            If Not isnull(Forms(frmCurrentForm.name).ID) Then ManageLicence.DeActivateUser Forms(frmCurrentForm.name).ID
            Forms(frmCurrentForm.name).Requery
    End Select

Exit_Function:
    Exit Sub
Catch_Error:
    MsgBox "Error (mdlContextMenu - " & CommandBars.ActionControl.parameter & "): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Sub


