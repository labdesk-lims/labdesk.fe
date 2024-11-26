Attribute VB_Name = "Deployment"
'################################################################################################
' This module provides support for deployment of new lim-systems as well as updating existing
' installations.
'################################################################################################

Option Compare Database
Option Explicit

Public Sub Prepare(Optional ByVal deploy As Boolean = True)
'-------------------------------------------------------------------------------
'Function:          PrepareSetup
'Date:              2022 February
'Purpose:           Create to setup a new installation of labdesk
'-------------------------------------------------------------------------------
    
    Dim td As TableDef
    Dim tbl As Variant
    
    '//Delete all temporary installation tables
    For Each tbl In CurrentDb.TableDefs
        If tbl.name Like "_*" Then CurrentDb.TableDefs.Delete tbl.name
    Next
    
    ' Translations
    SysCmd acSysCmdSetStatus, "Copy translations"
    If TableExists("_translation") Then CurrentDb.TableDefs.Delete "_translation"
    CreateTable "translation", "_translation"
    CopyTable "translation", "_translation"
    
    ' Roles
    SysCmd acSysCmdSetStatus, "Copy roles"
    If TableExists("_role") Then CurrentDb.TableDefs.Delete "_role"
    CreateTable "role", "_role"
    CopyTable "role", "_role"
    
    ' Permissions
    SysCmd acSysCmdSetStatus, "Copy permissions"
    If TableExists("_permission") Then CurrentDb.TableDefs.Delete "_permission"
    CreateTable "permission", "_permission"
    CopyTable "permission", "_permission"
    
    ' Role permission cross table
    SysCmd acSysCmdSetStatus, "Copy role configuration"
    If TableExists("_role_permission") Then CurrentDb.TableDefs.Delete "_role_permission"
    CreateTable "role_permission", "_role_permission"
    CopyTable "role_permission", "_role_permission"
    
    If deploy Then
        ' Unlink DSN less tables
        SysCmd acSysCmdSetStatus, "Unlink remote tables"
        UnAttachDSNLessTables config.DSNLessTables
        
        'Hide the navigation pane
        HideNavPane True
    
        'Disable full menu
        AddAppProperty "AllowFullMenus", dbBoolean, False
    
        ' Reset db settings
        SysCmd acSysCmdSetStatus, "Reset database"
        CurrentDb.Execute "UPDATE dbsetup SET server = Null, user = Null, password = Null, navpane = 0, devmode = 0"
        
        'Disable shift
        SysCmd acSysCmdSetStatus, "Disable shift"
        DisableShift
        
        'Set deployment mode
        pDeploy = True
    End If
    
    SysCmd acSysCmdSetStatus, "Preparation finished"
End Sub

Public Sub Install()
'-------------------------------------------------------------------------------
'Function:          SysInit
'Date:              2021 October
'Purpose:           Update the database with a local copy of all configurations
'-------------------------------------------------------------------------------
    ' InitTranslations
    UpdateTranslation
    UpdateRole
    UpdatePermission
    UpdateRolePermission
    AddAdmin
    
    'Set Frontend Version
    SetFeVersion
End Sub

Public Sub Update()
'-------------------------------------------------------------------------------
'Function:          SysInit
'Date:              2021 October
'Purpose:           Update the database with a local copy of all configurations
'-------------------------------------------------------------------------------
    ' InitTranslations
    UpdateTranslation
    'UpdateRole
    UpdatePermission
    'UpdateRolePermission
    
    'Set Frontend Version
    SetFeVersion
End Sub

Public Function UploadDependency(ByVal fname As String) As Boolean
'-------------------------------------------------------------------------------
'Function:          UploadDependency
'Date:              2024 January
'Purpose:           Upload any dependency to table dbfiles
'-> fname:          Path and name of file to upload
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim binary As Object
    Dim rs As Recordset
    
    CurrentDb.Execute "INSERT INTO dbfiles (FileName) VALUES('" & GetFileName(fname) & "')"
    
    Set rs = CurrentDb.OpenRecordset("SELECT binary FROM dbfiles WHERE FileName = '" & GetFileName(fname) & "'", dbOpenDynaset, dbSeeChanges)
    
    rs.Edit
    DbBlob.FileToBlob fname, rs!binary
    rs.Update
    rs.Close
    
    UploadDependency = True

Exit_Function:
    Set rs = Nothing
    Exit Function
Catch_Error:
    UploadDependency = False
    Resume Exit_Function
End Function

Public Function DeployDependency() As Boolean
'------------------------------------------------------------------------------
'Function:          DeployDependency
'Date:              2024 January
'Purpose:           Deploy all dependencies from table dbfiles to current dir
'------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim obj As Object
    Dim rs As Recordset
    
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM dbfiles", dbOpenDynaset, dbSeeChanges)
    
    rs.MoveFirst
    Do Until rs.EOF
        If rs!deploy Then DbBlob.BlobToFile CurrentProject.path & "\" & rs!FileName, rs!binary
        rs.MoveNext
    Loop
    
    rs.Close
    
    DeployDependency = True
    
Exit_Function:
    Set rs = Nothing
    Exit Function
Catch_Error:
    DeployDependency = False
    Resume Exit_Function
End Function

Private Sub AddTranslation(ByVal container As String, ByVal item As String, ByVal en As String, ByVal de As String)
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT COUNT(id) As cnt FROM translation WHERE container = '" & container & "' AND item = '" & item & "'", dbOpenDynaset, dbSeeChanges)
    
    If rs!cnt = 0 Then
        db.Execute "INSERT INTO translation (container, item, en, de) VALUES ('" & container & "', '" & item & "', '" & en & "', '" & de & "')"
    Else
        db.Execute "UPDATE translation SET en = '" & en & "', de = '" & de & "' WHERE container = '" & container & "' AND item = '" & item & "'"
    End If
    
Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Sub
Catch_Error:
    MsgBox "Error (mdlSysPrep - AddTranslation): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Sub

Private Sub UpdateTranslation()
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT * FROM _translation")
    
    Do While Not rs.EOF
        AddTranslation Nz(rs!container, ""), Nz(rs!item, ""), Nz(rs!en, ""), Nz(rs!de, "")
        rs.MoveNext
    Loop

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Sub
Catch_Error:
    MsgBox "Error (mdlSysPrep - UpdateTranslation): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Sub

Private Sub AddRole(ByVal title As String, ByVal description As String, ByVal administrative As Boolean)
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT COUNT(id) As cnt FROM role WHERE title = '" & title & "'", dbOpenDynaset, dbSeeChanges)
    
    If rs!cnt = 0 Then
        db.Execute "INSERT INTO role (title, description, administrative) VALUES ('" & title & "', '" & description & "', " & CInt(administrative) & ")"
    End If

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Sub
Catch_Error:
    MsgBox "Error (mdlSysPrep - AddRole): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Sub

Private Sub UpdateRole()
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT * FROM _role")
    
    Do While Not rs.EOF
        AddRole Nz(rs!title, ""), Nz(rs!description, ""), Nz(rs!administrative, "")
        rs.MoveNext
    Loop

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Sub
Catch_Error:
    MsgBox "Error (mdlSysPrep - UpdateRole): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Sub

Private Sub AddPermission(ByVal title As String, ByVal description As String)
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT COUNT(id) As cnt FROM permission WHERE title = '" & title & "'", dbOpenDynaset, dbSeeChanges)
    
    If rs!cnt = 0 Then
        db.Execute "INSERT INTO permission (title, description) VALUES ('" & title & "', '" & description & "')"
    End If

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Sub
Catch_Error:
    MsgBox "Error (mdlSysPrep - AddPermission): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Sub

Private Sub UpdatePermission()
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT * FROM _permission")
    
    Do While Not rs.EOF
        If Not IsNull(rs!description) Then
            AddPermission Nz(rs!title, ""), Nz(rs!description, "")
        Else
            AddPermission Nz(rs!title, ""), ""
        End If
        rs.MoveNext
    Loop

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Sub
Catch_Error:
    MsgBox "Error (mdlSysPrep - UpdatePermission): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Sub

Private Sub AddRolePermission(ByVal role As Long, ByVal permission As Long, ByVal can_create As Boolean, ByVal can_read As Boolean, ByVal can_update As Boolean, ByVal can_delete As Boolean)
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    
    db.Execute "UPDATE role_permission SET can_create = " & CInt(can_create) & ", can_read = " & CInt(can_read) & ", can_update = " & CInt(can_update) & ", can_delete = " & CInt(can_delete) & " WHERE permission = " & permission & " AND role = " & role

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Sub
Catch_Error:
    MsgBox "Error (mdlSysPrep - AddRolePermission): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Sub

Private Sub UpdateRolePermission()
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    Dim ra, rb, rx, ry As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT * FROM _role_permission")
    
    Do While Not rs.EOF
        Set ra = db.OpenRecordset("SELECT * FROM _role WHERE id = " & rs!role)
        Set rb = db.OpenRecordset("SELECT * FROM role WHERE title = '" & ra!title & "'", dbOpenDynaset, dbSeeChanges)
        Set rx = db.OpenRecordset("SELECT * FROM _permission WHERE id = " & rs!permission)
        Set ry = db.OpenRecordset("SELECT * FROM permission WHERE title = '" & rx!title & "'", dbOpenDynaset, dbSeeChanges)
        
        AddRolePermission rb!ID, ry!ID, True, rs!can_read, rs!can_update, rs!can_delete
        rs.MoveNext
    Loop

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Sub
Catch_Error:
    MsgBox "Error (mdlSysPrep - UpdateRolePermission): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Sub
