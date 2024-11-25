Attribute VB_Name = "AutoExec"
'################################################################################################
' This module will initialize the application to work properly with all components of relevance.
' InitApp will be called by the macro AutoExec during startup.
'################################################################################################

Option Compare Database
Option Explicit

' System wide configuration
Global config As New SysConfig

' User name and id of session
Global pUserName As String
Global pUserId As Integer
Global pDeploy As Boolean

' -----------------------------------------------------------------------------------------------
' Any code goes from here
' -----------------------------------------------------------------------------------------------

Public Function InitApp() As Boolean
'-------------------------------------------------------------------------------
' Function:     InitApp
' Date:         2022 January
' Purpose:      Init the application settings
' Parameters:   -
' Out:          Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    config.Init ".labdesk"
    
    'Deploy dependencies
    SysCmd acSysCmdSetStatus, "Deploy dependencies"
    Deployment.DeployDependency
    
    'Reset login password
    SysCmd acSysCmdSetStatus, "System cleanup"
    DbProcedures.ResetLoginPassword
    
    'Interface cosmetics
    SysCmd acSysCmdSetStatus, "Applying some interface cosmetics"
    ManageGui.AddAppProperty "AppTitle", dbText, config.AppTitle
    Application.RefreshTitleBar
    HideNavPane DbConnect.GetDbSetting("navpane")
    ManageGui.AddAppProperty "AppIcon", dbText, CurrentProject.path & "\icon.ico"
    CurrentDb.Properties("UseAppIconForFrmRpt") = True
    Application.RefreshTitleBar
    
    'Check for Runtime-Mode and disable shift
    SysCmd acSysCmdSetStatus, "Checking Runtime-Mode"
    If Not System.RunTimeMode And Not GetDbSetting("devmode") Then Err.Raise vbObjectError + 513, , "Runtime-Mode switched off. Startup aborted."
    
login_form:
    'Show login form
    SysCmd acSysCmdSetStatus, "User login"
    DoCmd.OpenForm "login", acNormal, , , acFormEdit, acDialog
    
    'Try to connect to server
    SysCmd acSysCmdSetStatus, "Try to connect to server"
    If Not PingOk(CStr(Split(DbConnect.GetDbSetting("server"), "/")(0))) Then
        'Show login form
        MsgBox "Can not connect to server " & DbConnect.GetDbSetting("server"), vbExclamation
        GoTo login_form
    End If
    
    'Try to attach tables
    SysCmd acSysCmdSetStatus, "Try to attach tables"
    If Not DbConnect.ConnectDb(DbConnect.GetDbSetting("server"), DbConnect.GetDbSetting("database"), config.DSNLessTables, DbConnect.GetDbSetting("winauth"), DbConnect.GetDbSetting("user"), DbConnect.GetDbSetting("password")) Then
        MsgBox "Wrong login credential provided", vbExclamation
        GoTo login_form
    End If
    
    'Check Version (raise error if backend does not match)
    SysCmd acSysCmdSetStatus, "Checking version integrity"
    If DbProcedures.GetBeVersion() <> config.BeVersion Then
        Err.Raise vbObjectError + 513, , "Backend version not supported. (Actual: " & DbProcedures.GetBeVersion() & " | Required: " & config.BeVersion & ")"
    End If
    
    'Install appplication if no users exists
    If InstallationPending Then
        If MsgBox("Press OK to install application. This may take a while.", vbOKCancel, "Information") = vbOK Then
            SysCmd acSysCmdSetStatus, "Install application"
            Deployment.Install
            MsgBox "System installation finished. Restart application to proceed.", vbInformation, "Information"
        End If
        Exit Function
    End If
    
    'Update application if actual version is higher than database entry
    If GetFeVersion < config.FeVersion Then
        If MsgBox("Press OK to update application. This may take a while.", vbOKCancel, "Information") = vbOK Then
            SysCmd acSysCmdSetStatus, "Update application"
            Deployment.Update
            SetFeVersion
            MsgBox "System update finished. Restart application to proceed.", vbInformation, "Information"
        End If
        Exit Function
    End If
    
    'Clean Cache (stored in the user folder under pCacheFolder)
    SysCmd acSysCmdSetStatus, "Clean cache"
    System.CleanCache
    
    'Init local tables
    SysCmd acSysCmdSetStatus, "Init local tables"
    LocalTables.InitLocalTables
    
    'Init context menus
    SysCmd acSysCmdSetStatus, "Init context menus"
    ContextMenus.ContextMenuInit
    
    'Check if user exists otherwise add
    SysCmd acSysCmdSetStatus, "Check user registration"
    DbProcedures.AddUser
    
    'Open the closing event form
    SysCmd acSysCmdSetStatus, "Init background worker"
    DoCmd.OpenForm "_background_worker_form", acNormal, , , acFormReadOnly, acWindowNormal
    Forms("_background_worker_form").TimerInterval = config.BackgroundWorkerInterval
    Forms("_background_worker_form").visible = False
    
    'Open desktop tab
    SysCmd acSysCmdSetStatus, "Open desktop"
    If Application.SysCmd(acSysCmdGetObjectState, acForm, "desktop") <> acObjStateOpen And Nz(GetFieldValue("setup", "show_desktop"), False) Then DoCmd.OpenForm "desktop", acNormal, , , acFormReadOnly, acWindowNormal
    
    'Fill status bar text with user name
    SysCmd acSysCmdSetStatus, "User: " & DbProcedures.GetUserName()
    
    InitApp = True
    
Exit_Function:
    Exit Function
Catch_Error:
    InitApp = False
    MsgBox "The following error was detected during startup: " & vbCrLf & Err.description, vbCritical, "Critical Error"
    Application.Quit
    Resume Exit_Function
End Function
