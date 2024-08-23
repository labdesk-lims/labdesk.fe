Attribute VB_Name = "DbProcedures"
'################################################################################################
' This module provides functions to call t-tql and local data procedures.
'################################################################################################

Option Compare Database
Option Explicit

' Used to store permissions
Public Type CRUD
    Create As Boolean
    Read As Boolean
    Update As Boolean
    Delete As Boolean
End Type

' Used to store column settings of a datasheetview
Public Type ColumnStyle
    Width As Integer
    hidden As Boolean
    order As Integer
End Type

Public Function GetFilterSetting(ByVal rfrm As String) As Variant
'-------------------------------------------------------------------------------
' Function:         GetFilterSetting
' Date:             2025 May
' Purpose:          Get the actual filter setting
' In:
' -> rfrm:          Form of interest to get filter settings from
' Out:              String Or Null
'------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    'Get filter setting
    Set rs = db.OpenRecordset("SELECT filter FROM filter WHERE active = 1 AND userid = '" & DbProcedures.GetUserName & "' AND form = '" & rfrm & "'", dbOpenDynaset, dbSeeChanges)
    
    If Not rs.EOF Then GetFilterSetting = rs(0) Else GetFilterSetting = Null
    
Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    GetFilterSetting = Null
    Resume Exit_Function
End Function

Public Sub SetFilterSetting(ByVal rfrm As String, filter As String)
'-------------------------------------------------------------------------------
' Function:         GetFilterSetting
' Date:             2025 May
' Purpose:          Get the actual filter setting
' In:
' -> rfrm:          Form of interest to get filter settings from
' Out:              String Or Null
'------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    'Get filter setting
    Set rs = db.OpenRecordset("SELECT id FROM filter WHERE global = False AND userid = '" & DbProcedures.GetUserName & "' AND form = '" & rfrm & "' AND filter = '" & filter & "'", dbOpenDynaset, dbSeeChanges)
    
    If Not rs.EOF Then db.Execute "UPDATE filter SET active = 1 WHERE id = " & rs(0), dbSeeChanges
    
Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Sub
Catch_Error:
    Resume Exit_Function
End Sub

Public Function SetColumnStyle(ByVal rfrm As Form, ByVal clmn As String, ByVal Width As Integer, ByVal hidden As Boolean, ByVal order As Integer) As Boolean
'-------------------------------------------------------------------------------
' Function:         SetColumnStyle
' Date:             2022 March
' Purpose:          Set the style of a column
' In:
' -> rfrm:          Form of interest to apply style settings
' -> clmn:          The column of choice
' -> width:         Width of the column (choose zero to hide)
' -> order:         The order/position of the column
' Out:              Done (T/F)
'------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    Dim cs As ColumnStyle
    
    Set db = CurrentDb()
    'Local column style switch off
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM lcl_columns WHERE user_id = '" & GetUserName() & "' AND table_id = '" & rfrm.name & "' AND column_id = '" & clmn & "'", dbOpenDynaset, dbSeeChanges)
    'Set rs = db.OpenRecordset("SELECT COUNT(*) FROM columns WHERE user_id = '" & GetUserName() & "' AND table_id = '" & rfrm.name & "' AND column_id = '" & clmn & "'", dbOpenDynaset, dbSeeChanges)
    
    If hidden Then Width = 0 'workaround to make the hidden property persistent
    
    If Not rs(0) > 0 Then
        'Local column style switch off
        db.Execute "INSERT INTO lcl_columns (user_id, table_id, column_id, column_width, column_hidden, column_order) VALUES('" & GetUserName() & "', '" & rfrm.name & "', '" & clmn & "', " & Width & ", " & CInt(hidden) & ", " & order & ")"
        'db.Execute "INSERT INTO columns (user_id, table_id, column_id, column_width, column_hidden, column_order) VALUES('" & GetUserName() & "', '" & rfrm.name & "', '" & clmn & "', " & Width & ", " & CInt(hidden) & ", " & order & ")"
    Else
        'Local column style switch off
        db.Execute "UPDATE lcl_columns SET column_width = " & Width & ", column_hidden = " & CInt(hidden) & ", column_order = " & order & " WHERE user_id = '" & GetUserName() & "' AND table_id = '" & rfrm.name & "' AND column_id = '" & clmn & "'"
        'db.Execute "UPDATE columns SET column_width = " & Width & ", column_hidden = " & CInt(hidden) & ", column_order = " & order & " WHERE user_id = '" & GetUserName() & "' AND table_id = '" & rfrm.name & "' AND column_id = '" & clmn & "'"
    End If
    
    SetColumnStyle = True
    
Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function GetColumnStyle(ByVal rfrm As Form, ByVal clmn As String) As ColumnStyle
'-------------------------------------------------------------------------------
' Function:         GetColumnStyle
' Date:             2022 March
' Purpose:          Set the style of a column
' In:
' -> rfrm:          Form of interest to apply style settings
' -> clmn:          The column of choice
' Out:              ColumnStyle (returns width = -3 in case of errors)
'------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    Dim cs As ColumnStyle
    
    Set db = CurrentDb()
    
    'Local column style switch off
    Set rs = db.OpenRecordset("SELECT * FROM lcl_columns WHERE user_id = '" & GetUserName() & "' AND table_id = '" & rfrm.name & "' AND column_id = '" & clmn & "'", dbOpenDynaset, dbSeeChanges)
    'Set rs = db.OpenRecordset("SELECT * FROM columns WHERE user_id = '" & GetUserName() & "' AND table_id = '" & rfrm.name & "' AND column_id = '" & clmn & "'", dbOpenDynaset, dbSeeChanges)
    
    If rs.EOF Or rs.BOF Then
        cs.Width = -3
        GoTo Exit_Function
    Else
        cs.Width = rs!column_width
        cs.hidden = rs!column_hidden
        cs.order = rs!column_order
    End If

Exit_Function:
    GetColumnStyle = cs
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    cs.Width = -3
    Resume Exit_Function
End Function

Public Function GetFieldValue(ByVal table As String, ByVal field As String, Optional ByVal ID As Long) As Variant
'-------------------------------------------------------------------------------
' Function:         GetFieldValue
' Date:             2022 Feburary
' Purpose:          Get value of a field
' In:
' -> table:         Table to open
' -> field;         Field to look at
' Out:              Response (Variant)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    If Nz(ID, 0) = 0 Then
        Set rs = db.OpenRecordset("SELECT TOP 1 " & field & " FROM " & table)
    Else
        Set rs = db.OpenRecordset("SELECT " & field & " FROM " & table & " WHERE id = " & ID)
    End If
    
    If rs(0) <> "" Then
        GetFieldValue = rs(0)
    Else
        GetFieldValue = Null
    End If
    
Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    GetFieldValue = Null
    Resume Exit_Function
End Function

Public Function SetFieldValue(ByVal table As String, ByVal field As String, ByVal ID As Long, ByVal value As Variant) As Boolean
'-------------------------------------------------------------------------------
' Function:         SetFieldValue
' Date:             2022 Feburary
' Purpose:          Get value of a field
' In:
' -> table:         Table to open
' -> field;         Field to look at
' Out:              Response (Variant)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    
    If IsNumeric(value) Then
        db.Execute "UPDATE " & table & " SET " & field & " = " & value
    Else
        db.Execute "UPDATE " & table & " SET " & field & " = '" & value & "'"
    End If
    
    SetFieldValue = True
    
Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function GetFeVersion() As Variant
'-------------------------------------------------------------------------------
' Function:         GetFeVersion
' Date:             2024 February
' Purpose:          Get the frontend version
' Out:              Version
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    Dim cs As ColumnStyle
    
    Set db = CurrentDb()
    
    'Local column style switch off
    Set rs = db.OpenRecordset("SELECT TOP 1 version_fe FROM setup", dbOpenDynaset, dbSeeChanges)
    
    If IsNull(rs!version_fe) Then db.Execute "UPDATE setup SET version_fe = '" & config.FeVersion & "'"
    GetFeVersion = Nz(rs!version_fe, config.FeVersion)

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    GetFeVersion = ""
    Resume Exit_Function
End Function

Public Sub SetFeVersion()
'-------------------------------------------------------------------------------
' Function:         SetFeVersion
' Date:             2024 February
' Purpose:          Set the frontend version
' Out:              Version
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    
    Set rs = db.OpenRecordset("SELECT COUNT(*) As cnt FROM setup", dbOpenDynaset, dbSeeChanges)
    
    If Nz(rs!cnt, 0) = 0 Then
        db.Execute "INSERT INTO setup (version_fe) VALUES ('" & config.FeVersion & "')"
    Else
        db.Execute "UPDATE setup SET version_fe = '" & config.FeVersion & "'"
    End If

Exit_Function:
    Set db = Nothing
    Exit Sub
Catch_Error:
    Resume Exit_Function
End Sub

Public Function InstallationPending() As Boolean
'-------------------------------------------------------------------------------
' Function:         SetFeVersion
' Date:             2024 February
' Purpose:          Set the frontend version
' Out:              Version
'-------------------------------------------------------------------------------
  On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    
    Set rs = db.OpenRecordset("SELECT COUNT(*) As cnt FROM users", dbOpenDynaset, dbSeeChanges)
    
    If Nz(rs!cnt, 0) = 0 Then InstallationPending = True

Exit_Function:
    Set db = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function GetBeVersion() As Variant
'-------------------------------------------------------------------------------
' Function:         GetBeVersion
' Date:             2022 January
' Purpose:          Get the backend version
' Out:              Version
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "version_be"
    cmd.Parameters.Append cmd.CreateParameter("@version_be", adVarChar, adParamOutput, 256)
    cmd.Execute
    
    GetBeVersion = cmd.Parameters.item("@version_be")
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    GetBeVersion = ""
    Resume Exit_Function
End Function

Public Function GetUserName() As Variant
'-------------------------------------------------------------------------------
' Function:  GetUserName
' Date:      2022 January
' Purpose:   Get the name of the logged in user
' In:        -
' Out:       User name (will be null in case of any error)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    
    If pUserName <> "" Then GoTo Exit_Function
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "users_get_name"
    cmd.Parameters.Append cmd.CreateParameter("@response_message", adVarChar, adParamOutput, 256)
    cmd.Execute
    
    GetUserName = cmd.Parameters.item("@response_message")
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    GetUserName = Null
    Resume Exit_Function
End Function

Public Function GetUserId() As Variant
'-------------------------------------------------------------------------------
' Function:  GetUserId
' Date:      2022 March
' Purpose:   Get the id of the logged in user
' In:        -
' Out:       User id (will be null in case of any error)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    If pUserId <> 0 Then GoTo Exit_Function
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT id FROM users WHERE name = '" & GetUserName() & "'", dbOpenDynaset, dbSeeChanges)
    
    If rs.EOF Or rs.BOF Then
        GetUserId = Null
    Else
        GetUserId = rs(0)
    End If

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    GetUserId = Null
    Resume Exit_Function
End Function

Public Function GetTranslation(ByVal container As String, ByVal item As String, ByVal language As String) As Variant
'-------------------------------------------------------------------------------
' Function:         GetTranslation
' Date:             2022 January
' Purpose:          Get the translation of a specific item in a container
' In:
' -> container:     Container with item of interes (e.g. Form, MsgBox, . . .)
' -> item:          Item to be translated (e.g. name of the control)
' -> language:      Language code (e.g. 'de')
' Out:              Translation (will be null in case of any error)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT " & language & " FROM lcl_translation WHERE container ='" & container & "' AND item = '" & item & "'")
    
    If rs.EOF Or rs.BOF Then
        db.Execute "INSERT INTO translation(container, item) VALUES('" & container & "', '" & item & "')"
        GetTranslation = item
    Else
        GetTranslation = Nz(rs(0), item)
    End If

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    GetTranslation = item
    Resume Exit_Function
End Function

Public Function GetPermission(ByVal title As String) As CRUD
'-------------------------------------------------------------------------------
' Function:         GetPermission
' Date:             2022 January
' Purpose:          Get the permission of an item called title
' In:
' -> title:         Title of the permission
' Out:              Permission according CRUD taxonomy
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim response As CRUD
    Dim db As database
    Dim rsa, rsb, rsc As Recordset
    
    Set db = CurrentDb()
    Set rsa = db.OpenRecordset("SELECT id FROM permission WHERE title = '" & title & "'", dbOpenDynaset, dbSeeChanges)
    Set rsb = db.OpenRecordset("SELECT role FROM users WHERE name = '" & GetUserName & "'", dbOpenDynaset, dbSeeChanges)
    
    If rsa.EOF Or rsa.BOF Or rsb.EOF Or rsb.BOF Then
        GetPermission = response
        db.Execute "INSERT INTO permission(title) VALUES('" & title & "')"
        Exit Function
    End If
    
    Set rsc = db.OpenRecordset("SELECT * FROM role_permission WHERE role = " & rsb(0) & " AND permission = " & rsa(0), dbOpenDynaset, dbSeeChanges)
    
    response.Create = rsc("can_create")
    response.Read = rsc("can_read")
    response.Update = rsc("can_update")
    response.Delete = rsc("can_delete")
    
    GetPermission = response

Exit_Function:
    Set db = Nothing
    Set rsa = Nothing
    Set rsb = Nothing
    Set rsc = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function DuplicateRecord(ByVal table As String, ByVal ID As Long) As Boolean
'-------------------------------------------------------------------------------
' Function:         DuplicateRecord
' Date:             2022 January
' Purpose:          Perform the duplication of the record if applies
' In:
' -> rfrm:          Name of the table to duplicate a record
' Out:              Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    cmd.CommandType = adCmdStoredProc
    
    Select Case table
        Case "role"
            cmd.CommandText = "role_duplicate"
            cmd.Parameters.Append cmd.CreateParameter("@pRole", adInteger, adParamInput, , ID)
            cmd.Execute
        Case Else
            MsgBox GetTranslation("msgbox", "duplication_not_supported", GetDbSetting("language")), vbExclamation, GetTranslation("msgbox", "vbExclamation", GetDbSetting("language"))
    End Select
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (mdlDbProcedures - DuplicateRecord): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function TestCalculation(ByVal analysis As Long) As Double
'-------------------------------------------------------------------------------
' Function:         TestCalculation
' Date:             2022 January
' Purpose:          Validate equation by calling calculation test procedure
' In:
' -> analysis:      Analysis service of interest
' Out:              Result (Float)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "calculation_test"
    cmd.Parameters.Append cmd.CreateParameter("@analysis", adInteger, adParamInput, , analysis)
    cmd.Parameters.Append cmd.CreateParameter("@response_message", adDouble, adParamOutput)
    cmd.Execute
    
    TestCalculation = cmd.Parameters.item("@response_message")
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (mdlDbProcedures - TestCalculation): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function IterateCalculation(ByVal request As Integer) As Boolean
'-------------------------------------------------------------------------------
' Function:         TestCalculation
' Date:             2022 January
' Purpose:          Validate equation cy calculation test
' In:
' -> analysis:      Analysis service of interest
' Out:              Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "calculation_iterate"
    cmd.Parameters.Append cmd.CreateParameter("@request", adInteger, adParamInput, , request)
    cmd.Execute
    
    IterateCalculation = True
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (mdlDbProcedures - IterateCalculation): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function SendMail(ByVal recipients As String, ByVal subject As String, ByVal body As String) As Boolean
'-------------------------------------------------------------------------------
' Function:         SendMail
' Date:             2022 February
' Purpose:          Send a mail using the server mail profile
' In:
' -> recipients:    Recipients
' -> subject:       Subject
' -> body:          Body
' Out:              Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "mail_send"
    cmd.Parameters.Append cmd.CreateParameter("@p_recipients", adLongVarWChar, adParamInput, -1, recipients)
    cmd.Parameters.Append cmd.CreateParameter("@p_subject", adVarChar, adParamInput, 256, subject)
    cmd.Parameters.Append cmd.CreateParameter("@p_body", adLongVarWChar, adParamInput, -1, body)
    cmd.Parameters.Append cmd.CreateParameter("@p_filenames", adLongVarWChar, adParamInput, -1, "")
    cmd.Execute
    
    SendMail = True
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (mdlDbProcedures - SendMail): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function AttachToMailQueue(ByVal request As Variant, ByVal billing_customer As Variant, ByVal subject As String, ByVal body As String, Optional ByVal recipients As Variant) As Boolean
'-------------------------------------------------------------------------------
' Function:         AttachToMailQueue
' Date:             2022 March
' Purpose:          Attach a mail to queue
' In:
' -> request:       Request
' -> subject:       Subject
' -> body:          Body
' -> recipients:    Recipients (Optional)
' Out:              Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    
    If IsNull(recipients) Or recipients = "" Or IsNull(request) Then Err.Raise vbObjectError + 513, , GetTranslation("msgbox", "mailadress_error", GetDbSetting("language"))
    
    Set db = CurrentDb
    
    If recipients = "" Then
        db.Execute "INSERT INTO mailqueue (subject, body, request, billing_customer) VALUES('" & subject & "', '" & body & "', " & Nz(request, "NULL") & ", " & Nz(billing_customer, "NULL") & ")", dbOpenDynaset
    Else
        db.Execute "INSERT INTO mailqueue (recipients, subject, body, request, billing_customer) VALUES('" & recipients & "', '" & subject & "', '" & body & "', " & Nz(request, "NULL") & ", " & Nz(billing_customer, "NULL") & ")", dbOpenDynaset
    End If
    
    AttachToMailQueue = True

Exit_Function:
    Set db = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (mdlDbProcedures - AttachToMailQueue): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function CreateSubRequest(ByVal request As Integer) As Boolean
'-------------------------------------------------------------------------------
' Function:         CreateSubRequest
' Date:             2022 February
' Purpose:          Create a subrequest (sample)
' In:
' -> recipients:    Request of choice
' Out:              Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "request_create_subrequest"
    cmd.Parameters.Append cmd.CreateParameter("@p_id", adInteger, adParamInput, , request)
    cmd.Execute
    
    CreateSubRequest = True
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (mdlDbProcedures - CreateSubRequest): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function AddUser() As Boolean
'-------------------------------------------------------------------------------
' Function:         AddUser
' Date:             2022 March
' Purpose:          Add user to table users
' In:               -
' Out:              Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM users WHERE name = '" & GetUserName() & "'")
    
    If rs(0) = 0 Then db.Execute "INSERT INTO users (name) VALUES('" & GetUserName() & "')", dbOpenDynaset
    
     AddUser = True
     
Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function AddAdmin() As Boolean
'-------------------------------------------------------------------------------
' Function:         AddAdmin
' Date:             2022 March
' Purpose:          Add admin to table users
' In:               -
' Out:              Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    Dim rx As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT COUNT(*) FROM users WHERE name = '" & GetUserName() & "'", dbOpenDynaset, dbSeeChanges)
    Set rx = db.OpenRecordset("SELECT id FROM role WHERE administrative = 1", dbOpenDynaset, dbSeeChanges)

    If rs(0) = 0 Then db.Execute "INSERT INTO users (name, role) VALUES('" & GetUserName() & "', " & rx!ID & ")", dbOpenDynaset
    
    AddAdmin = True
     
Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function CreateSPA(ByVal uid As String, ByVal profile As Integer, ByVal analysis As Integer, ByVal from As Date, ByVal till As Date) As Boolean
'-------------------------------------------------------------------------------
' Function:         CreateSPA
' Date:             2022 March
' Purpose:          Create a statistical process analysis (SPA)
' In:
' -> UID:           Unique identifier to present data of interest in charts
' -> profile:       The profile of choice to be analyzed
' -> analysis:      The analysis of choice to be plotted/calculated
' -> from:          The date from to analyze the data
' -> till:          The date till to analyze the data
' Out:              Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "spa_create"
    cmd.Parameters.Append cmd.CreateParameter("@uid", adVarChar, adParamInput, 256, uid)
    cmd.Parameters.Append cmd.CreateParameter("@profile", adInteger, adParamInput, , profile)
    cmd.Parameters.Append cmd.CreateParameter("@analysis", adInteger, adParamInput, , analysis)
    cmd.Parameters.Append cmd.CreateParameter("@from", adDate, adParamInput, , from)
    cmd.Parameters.Append cmd.CreateParameter("@till", adDate, adParamInput, , till)
    cmd.Execute
    
    CreateSPA = True
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function GetAuditTrail(ByVal table As String, ByVal ID As Long) As String
'-------------------------------------------------------------------------------
' Function:         etAuditTrail
' Date:             2022 March
' Purpose:          Get audit trail of table
' In:
' -> table:         Table of interest
' -> id:            id to diff
' -> from:          The date from to analyze the data
' -> till:          The date till to analyze the data
' Out:              Recordset
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim fld As field
    Dim s As String
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "audit_xml_diff"
    cmd.Parameters.Append cmd.CreateParameter("@table_name", adVarChar, adParamInput, 128, table)
    cmd.Parameters.Append cmd.CreateParameter("@id", adInteger, adParamInput, , ID)
    
    Set rs = cmd.Execute()
    
    While Not rs.EOF
        If Len(Nz(rs!value_old, "null")) < 255 And Len(Nz(rs!value_new, "null")) < 255 Then
            s = s & "[" & rs!changed_at & "] " & rs!changed_by & "<br>" & "<b>" & rs!elem_name & "</b> " & Nz(rs!value_old, "null") & " -> " & Nz(rs!value_new, "null") & "<br><br>"
        Else
            s = s & "[" & rs!changed_at & "] " & rs!changed_by & "<br>" & "<b>" & rs!elem_name & "</b> " & Nz(Left(rs!value_old, 255) & " . . .", "null") & " -> " & Nz(Left(rs!value_new, 255) & " . . .", "null") & "<br><br>"
        End If
        rs.MoveNext
    Wend
        
    GetAuditTrail = s

Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function GetStateCode(ByVal state As Integer) As Variant
'-------------------------------------------------------------------------------
' Function:         GetStateCode
' Date:             2022 February
' Purpose:          Get the state code from table state id
' Out:              State Code
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT state FROM state WHERE id = " & state)
    
    If rs.EOF Or rs.BOF Then
        GetStateCode = Null
    Else
        GetStateCode = rs(0)
    End If

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    GetStateCode = Null
    MsgBox "Error (mdlDbProcedures - GetVersion): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function GetCustomerID() As Variant
'-------------------------------------------------------------------------------
' Function:         GetCustomerID
' Date:             2022 April
' Purpose:          Get customer id of actual user
' Out:              Customer ID
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "users_get_customer"
    cmd.Parameters.Append cmd.CreateParameter("@response_message", adInteger, adParamOutput)
    cmd.Execute
    
    GetCustomerID = cmd.Parameters.item("@response_message")
    
    Set rs = cmd.Execute()

Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function ReplaceSql(ByVal message As String) As String
'-------------------------------------------------------------------------------
' Function:         PrepareSqlMessage
' Date:             2022 July
' Purpose:          Substitute SQL statement by value in string
' In:               Message with SQL statement to be processed
' Out:              Message
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "message_substitute_sql"
    cmd.Parameters.Append cmd.CreateParameter("@p_message", adLongVarWChar, adParamInput, -1, message)
    cmd.Parameters.Append cmd.CreateParameter("@return_message", adBSTR, adParamOutput, -1)
    cmd.Execute
    
    ReplaceSql = cmd.Parameters.item("@return_message")
    
    Set rs = cmd.Execute()

Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    'ReplaceSql = "Error rendering sticker. Check for missing informations."
    ReplaceSql = message
    Resume Exit_Function
End Function

Public Function PoolMeasurement(ByVal measurement As Integer) As Variant
'-------------------------------------------------------------------------------
' Function:         PoolMeasurement
' Date:             2022 January
' Purpose:          Pool measurements in batch
' In:
' -> Measurement:   ID of measurement to pool in batch
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "measurement_pool"
    cmd.Parameters.Append cmd.CreateParameter("@p_id", adInteger, adParamInput, , measurement)
    cmd.Execute
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (mdlDbProcedures - PoolMeasurement): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function SetTableFlag(ByVal username As String, ByVal tablename As String, ByVal tableid As Long) As Boolean
'-------------------------------------------------------------------------------
' Function:         SetTableFlag
' Date:             2022 November
' Purpose:          Set a flag that specific record in table is in use.
'                   Set alert to true if user needs to get notified.
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    
    Set db = CurrentDb()
    db.Execute "INSERT INTO tableflag (user_name, table_name, table_id) VALUES ('" & username & "', '" & tablename & "', " & tableid & ")"

    SetTableFlag = True

Exit_Function:
    Set db = Nothing
    Exit Function
Catch_Error:
    SetTableFlag = False
    Resume Exit_Function
End Function

Public Function GetTableFlag(ByVal tablename As String, ByVal tableid As Long) As Variant
'-------------------------------------------------------------------------------
' Function:         GetTableFlag
' Date:             2022 November
' Purpose:          Gets the flag of a specific record.
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT user_name FROM tableflag WHERE table_name = '" & tablename & "' AND table_id = " & tableid, dbOpenDynaset, dbSeeChanges)
    
    GetTableFlag = rs(0)

Exit_Function:
    Set db = Nothing
    Set rs = Nothing
    Exit Function
Catch_Error:
    GetTableFlag = Null
    Resume Exit_Function
End Function

Public Function RemoveTableFlag(ByVal tablename As String, ByVal tableid As Long) As Boolean
'-------------------------------------------------------------------------------
' Function:         RemoveTableFlag
' Date:             2022 November
' Purpose:          Remove the flag for table in use.
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    
    Set db = CurrentDb()
    db.Execute "DELETE FROM tableflag WHERE table_name = '" & tablename & "' AND table_id = " & tableid, dbSeeChanges

    RemoveTableFlag = True

Exit_Function:
    Set db = Nothing
    Exit Function
Catch_Error:
    RemoveTableFlag = False
    Resume Exit_Function
End Function

Public Function AddErrorLog(ByVal ID As Long, ByVal description As String) As Boolean
'-------------------------------------------------------------------------------
' Function:  AddErrorLog
' Date:      2022 November
' Purpose:   Add a log entry into table errorlog
' In:        error_id, error_description
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    
    Set db = CurrentDb()
    db.Execute "INSERT INTO errorlog (error_id, error_description) VALUES (" & ID & ", '" & description & "')"
    
    AddErrorLog = True

Exit_Function:
    Set db = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function GetReports() As Variant
'-------------------------------------------------------------------------------
' Function:  GetReports
' Date:      2023 January
' Purpose:   Get a list of all reports
'-------------------------------------------------------------------------------
    Dim rpt As AccessObject, db As Object, a() As String, i As Integer
    
    Set db = Application.CurrentProject
    
    'Count reports
    i = 0
    For Each rpt In db.AllReports
        If InStr(1, rpt.name, config.ReportId, vbTextCompare) = 1 Then i = i + 1
    Next rpt
    
    'Fill array with report names
    ReDim a(i)
    i = 1
    For Each rpt In db.AllReports
        If InStr(1, rpt.name, config.ReportId, vbTextCompare) = 1 Then
            a(i) = rpt.name
            i = i + 1
        End If
    Next rpt
    
    GetReports = a
End Function

Public Function UploadPicture(ByRef fref As Object, Optional FileName As String) As Boolean
'-------------------------------------------------------------------------------
'Function:          UploadPicture
'Date:              2022 March
'Purpose:           Uploads a picture to the database
'In:
' -> fref:          BLOB field to store the serialized data
'Out:               Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim oShell As Object
    Dim oFSO As Object
    Dim fname As String
    
    Set oShell = CreateObject("WScript.Shell")
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    If FileName = "" Then
        fname = Dialog.OpenFileDialog()
    Else
        fname = FileName
    End If
    
    If fname = "" Then Exit Function
    DbBlob.FileToBlob fname, fref
    
    UploadPicture = True
    
Exit_Function:
    Exit Function
Catch_Error:
    MsgBox "Error (dbProdures - UploadPicture): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function ShowPicture(ByRef fref As Object, ByRef pref As Object) As Boolean
'-------------------------------------------------------------------------------
'Function:          ShowPicture
'Date:              2022 March
'Purpose:           Show the picture of the BLOB
'In:
' -> fref:          BLOB field of the serialized data
' -> pref           Picture field to present the picture
'Out:               Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim oShell As Object
    Dim oFSO As Object
    Dim fpath As String
    Dim fname As String
    
    Set oShell = CreateObject("WScript.Shell")
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    
    fpath = oShell.ExpandEnvironmentStrings("%USERPROFILE%\") + config.CacheFolder
    If Not oFSO.FolderExists(fpath) Then MkDir fpath
    
    fname = fpath & "\" & "tmp_" & CreateGuid()
    If Not IsNull(fref) Then DbBlob.BlobToFile fname, fref
    pref.Picture = fname
    
    ShowPicture = True
    
Exit_Function:
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function ConcatRelated(strField As String, strTable As String, Optional strWhere As String, Optional strOrderBy As String, Optional strSeparator = ", ") As Variant
'-------------------------------------------------------------------------------
'Function:          ConcatRelated
'Date:              2023 April
'Purpose:           Concat results from related values
'In:
' -> strField       Name of the field
' -> strTable       Name of the table
' -> strWhere       Optional where clause
' -> strOrderBy     Optional order by clause
' -> strSeparatur   Separator to be used in output string
'Out:               Concatenated values as string
'-------------------------------------------------------------------------------
On Error GoTo Err_Handler
    Dim rs As DAO.Recordset         'Related records
    Dim rsMV As DAO.Recordset       'Multi-valued field recordset
    Dim strSQL As String            'SQL statement
    Dim strOut As String            'Output string to concatenate to.
    Dim lngLen As Long              'Length of string.
    Dim bIsMultiValue As Boolean    'Flag if strField is a multi-valued field.
    
    'Initialize to Null
    ConcatRelated = Null
    
    'Build SQL string, and get the records.
    strSQL = "SELECT " & strField & " FROM " & strTable
    If strWhere <> vbNullString Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    If strOrderBy <> vbNullString Then
        strSQL = strSQL & " ORDER BY " & strOrderBy
    End If
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset)
    'Determine if the requested field is multi-valued (Type is above 100.)
    bIsMultiValue = (rs(0).type > 100)
    
    'Loop through the matching records
    Do While Not rs.EOF
        If bIsMultiValue Then
            'For multi-valued field, loop through the values
            Set rsMV = rs(0).value
            Do While Not rsMV.EOF
                If Not IsNull(rsMV(0)) Then
                    strOut = strOut & rsMV(0) & strSeparator
                End If
                rsMV.MoveNext
            Loop
            Set rsMV = Nothing
        ElseIf Not IsNull(rs(0)) Then
            strOut = strOut & rs(0) & strSeparator
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    'Return the string without the trailing separator.
    lngLen = Len(strOut) - Len(strSeparator)
    If lngLen > 0 Then
        ConcatRelated = Left(strOut, lngLen)
    End If

Exit_Handler:
    'Clean up
    Set rsMV = Nothing
    Set rs = Nothing
    Exit Function

Err_Handler:
    Resume Exit_Handler
End Function

Public Function ResetLoginPassword() As Boolean
'-------------------------------------------------------------------------------
' Function:  ResetLoginPassword
' Date:      2022 November
' Purpose:   Reset login password
' In:        -
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    db.Execute "UPDATE dbsetup SET password = ''"
    
    ResetLoginPassword = True

Exit_Function:
    Set db = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function GetDbDate() As Date
'-------------------------------------------------------------------------------
' Function:  GetDbDate
' Date:      2023 October
' Purpose:   Get date from sql server
' Out:       Date
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "datetime_get"
    cmd.Parameters.Append cmd.CreateParameter("@response_message", adDate, adParamOutput, 256)
    cmd.Execute
    
    GetDbDate = cmd.Parameters.item("@response_message")
    
Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (mdlDbProcedures - GetDbDate): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function RunTemplate(ByVal template As Integer, ByVal priority As Integer, ByVal workflow As Integer) As Variant
'-------------------------------------------------------------------------------
' Function:         RunTemplate
' Date:             2024 May
' Purpose:          Run/Execute template and create sample with subsamples
' Out:              Customer ID
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "template_run"
    cmd.Parameters.Append cmd.CreateParameter("@template", adInteger, adParamInput, -1, template)
    cmd.Parameters.Append cmd.CreateParameter("@priority", adInteger, adParamInput, -1, priority)
    cmd.Parameters.Append cmd.CreateParameter("@workflow", adInteger, adParamInput, -1, workflow)
    cmd.Execute

Exit_Function:
    Set cmd = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (mdlDbProcedures - RunTemplate): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function
