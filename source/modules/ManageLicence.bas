Attribute VB_Name = "ManageLicence"
'################################################################################################
' This module manages user licences in table users
'################################################################################################

Option Compare Database
Option Explicit

Public Function GetUserUid(ByVal ID As Integer) As String
On Error GoTo Catch_Error
    Dim db As database
    Dim rs1 As Recordset
    Dim encryptedString As String
    
    Set db = CurrentDb()
    
    Set rs1 = db.OpenRecordset("SELECT uid FROM users WHERE id = " & ID, dbOpenDynaset, dbSeeChanges)
    
    If rs1.EOF Then Err.Raise vbObjectError + 513, , "User ID not found."
    
    GetUserUid = rs1(0).value
    
Exit_Function:
    Set db = Nothing
    Set rs1 = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (ManageLicence - GetUserUak): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function GetUserUak(ByVal ID As Integer) As String
On Error GoTo Catch_Error
    Dim db As database
    Dim rs1 As Recordset
    Dim encryptedString As String
    
    Set db = CurrentDb()
    
    Set rs1 = db.OpenRecordset("SELECT uid FROM users WHERE id = " & ID, dbOpenDynaset, dbSeeChanges)
    
    If rs1.EOF Then Err.Raise vbObjectError + 513, , "User ID not found."

    encryptedString = CipherAES.StoreEncryptAES(rs1(0), config.MasterKey & ID, 1)
    GetUserUak = encryptedString
    
Exit_Function:
    Set db = Nothing
    Set rs1 = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (ManageLicence - GetUserUak): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function ActivateUser(ByVal ID As Integer) As Boolean
On Error GoTo Catch_Error
    Dim db As database
    Dim rs1 As Recordset
    Dim rs2 As Recordset
    Dim encryptedString As String
    
    Set db = CurrentDb()
    
    'Write activation key
    Set rs1 = db.OpenRecordset("SELECT uid FROM users WHERE id = " & ID, dbOpenDynaset, dbSeeChanges)
    
    If rs1.EOF Then Err.Raise vbObjectError + 513, , "User ID not found. Activation failed."
    
    encryptedString = CipherAES.StoreEncryptAES(rs1(0), config.MasterKey & ID, 1)
    db.Execute "UPDATE users SET uak = """ & encryptedString & """ WHERE id = " & ID, dbSeeChanges
    
    'Validate activation key
    Set rs2 = db.OpenRecordset("SELECT uak FROM users WHERE id = " & ID, dbOpenDynaset, dbSeeChanges)
    
    If rs2.EOF Then Err.Raise vbObjectError + 513, , "Activation key not found. Activation failed"
    If CipherAES.RetrieveDecryptAES(rs2(0), config.MasterKey & ID, 1) <> rs1(0) Then Err.Raise vbObjectError + 513, , "Activation key not properly written. Activation failed"
    
    ActivateUser = True
    
Exit_Function:
    Set db = Nothing
    Set rs1 = Nothing
    Set rs2 = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (ManageLicence - ActivateUser): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function DeActivateUser(ByVal ID As Integer) As Boolean
On Error GoTo Catch_Error
    Dim db As database
    Dim rs1 As Recordset
    Dim encryptedString As String
    
    Set db = CurrentDb()
    
    'Write activation key
    Set rs1 = db.OpenRecordset("SELECT uid FROM users WHERE id = " & ID, dbOpenDynaset, dbSeeChanges)
    
    If rs1.EOF Then Err.Raise vbObjectError + 513, , "User ID not found. Activation failed."
    
    db.Execute "UPDATE users SET uak = Null WHERE id = " & ID, dbSeeChanges
    
    DeActivateUser = True
    
Exit_Function:
    Set db = Nothing
    Set rs1 = Nothing
    Exit Function
Catch_Error:
    MsgBox "Error (ManageLicence - DeActivateUser): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function UserHasLicence(ByVal ID As Integer) As Boolean
On Error GoTo Catch_Error
    Dim db As database
    Dim rs1 As Recordset
    Dim rs2 As Recordset
    Dim encryptedString As String
    
    Set db = CurrentDb()
    
    'Get activation key
    Set rs1 = db.OpenRecordset("SELECT uid FROM users WHERE id = " & ID, dbOpenDynaset, dbSeeChanges)
    
    If rs1.EOF Then Exit Function
    
    encryptedString = CipherAES.StoreEncryptAES(rs1(0), config.MasterKey & ID, 1)
    
    'Validate activation key
    Set rs2 = db.OpenRecordset("SELECT uak FROM users WHERE id = " & ID, dbOpenDynaset, dbSeeChanges)
    
    If rs2.EOF Or isnull(rs2(0)) Then Exit Function
    If CipherAES.RetrieveDecryptAES(rs2(0), config.MasterKey & ID, 1) <> rs1(0) Then Err.Raise vbObjectError + 513, , "Activation key not properly written. Activation failed"
    
    UserHasLicence = True
    
Exit_Function:
    Set db = Nothing
    Set rs1 = Nothing
    Set rs2 = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function GetLicenceKey(ByVal ID As Integer, ByVal uid As String) As String
    GetLicenceKey = CipherAES.StoreEncryptAES(uid, config.MasterKey & ID, 1)
End Function

Public Sub LicenceAllUsers()
    Dim rs As DAO.Recordset
    
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM users", dbOpenDynaset, dbSeeChanges)

    'Loop through all users and write licence key
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF = True
            rs.Edit
            rs!uak = GetLicenceKey(rs!ID, GetUserUid(rs!ID))
            rs.Update
            
            'Prompt licensed user
            SysCmd acSysCmdSetStatus, "User " & rs!ID & " licensed."
            
            rs.MoveNext
        Loop
    End If
    
    SysCmd acSysCmdSetStatus, "User: " & DbProcedures.GetUserName() & IIf(config.DemoMode, " (not licensed)", " (licensed)")
    
    rs.Close 'Close the recordset
    Set rs = Nothing 'Clean up
End Sub
