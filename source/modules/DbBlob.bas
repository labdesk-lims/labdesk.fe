Attribute VB_Name = "DbBlob"
'################################################################################################
' This module provides file dialogues, serializes and deserializes data and writes it to the sql
' database and back to the harddisc of choice.
'################################################################################################

Option Compare Database
Option Explicit
    
Public Function BlobToFile(ByVal strFile As String, ByRef Field As Object) As Long
'-------------------------------------------------------------------------------
'Function:          BlobToFile
'Date:              2021 October
'Purpose:           Show a File Save dialog
'In:
'-> strFile:        Path were data will be saved to
'-> field:          Recordset field with serialized data
'Out:               Number of files processed
'-------------------------------------------------------------------------------
On Error GoTo Err_BlobToFile
    Dim nFileNum As Integer
    Dim abytData() As Byte
    BlobToFile = 0
    nFileNum = FreeFile

    Open strFile For Binary Access Write As nFileNum
    abytData = Field
    Put #nFileNum, , abytData
    BlobToFile = LOF(nFileNum)

Exit_BlobToFile:
    If nFileNum > 0 Then Close nFileNum
    Exit Function

Err_BlobToFile:
    MsgBox "Function BlobToFile Error ID(" & Err.Number & "): " & Err.description & "(" & strFile & ")", vbInformation
    BlobToFile = 0
    Resume Exit_BlobToFile
End Function

Public Function FileToBlob(ByVal strFile As String, ByRef Field As Object) As Boolean
'-------------------------------------------------------------------------------
'Function:          FileToBlob
'Date:              2021 October
'Purpose:           Show a File Save dialog
'In:
'-> strFile:        Path of the file to be serialized
'-> field:          Recordset field to store serialized data
'Out:               Done (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Err_FileToBlob

    'Test to see if the file exists. Exit if it does not.
    If Dir(strFile) = "" Then Exit Function

    FileToBlob = True

    'Create a connection object
    Dim cn As ADODB.Connection
    Set cn = CurrentProject.Connection

    'Create our other variables
    Dim rs As ADODB.Recordset
    Dim mstream As ADODB.stream
    Set rs = New ADODB.Recordset

    'Open our Binary Stream object and load our file into it
    Set mstream = New ADODB.stream
    mstream.Open
    mstream.type = adTypeBinary
    mstream.LoadFromFile strFile

    'read our binary file into the OLE Field
    Field = mstream.Read

    'Edit: Removed some cleanup code I had inadvertently left here.

CleanUp:
    On Error Resume Next
    rs.Close
    mstream.Close
    Set mstream = Nothing
    Set rs = Nothing
    Set cn = Nothing

    Exit Function

Err_FileToBlob:
    MsgBox "Function FileToBlob Error ID(" & Err.Number & "): " & Err.description, vbInformation
    FileToBlob = False
    Resume CleanUp
End Function




