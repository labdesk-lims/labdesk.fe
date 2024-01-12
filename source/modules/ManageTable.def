Attribute VB_Name = "ManageTable"
'################################################################################################
' This module supports asynchronous data exchange with forms. Data selected from the sql server
' will be copied to local tables and any change will be written back afterwards. Because
' of efficiency reasons a 'where clause' is used to select relevant data only. Binary data is
' not copied over to local table for performance reasons. Downloads need to handled directly from
' from the corresponding source table.
'################################################################################################

Option Compare Database
Option Explicit

Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Function CreateGuid() As String
'-------------------------------------------------------------------------------
'Function:  CreateGuid
'Date:      2021 October
'Purpose:   Creates a GUID as string
'Out:       GUID
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim udtGUID As GUID
    If (CoCreateGuid(udtGUID) = 0) Then
        CreateGuid = _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & String(4 - Len(Hex$(udtGUID.Data2)), "0") & _
        Hex$(udtGUID.Data2) & String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & IIf((udtGUID.Data4(1) < &H10), "0", "") & _
        Hex$(udtGUID.Data4(1)) & IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & IIf((udtGUID.Data4(4) < &H10), "0", "") & _
        Hex$(udtGUID.Data4(4)) & IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & IIf((udtGUID.Data4(7) < &H10), "0", "") & _
        Hex$(udtGUID.Data4(7))
    End If
    
Exit_Function:
    Exit Function
Catch_Error:
    MsgBox "Error (mdlAsyncTable - CreateGuid): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function GetIdentityColumn(ByVal table As String) As String
'-------------------------------------------------------------------------------
'Function:  GetIdentityColumn
'Date:      2021 October
'Purpose:   Checks which column is the identitiy column and returns it as string
'In:        Table for inspection
'Out:       Name of identity column
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As DAO.database, tbl As DAO.TableDef
    Dim f As DAO.field
    Dim id_clmn As String
    
    Set db = CurrentDb()
    
    For Each tbl In db.TableDefs
        If tbl.name = table Then
            For Each f In tbl.Fields
                If f.attributes = 17 And f.type = 4 Then    '(17 = dbAutoIncrField (16) + dbFixedField(1),  4 = dbLong)
                    id_clmn = f.name
                End If
            Next f
        End If
    Next tbl
    
    GetIdentityColumn = id_clmn
    
Exit_Function:
    Exit Function
Catch_Error:
    MsgBox "Error (mdlAsyncTable - GetIdentityColumn): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function TableExists(ByVal table As String) As Boolean
'-------------------------------------------------------------------------------
'Function:  TableExists
'Date:      2021 October
'Purpose:   Checks if the table exists
'In:        Table to search for
'Out:       Found (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim tbl
    tbl = DCount("*", table)
    
    TableExists = True
    
Exit_Function:
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function CleanTable(ByVal table As String) As Boolean
'-------------------------------------------------------------------------------
'Function:  Clean Table
'Date:      2021 October
'Purpose:   Delete all rows in a table
'In:        Table for cleaning
'Out:       Cleaned(T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    CurrentDb.Execute "DELETE * FROM " & table
    
    CleanTable = True
    
Exit_Function:
    Exit Function
Catch_Error:
    MsgBox "Error (mdlAsyncTable - CleanTable): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function CreateTable(ByVal table As String, ByVal tableCopy As String) As Boolean
'-------------------------------------------------------------------------------
'Function:      CreateTable
'Date:          2021 October
'Purpose:       Creates a temporary table from a referenced table
'Parameters:
'-> table:      Table which will be the template
'-> tableCopy:  Table which will be created based on the template
'Out:           Created (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database
    Dim tdf As TableDef, fldNew As field
    Dim fldOld As field, rst As Recordset
    Dim fldPrp As DAO.Property
    
    Set db = CurrentDb
    Set rst = CurrentDb.OpenRecordset("Select * FROM " & table, dbOpenDynaset, dbSeeChanges, dbOptimistic)
    'Create new table
    Set tdf = db.CreateTableDef(tableCopy)
    
    'Copy record layout to new table
    For Each fldOld In rst.Fields
        Set fldNew = tdf.CreateField(fldOld.name, fldOld.type, fldOld.Size)
        tdf.Fields.Append fldNew
    Next fldOld
    
    'Append the table to the tabledefs collection
    db.TableDefs.Append tdf

    CreateTable = True

Exit_Function:
    Exit Function
Catch_Error:
    MsgBox "Error (mdlAsyncTable - CreateTable): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function CopyTable(ByVal table As String, ByVal tableCopy As String, Optional ByVal whereClause As String) As Boolean
'-------------------------------------------------------------------------------
'Function:          CopyTable
'Date:              2021 October
'Purpose:           Copy data from one table to another temporary table
'Parameters:
'-> table:          Table which will be copied over
'-> tableCopy:      Table where data will be copied to
'-> whereClause:    Optional where clause to cur down data copied
'Out:               Copied (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database, rstTable As Recordset, rstTemp As Recordset
    Dim lngi As Long
    
    Set db = CurrentDb
    
    'Clear old data from the temp table
    If TableExists(tableCopy) Then
        CleanTable tableCopy
    Else
        CreateTable table, tableCopy
    End If
    
    'Copy data from table to copyTable
    db.Execute "INSERT INTO " & tableCopy & " SELECT * FROM " & table & " " & whereClause, dbSeeChanges
    
    CopyTable = True

Exit_Function:
    Exit Function
Catch_Error:
    MsgBox "Error (mdlAsyncTable - CopyTable): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function UpdateTable(ByVal table As String, ByVal tableCopy As String) As Boolean
'-------------------------------------------------------------------------------
'Function:      UpdateTable
'Date:          2021 October
'Purpose:       Updates the source from a temporary table
'Parameters:
'-> table:      Table which will be updated in the source
'-> tableCopy:  Table where data will taken from to copy over
'Out:           Updated (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim db As database, rstForm As Recordset, rstTemp As Recordset
    Dim fld As field
    Dim id_column As String
    
    Set db = CurrentDb

    Set rstTemp = CurrentDb.OpenRecordset("Select * FROM " & tableCopy & " WHERE id <> NULL", dbOpenDynaset, dbSeeChanges, dbOptimistic)
    id_column = GetIdentityColumn(table)

    If Not (rstTemp.BOF And rstTemp.EOF) Then
        'Update all of the form records which have been edited
        rstTemp.MoveFirst
        Do Until rstTemp.EOF
            Set rstForm = CurrentDb.OpenRecordset("Select * FROM " & table & " WHERE id = " & rstTemp(id_column).value, dbOpenDynaset, dbSeeChanges, dbOptimistic)
            If Not rstForm.EOF Then
                rstForm.Edit
                For Each fld In rstTemp.Fields
                    If fld.name <> id_column Then
                        If fld.type <> dbLongBinary Then
                            rstForm(fld.name).value = fld
                        Else
                            If Not IsNull(fld) Then rstForm(fld.name).value = fld
                        End If
                    End If
                Next fld
                rstForm.Update
            End If
            rstTemp.MoveNext
        Loop
    End If
    
    Set rstTemp = CurrentDb.OpenRecordset("Select * FROM " & tableCopy & " WHERE id IS NULL", dbOpenDynaset, dbSeeChanges, dbOptimistic)
    Set rstForm = db.OpenRecordset(table, dbOpenDynaset, dbSeeChanges, dbOptimistic)
    id_column = GetIdentityColumn(table)

    If Not (rstTemp.BOF And rstTemp.EOF) Then
        'Add all of the form records which have been added
        rstTemp.MoveFirst
        Do Until rstTemp.EOF
            rstForm.AddNew
            For Each fld In rstTemp.Fields
                If fld.name <> id_column Then
                    rstForm(fld.name).value = fld
                End If
            Next fld
        rstForm.Update
        rstTemp.MoveNext
        Loop
    End If
    
    Set rstTemp = CurrentDb.OpenRecordset("Select " & tableCopy & "_cpy" & ".id" & " FROM " & tableCopy & "_cpy" & " LEFT JOIN " & tableCopy & " ON " & tableCopy & "_cpy" & ".id = " & tableCopy & ".id WHERE " & tableCopy & ".id Is Null", dbOpenDynaset, dbSeeChanges, dbOptimistic)
    If Not (rstTemp.EOF) Then
        'Delete all of the form records which have been deleted
        rstTemp.MoveFirst
        Do Until rstTemp.EOF
            CurrentDb.Execute "DELETE FROM " & table & " WHERE id = " & rstTemp("id").value, dbSeeChanges
            rstTemp.MoveNext
        Loop
    End If
    
    UpdateTable = True

Exit_Function:
    Exit Function
Catch_Error:
    MsgBox "Error (mdlAsyncTable - UpdateTable): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function

Public Function CreateMultiReportTable(ByVal table As String, ByVal profile As Long, ByVal smppoint As Long, ByVal top As Long, ByVal startFrom As Long) As Recordset
'-------------------------------------------------------------------------------
'Function:      CreateMultiReportTable
'Date:          2023 December
'Purpose:       Create a RecordSet for a multi sample report
'Parameters:
'-> table:      Name of the table where data will be stored
'-> profile:    The applied profile for each sample to be displayed
'-> smppoint:   The sample point of interest
'-> top:        Number of samples to be taken into account
'-> startFrom:  The smalles uid to start from
'Out:           Updated (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim rx As Recordset
    Dim i As Long
    Dim tdf As TableDef
    Dim fld As field
    
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = ADODBConStr
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "report_multiple"
    cmd.Parameters.Append cmd.CreateParameter("@profile", adInteger, adParamInput, , Nz(profile, 0))
    cmd.Parameters.Append cmd.CreateParameter("@smppoint", adInteger, adParamInput, , Nz(smppoint, 0))
    cmd.Parameters.Append cmd.CreateParameter("@top", adInteger, adParamInput, , top)
    cmd.Parameters.Append cmd.CreateParameter("@from", adInteger, adParamInput, , startFrom)
    
    Set rs = cmd.Execute
    
    If TableExists(table) Then CurrentDb.TableDefs.Delete table
    
    Set tdf = CurrentDb.CreateTableDef(table)
    
    ' Create columns
    For i = 0 To rs.Fields.count - 1
        Set fld = tdf.CreateField(rs.Fields(i).name, dbText, 255)
        tdf.Fields.Append fld
    Next i
    
    CurrentDb.TableDefs.Append tdf
    
    ' Update data
    Set rx = CurrentDb.OpenRecordset(table, dbOpenDynaset)
    rs.MoveFirst
    Do While Not rs.EOF
        rx.AddNew
        For i = 0 To rs.Fields.count - 1
            rx(i) = rs(i)
        Next i
        rs.MoveNext
        rx.Update
    Loop
    
    Set CreateMultiReportTable = rx
    
Exit_Function:
    Set cmd = Nothing
    Set rs = Nothing
    Set rx = Nothing
    Exit Function
Catch_Error:
    'MsgBox "Error (mdlDbProcedures - GetMultiReport): " & Err.description, vbCritical, "Error"
    Resume Exit_Function
End Function
