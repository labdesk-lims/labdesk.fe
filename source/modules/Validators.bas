Attribute VB_Name = "Validators"
 '################################################################################################
' Here you will find functions to validate data of fields
'################################################################################################

Option Compare Database
Option Explicit

Public Function IsEqual(value As Variant, comparison As Variant) As Boolean
    If value = comparison Then IsEqual = True Else IsEqual = False
End Function

Public Function ValidEmailAddress(ByVal strEmailAddress As String) As Boolean
'-------------------------------------------------------------------------------
'Function:          ValidateEmailAddress
'Date:              2021 October
'Purpose:           Validate email address
'In:                Email address as string
'Out:               Valid (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    
    Dim objRegExp As New RegExp
    Dim blnIsValidEmail As Boolean
    
    If IsNull(strEmailAddress) Then Resume Exit_Function
    
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = "^([a-zA-Z0-9_\-\.]+)@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,3})$"
    
    blnIsValidEmail = objRegExp.test(strEmailAddress)
    ValidEmailAddress = blnIsValidEmail
      
Exit_Function:
    Set objRegExp = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function ValidEmailArray(ByVal strEmailArray As String) As Boolean
'-------------------------------------------------------------------------------
'Function:          ValidateEmailArray
'Date:              2021 October
'Purpose:           Validate email array (like: a@b.de;c@de.en)
'In:                Email array as string
'Out:               Valid (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim cc_email As Variant
    Dim s As Variant
    Dim b As Boolean
    
    cc_email = Split(strEmailArray, ";")
    
    b = True
    For Each s In cc_email
        b = b And ValidEmailAddress(s)
    Next

    ValidEmailArray = b
    
Exit_Function:
    Exit Function
 
Catch_Error:
    Resume Exit_Function
End Function

Public Function IsValidFolderName(ByVal sFolderName As String) As Boolean
'-------------------------------------------------------------------------------
'Function:          IsValidFileNameOrPath
'Date:              2021 October
'Purpose:           Validate folder name
'In:                Folder name
'Out:               Valid (T/F)
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    Dim objRegExp As New RegExp
    
    'Check to see if any illegal characters have been used
    objRegExp.Pattern = "[&lt;&gt;:""/\\\|\?\*]"
    IsValidFolderName = objRegExp.test(sFolderName)
    
    'Ensure the folder name does end with a . or a blank space
    If Right(sFolderName, 1) = "." Then IsValidFolderName = False
    If Right(sFolderName, 1) = " " Then IsValidFolderName = False
    
Exit_Function:
    Set objRegExp = Nothing
    Exit Function
Catch_Error:
    Resume Exit_Function
End Function

Public Function IsSubForm(ByRef rfrm As Form) As Boolean
'-------------------------------------------------------------------------------
'Function:          IsSubForm
'Date:              2024 February
'Purpose:           Check if form is a subform
'rfrm:              name of the form to check
'Out:               T/F
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    IsSubForm = Not (rfrm.Parent Is Nothing)
    Exit Function
Catch_Error:
    IsSubForm = False
End Function

Public Function IsFormView(ByRef rfrm As Form) As Boolean
'-------------------------------------------------------------------------------
'Function:          IsFormView
'Date:              2024 February
'Purpose:           Check form view
'rfrm:              name of the form to check
'Out:               T/F
'-------------------------------------------------------------------------------
On Error GoTo Catch_Error
    IsFormView = (rfrm.CurrentView = 1)
    Exit Function
Catch_Error:
    IsFormView = False
End Function

Public Sub ComboBoxSearch(ByRef combo As comboBox, ByVal lookupField As String, ByVal pk As String)
'-------------------------------------------------------------------------------
'Function:          ComboBoxSearch
'Date:              2023 May
'Purpose:           Google like search for combobox
'In:
'-> combo           ComboBox object
'-> lookupField     Field to be selected for combobox
'-> pk              Primary key of table
'Out:               Valid (T/F)
'-------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = combo.RowSource
    combo.RowSource = "SELECT " & lookupField & ", " & pk & " FROM (" & Replace(strSQL, ";", "") & ") WHERE " & lookupField & " Like '*" & combo.Text & "*'"
    'combo.Dropdown '<- activate to open dropdown combobox automatically
End Sub

Public Sub EncryptTest()
    MsgBox RetrieveDecryptAES(StoreEncryptAES("test", "123", 1), "123", 1, True)
End Sub
