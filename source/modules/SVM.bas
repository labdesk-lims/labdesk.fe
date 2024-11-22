Attribute VB_Name = "Svm"
'################################################################################################
' This class exports and imports all source code, properties and references.
' To properly import reference Microsoft.Scripting and Microsoft.XML need to be activated
'################################################################################################

Option Compare Database
Option Private Module
Option Explicit

Private Const VB_MODULE As Integer = 1
Private Const VB_CLASS As Integer = 2
Private Const VB_FORM As Integer = 100
Private Const EXT_TABLE = ".tbl"
Private Const EXT_QUERY = ".qry"
Private Const EXT_FORM = ".frm"
Private Const EXT_REPORT = ".rpt"
Private Const EXT_MACRO = ".bas"
Private Const EXT_MODULE = ".bas"
Private Const EXT_PROPERTY = ".prp"
Private Const EXT_REFERENCE = ".ref"
Private Const SRC_FLD As String = "source"
 
Public Sub Export()
'-------------------------------------------------------------------------------
' Function:         Export
' Date:             2023 December
' Purpose:          Exports all code as separate files to the folder $SRC_FLD
' Out:              -
'-------------------------------------------------------------------------------
    Dim obj As Object
    Dim fso As Object
    Dim strPath As String
    Dim strFileName As String
    
    Set fso = CreateObject("Scripting.FilesystemObject")
    
    SysCmd acSysCmdSetStatus, "Delete dated files . . ."
    deleteFolder fso.GetParentFolderName(CurrentProject.path), SRC_FLD
    strPath = addFolder(fso.GetParentFolderName(CurrentProject.path), SRC_FLD)
    
    'Tables
    SysCmd acSysCmdSetStatus, "Export tables . . ."
    For Each obj In CurrentDb.TableDefs
        If Left(obj.name, 4) <> "MSys" Then
            SysCmd acSysCmdSetStatus, "Export table " & obj.name
            strFileName = addFolder(strPath, "tables") & "\" & obj.name & EXT_TABLE
            Application.ExportXML acExportTable, obj.name, strFileName, strFileName & ".XSD", strFileName & ".XSL", , acUTF8, acEmbedSchema + acExportAllTableAndFieldProperties
        End If
    Next
    
    'Queries
    SysCmd acSysCmdSetStatus, "Export queries . . ."
    For Each obj In CurrentDb.QueryDefs
        If Left(obj.name, 1) <> "~" Then
            SysCmd acSysCmdSetStatus, "Export query " & obj.name
            strFileName = addFolder(strPath, "queries") & "\" & obj.name & EXT_QUERY
            Application.SaveAsText acQuery, obj.name, strFileName
        End If
    Next
    
    'Forms
    SysCmd acSysCmdSetStatus, "Export forms . . ."
    For Each obj In CurrentProject.AllForms
        SysCmd acSysCmdSetStatus, "Export form " & obj.name
        strFileName = addFolder(strPath, "forms") & "\" & obj.name & EXT_FORM
        Application.SaveAsText acForm, obj.name, strFileName
    Next
    
    'Reports
    SysCmd acSysCmdSetStatus, "Export reports . . ."
    For Each obj In CurrentProject.AllReports
        SysCmd acSysCmdSetStatus, "Export report " & obj.name
        strFileName = addFolder(strPath, "reports") & "\" & obj.name & EXT_REPORT
        Application.SaveAsText acReport, obj.name, strFileName
    Next

    'Macros
    SysCmd acSysCmdSetStatus, "Export macros . . ."
    For Each obj In CurrentProject.AllMacros
        SysCmd acSysCmdSetStatus, "Export macro " & obj.name
        strFileName = addFolder(strPath, "macros") & "\" & obj.name & EXT_MACRO
        Application.SaveAsText acMacro, obj.name, strFileName
    Next
    
    'Modules
    SysCmd acSysCmdSetStatus, "Export modules . . ."
    For Each obj In Application.VBE.ActiveVBProject.VBComponents
        SysCmd acSysCmdSetStatus, "Export  module " & obj.name
        strFileName = addFolder(strPath, "modules") & "\" & obj.name & EXT_MODULE
        Select Case obj.type
            Case VB_MODULE
                obj.Export strFileName
            Case VB_CLASS
                obj.Export strFileName
            Case VB_FORM
                ' Do not export form modules (already exported the complete forms)
            Case Else
                Debug.Print "Unknown module type: " & obj.type, obj.name
        End Select
    Next
    
    'Properties
    SysCmd acSysCmdSetStatus, "Export properties . . ."
    strFileName = addFolder(strPath, "properties") & "\" & "properties" & EXT_PROPERTY
    Application.SaveAsAXL acDatabaseProperties, CurrentProject.name, strFileName
    
    'References
    SysCmd acSysCmdSetStatus, "Export references . . ."
    strFileName = addFolder(strPath, "references") & "\" & "references" & EXT_REFERENCE
    ExportReferences strFileName
    
    SysCmd acSysCmdSetStatus, "Export successfully finished!"
End Sub

Public Sub Import()
'-------------------------------------------------------------------------------
' Function:         Import
' Date:             2023 December
' Purpose:          Imports all code as separate files to the folder $SRC_FLD
' Note:             For proper import all tables need to be attached dto atabase
' Out:              -
'-------------------------------------------------------------------------------
    Dim obj As Object
    Dim fso As Object
    Dim oFile As File
    Dim strPath As String
    Dim strFileName As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    strPath = fso.GetParentFolderName(CurrentProject.path) + "\" + SRC_FLD
    
    'Tables
    SysCmd acSysCmdSetStatus, "Import tables . . ."
    For Each oFile In fso.GetFolder(strPath & "\tables").files
        SysCmd acSysCmdSetStatus, "Import table " & oFile.path
        If fso.GetExtensionName(oFile.path) = Replace(EXT_TABLE, ".", "") Then
            DeleteObject acTable, fso.GetBaseName(oFile.path)
            Application.ImportXML oFile.path, acStructureAndData
        End If
    Next
    
    'Queries
    SysCmd acSysCmdSetStatus, "Import queries . . ."
     For Each oFile In fso.GetFolder(strPath & "\queries").files
        SysCmd acSysCmdSetStatus, "Import query " & oFile.path
        If fso.GetExtensionName(oFile.path) = Replace(EXT_QUERY, ".", "") Then
            DeleteObject acQuery, fso.GetBaseName(oFile.path)
            Application.LoadFromText acQuery, fso.GetBaseName(oFile.path), oFile.path
        End If
    Next
    
    'Forms
    SysCmd acSysCmdSetStatus, "Import forms . . ."
     For Each oFile In fso.GetFolder(strPath & "\forms").files
        SysCmd acSysCmdSetStatus, "Import form " & oFile.path
        If fso.GetExtensionName(oFile.path) = Replace(EXT_FORM, ".", "") Then
            DeleteObject acForm, fso.GetBaseName(oFile.path)
            Application.LoadFromText acForm, fso.GetBaseName(oFile.path), oFile.path
        End If
    Next
    
    'Reports
    SysCmd acSysCmdSetStatus, "Import reports . . ."
    For Each oFile In fso.GetFolder(strPath & "\reports").files
        SysCmd acSysCmdSetStatus, "Import report " & oFile.path
        If fso.GetExtensionName(oFile.path) = Replace(EXT_REPORT, ".", "") Then
            DeleteObject acReport, fso.GetBaseName(oFile.path)
            Application.LoadFromText acReport, fso.GetBaseName(oFile.path), oFile.path
        End If
    Next
    
    'Macros
    SysCmd acSysCmdSetStatus, "Import macros . . ."
    For Each oFile In fso.GetFolder(strPath & "\macros").files
        SysCmd acSysCmdSetStatus, "Import macro " & oFile.path
        If fso.GetExtensionName(oFile.path) = Replace(EXT_MACRO, ".", "") Then
            DeleteObject acMacro, fso.GetBaseName(oFile.path)
            Application.LoadFromText acMacro, fso.GetBaseName(oFile.path), oFile.path
        End If
    Next
    
    'Modules
    SysCmd acSysCmdSetStatus, "Import modules . . ."
    For Each oFile In fso.GetFolder(strPath & "\modules").files
        SysCmd acSysCmdSetStatus, "Import module " & oFile.path
        If fso.GetExtensionName(oFile.path) = Replace(EXT_MODULE, ".", "") And fso.GetBaseName(oFile.path) <> "Svm" Then
            DeleteObject acModule, fso.GetBaseName(oFile.path)
            Application.VBE.ActiveVBProject.VBComponents.Import oFile.path
        End If
    Next
    
    'Properties
    SysCmd acSysCmdSetStatus, "Import properties . . ."
    For Each oFile In fso.GetFolder(strPath & "\properties").files
        If fso.GetExtensionName(oFile.path) = Replace(EXT_PROPERTY, ".", "") Then
            Application.LoadFromAXL acDatabaseProperties, fso.GetBaseName(oFile.path), oFile.path
        End If
    Next
    
    'References
    SysCmd acSysCmdSetStatus, "Import references . . ."
    For Each oFile In fso.GetFolder(strPath & "\references").files
        If fso.GetExtensionName(oFile.path) = Replace(EXT_REFERENCE, ".", "") Then
            ImportReferences oFile.path
        End If
    Next
    
    SysCmd acSysCmdSetStatus, "Import successfully finished!"
End Sub

Private Function DeleteObject(acObject As AcObjectType, name As String) As Boolean
On Error GoTo Skip
    DoCmd.DeleteObject acObject, name
    DeleteObject = True
    Exit Function
Skip:
    DeleteObject = False
    Exit Function
End Function

Private Function addFolder(ByVal strPath As String, ByVal strFolder As String) As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    addFolder = strPath & "\" & strFolder
    If Not fso.FolderExists(addFolder) Then MkDir addFolder
End Function

Public Function deleteFolder(ByVal strPath As String, ByVal strFolder As String) As String
On Error GoTo Skip
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    deleteFolder = strPath & "\" & strFolder
    If Not fso.FolderExists(deleteFolder) Then Exit Function
    fso.deleteFolder deleteFolder
Skip:
End Function

Private Function AddFormat(objXML As DOMDocument, strEncoding As String)
    Dim objPI As IXMLDOMProcessingInstruction
    
    Set objPI = objXML.createProcessingInstruction("xml", "version='1.0' encoding='" & strEncoding & "'")
    objXML.appendChild objPI
End Function

Private Function AppendElement(name As String, value As String, objRoot As IXMLDOMElement, objXML As DOMDocument)
    Dim elm As IXMLDOMElement
    
    Set elm = objXML.createElement(name)
    elm.Text = value
    objRoot.appendChild elm
End Function

Private Function ReplacePathWithEnviron(path As String) As String
    Dim str As String
    
    str = Replace(path, Environ$("SystemRoot"), "%SystemRoot%")
    str = Replace(str, Environ$("CommonProgramFiles"), "%CommonProgramFiles%")
    str = Replace(str, Environ$("ProgramFiles"), "%ProgramFiles%")
    str = Replace(str, CurrentProject.path, "%CurrentPath%")
    ReplacePathWithEnviron = str
End Function

Private Function ReplaceEnvironWithPath(path As String) As String
    Dim str As String
    
    str = Replace(path, "%SystemRoot%", Environ$("SystemRoot"))
    str = Replace(str, "%CommonProgramFiles%", Environ$("CommonProgramFiles"))
    str = Replace(str, "%ProgramFiles%", Environ$("ProgramFiles"))
    str = Replace(str, "%CurrentPath%", CurrentProject.path)
    ReplaceEnvironWithPath = str
End Function

Private Function ReferenceExists(ByVal fname As String) As Boolean
    Dim obj As Object
    For Each obj In Access.References
        If fname = obj.FullPath Then
            ReferenceExists = True
            Exit Function
        End If
    Next
    ReferenceExists = False
End Function

Private Sub ExportReferences(ByVal fname As String)
    Dim prj As VBProject
    Dim ref As VBIDE.Reference
    
    Dim objXML As DOMDocument
    Dim objPI As IXMLDOMProcessingInstruction
    
    Dim objReference As IXMLDOMElement

    Set objXML = New DOMDocument
    AddFormat objXML, "iso-8859-1"

    For Each prj In VBE.VBProjects
        Set objReference = objXML.createElement("Reference")
        
        For Each ref In prj.References
            AppendElement "Name", ref.name, objReference, objXML
            AppendElement "FullPath", ReplacePathWithEnviron(ref.FullPath), objReference, objXML
            AppendElement "Guid", ref.GUID, objReference, objXML
            AppendElement "Major", ref.Major, objReference, objXML
            AppendElement "Minor", ref.Minor, objReference, objXML
        Next ref
    
        objXML.appendChild objReference
        objXML.Save fname
    Next prj
End Sub

Private Function ImportReferences(ByVal fname As String) As Boolean
    Dim objXML As DOMDocument
    Dim objRoot As IXMLDOMElement
    Dim objL1 As Object
    Dim objL2 As Object
    
    Set objXML = New DOMDocument
    
    If Not objXML.Load(fname) Then Err.Raise vbObjectError + 513, , "Loading references failed."
    
    Set objRoot = objXML.DocumentElement
    
    For Each objL1 In objXML.DocumentElement.ChildNodes
        For Each objL2 In objL1.ChildNodes

            If objL1.nodeName = "FullPath" Then
                If Not ReferenceExists(ReplaceEnvironWithPath(objL2.Text)) Then Access.References.AddFromFile ReplaceEnvironWithPath(objL2.Text)
            End If
        Next
    Next
    
    ImportReferences = True
End Function
