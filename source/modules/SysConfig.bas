VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SysConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'####################################################################################################
' This module encompasses the configuration information of the application. This includes public and
' global constants.
'####################################################################################################

Option Compare Database
Option Explicit

' Database connection timeout in seconds
Private Const pConnectionTimeout = 3

' Application identifier
Private Const pAppId = "labdesk-ui"

' Application title shown as form title
Private Const pAppTitle = "LABDESK - Laboratory Information Management System"

' Background worker timer interval in milliseconds
Private Const pBackgroundWorkerInterval = 3 * 60000

' Semantic Versioning according to https://semver.org/
' MAJOR-MINOR-PATCH
Private Const pFeVersion = "v2.1.1" 'Presented frontend version
Private Const pBeVersion = "v2.2.2" 'Required backend version

' Identifiers used for reports and labels
Private Const pReportId = "RPT-" 'Identifier for selectable reports
Private Const pInvoiceId = "INV-" 'Identifier for selectable reports
Private Const pLabelId = "LBL-" 'Identifier for selectable labels
Private Const pWorksheetId = "WKS-" 'Identifier for selectable Worksheets

' Temporary cache folder (the local user folder will be used) .labdesk
Private pCacheFolder As String

Public Property Get ConnectionTimeout() As Integer
    ConnectionTimeout = pConnectionTimeout
End Property

Public Property Get AppId() As String
    AppId = pAppId
End Property

Public Property Get AppTitle() As String
    AppTitle = pAppTitle
End Property

Public Property Get BackgroundWorkerInterval()
    BackgroundWorkerInterval = pBackgroundWorkerInterval
End Property

Public Property Get FeVersion() As String
    FeVersion = pFeVersion
End Property

Public Property Get BeVersion() As String
    BeVersion = pBeVersion
End Property

Public Property Get ReportId() As String
    ReportId = pReportId
End Property

Public Property Get InvoiceId() As String
    InvoiceId = pInvoiceId
End Property

Public Property Get LabelId() As String
    LabelId = pLabelId
End Property

Public Property Get WorksheetId() As String
    WorksheetId = pWorksheetId
End Property

Public Property Get CacheFolder() As String
    CacheFolder = pCacheFolder
End Property

Public Property Get DSNLessTables() As Variant
    Dim a As Variant
    a = Array( _
    "filter", "translation", "customfield", "setup", _
    "contact", "laboratory", "customer", "manufacturer", "smppoint", "smpcontainer", "smpmatrix", "material", "service", _
    "cfield", "cvalidate", "condition", "uncertainty", "analysis", "attribute", "method_analysis", "method_smptype", "method", "qualification", "instrument_method", "department", "workplace", "instrument", "certificate", "attachment", "instype", "supplier", _
    "request", "request_customfield", "request_material", "request_service", "measurement_cfield", "measurement", "measurement_condition", "view_request_measurement", _
    "profile_analysis", "profile", "request_analysis", "step", "state", "workflow", "smptype", "smpcondition", "smppreservation", "priority", "technique", _
    "template", "template_profile", _
    "batch", "btcposition", "storage", "strposition", "handbook", _
    "project", "project_customfield", "project_member", "formulation", "component", "task", "task_workload", "view_task", "task_service", "task_material", _
    "view_measurement", "view_request_owner", "view_project_owner", _
    "users", "role", "role_permission", "permission", "traversal", "audit", "spa", "columns", _
    "view_labreport_details", "view_attachment_revision", "view_worksheet_details", _
    "billing", "billing_customer", "billing_position", "view_billing_position", _
    "mailqueue", _
    "tableflag", "errorlog" _
    )
    DSNLessTables = a
End Property

Public Sub Init(DemoMode As Boolean, CacheFolder As String)
    pCacheFolder = CacheFolder
End Sub

