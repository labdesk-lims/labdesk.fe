Attribute VB_Name = "ModBus"
Option Compare Database
Option Explicit

'-------------------------------------------------------------------------------
'
' This VB module is a collection of routines to perform serial port I/O without
' using the Microsoft Comm Control component.  This module uses the Windows API
' to perform the overlapped I/O operations necessary for serial communications.
'
' The routine can handle up to the number of serial ports defined by constant
' MAX_PORTS which are identified with a Port ID.
'
' All routines (with the exception of CommRead and CommWrite) return an error
' code or 0 if no error occurs.  The routine CommGetError can be used to get
' the complete error message.
'
'-------------------------------------------------------------------------------
' Public Constants
'-------------------------------------------------------------------------------

' Output Control Lines (CommSetLine)
Const LINE_BREAK = 1
Const LINE_DTR = 2
Const LINE_RTS = 3

' Input Control Lines  (CommGetLine)
Const LINE_CTS = &H10&
Const LINE_DSR = &H20&
Const LINE_RING = &H40&
Const LINE_RLSD = &H80&
Const LINE_CD = &H80&


' Constants for dwFlags of WINHTTP_AUTOPROXY_OPTIONS
Const WINHTTP_AUTOPROXY_AUTO_DETECT = 1
Const WINHTTP_AUTOPROXY_CONFIG_URL = 2

' Constants for dwAutoDetectFlags
Const WINHTTP_AUTO_DETECT_TYPE_DHCP = 1
Const WINHTTP_AUTO_DETECT_TYPE_DNS = 2

'-------------------------------------------------------------------------------
' System Constants
'-------------------------------------------------------------------------------
Private Const ERROR_IO_INCOMPLETE = 996&
Private Const ERROR_IO_PENDING = 997
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const OPEN_EXISTING = 3
Public Const HTTPREQUEST_PROXYSETTING_PROXY = 2

' COMM Functions
Private Const MS_CTS_ON = &H10&
Private Const MS_DSR_ON = &H20&
Private Const MS_RING_ON = &H40&
Private Const MS_RLSD_ON = &H80&
Private Const PURGE_RXABORT = &H2
Private Const PURGE_RXCLEAR = &H8
Private Const PURGE_TXABORT = &H1
Private Const PURGE_TXCLEAR = &H4

Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2

' COMM Escape Functions
Private Const CLRBREAK = 9
Private Const CLRDTR = 6
Private Const CLRRTS = 4
Private Const SETBREAK = 8
Private Const SETDTR = 5
Private Const SETRTS = 3

'-------------------------------------------------------------------------------
' Modbus Constants
'-------------------------------------------------------------------------------

Public Const MBJabberTimeout = 100    'streaming limit = gMBTimeout * MBJabberTimeout

'-------------------------------------------------------------------------------
' Global Variables
'-------------------------------------------------------------------------------

Public data_buffer(1024) As Byte
Public g_MB_errors As Integer
Public g_MB_buffer(256) As Long
Public gintStatusPointer As Long
Public lkfoiurop As String
Public gHalt As Integer
Public MBTimeoutMAX As Integer      'total timeout equals gMBTimeout * MBTimeoutMAX
Public gScanNowPushed As Boolean
Public gscreenUpdateState As Integer
Public gstatusBarState As Integer
Public gcalcState As Integer
Public geventsState As Integer
Public gdisplayPageBreakState As Integer
Public gCOMportStatus As String
Public dfhsdf12 As Integer
Public gCancel As Integer
Public gRefresh As Integer
Public gretBPS As Integer   'selected bit rate
Public gMBTimeout As Integer
Public goptContinuousScan As Integer

'**** test code
Public giReportRow As Long


'-------------------------------------------------------------------------------
' System Structures
'-------------------------------------------------------------------------------
Private Type COMSTAT
        fBitFields As Long ' See Comment in Win32API.Txt
        cbInQue As Long
        cbOutQue As Long
End Type

Private Type COMMTIMEOUTS
        ReadIntervalTimeout As Long
        ReadTotalTimeoutMultiplier As Long
        ReadTotalTimeoutConstant As Long
        WriteTotalTimeoutMultiplier As Long
        WriteTotalTimeoutConstant As Long
End Type

'
' The DCB structure defines the control setting for a serial
' communications device.
'
Private Type DCB
        DCBlength As Long
        BaudRate As Long
        fBitFields As Long ' See Comments in Win32API.Txt
        wReserved As Integer
        XonLim As Integer
        XoffLim As Integer
        ByteSize As Byte
        Parity As Byte
        StopBits As Byte
        XonChar As Byte
        XoffChar As Byte
        ErrorChar As Byte
        EofChar As Byte
        EvtChar As Byte
        wReserved1 As Integer 'Reserved; Do Not Use
End Type

Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

' My app proxy information type
Public Type yoasu20lsk
        active As Boolean
        proxy As String
        proxyBypass As String
End Type

' Structure to receive IE proxy settings
Public Type WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
        fAutoDetect As Long
        lpszAutoConfigUrl As Long
        lpszProxy As Long
        lpszProxyBypass As Long
End Type

Public Type WINHTTP_AUTOPROXY_OPTIONS
        dwFlags As Long
        dwAutoDetectFlags As Long
        lpszAutoConfigUrl As Long
        lpvReserved As Long
        dwReserved As Long
        fAutoLogonIfChallenged As Long
End Type

Public Type WINHTTP_PROXY_INFO
        dwAccessType As Long
        lpszProxy As Long
        lpszProxyBypass As Long
End Type

'-------------------------------------------------------------------------------
' System Functions
'-------------------------------------------------------------------------------
'
' Fills a specified DCB structure with values specified in
' a device-control string.
'
#If VBA7 Then
    Private Declare PtrSafe Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" _
    (ByVal lpDef As String, lpDCB As DCB) As Long
#Else
    Private Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" _
    (ByVal lpDef As String, lpDCB As DCB) As Long
#End If

'
' Retrieves information about a communications error and reports
' the current status of a communications device. The function is
' called when a communications error occurs, and it clears the
' device's error flag to enable additional input and output
' (I/O) operations.
'
#If VBA7 Then
    Private Declare PtrSafe Function ClearCommError Lib "kernel32" _
    (ByVal hFile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
#Else
    Private Declare Function ClearCommError Lib "kernel32" _
    (ByVal hFile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
#End If

'
' Closes an open communications device or file handle.
'
#If VBA7 Then
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
#Else
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
#End If

'
' Creates or opens a communications resource and returns a handle
' that can be used to access the resource.
'
#If VBA7 Then
    Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" _
    (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
#Else
    Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
    (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
#End If

'
' Directs a specified communications device to perform a function.
'
#If VBA7 Then
    Private Declare PtrSafe Function EscapeCommFunction Lib "kernel32" _
    (ByVal nCid As Long, ByVal nFunc As Long) As Long
#Else
    Private Declare Function EscapeCommFunction Lib "kernel32" _
    (ByVal nCid As Long, ByVal nFunc As Long) As Long
#End If

'
' Formats a message string such as an error string returned
' by anoher function.
'
#If VBA7 Then
    Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long
#Else
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long
#End If

'
' Retrieves modem control-register values.
'
#If VBA7 Then
    Private Declare PtrSafe Function GetCommModemStatus Lib "kernel32" _
    (ByVal hFile As Long, lpModemStat As Long) As Long
#Else
    Private Declare Function GetCommModemStatus Lib "kernel32" _
    (ByVal hFile As Long, lpModemStat As Long) As Long
#End If

'
' Retrieves the current control settings for a specified
' communications device.
'
#If VBA7 Then
    Private Declare PtrSafe Function GetCommState Lib "kernel32" _
    (ByVal nCid As Long, lpDCB As DCB) As Long
#Else
    Private Declare Function GetCommState Lib "kernel32" _
    (ByVal nCid As Long, lpDCB As DCB) As Long
#End If

'
' Retrieves the calling thread's last-error code value.
'
#If VBA7 Then
    Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
#Else
    Private Declare Function GetLastError Lib "kernel32" () As Long
#End If

'
' Retrieves the results of an overlapped operation on the
' specified file, named pipe, or communications device.
'
#If VBA7 Then
    Private Declare PtrSafe Function GetOverlappedResult Lib "kernel32" _
    (ByVal hFile As Long, lpOverlapped As OVERLAPPED, _
    lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
#Else
    Private Declare Function GetOverlappedResult Lib "kernel32" _
    (ByVal hFile As Long, lpOverlapped As OVERLAPPED, _
    lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
#End If

'
' Discards all characters from the output or input buffer of a
' specified communications resource. It can also terminate
' pending read or write operations on the resource.
'
#If VBA7 Then
    Private Declare PtrSafe Function PurgeComm Lib "kernel32" _
    (ByVal hFile As Long, ByVal dwFlags As Long) As Long
#Else
    Private Declare Function PurgeComm Lib "kernel32" _
    (ByVal hFile As Long, ByVal dwFlags As Long) As Long
#End If

'
' Reads data from a file, starting at the position indicated by the
' file pointer. After the read operation has been completed, the
' file pointer is adjusted by the number of bytes actually read,
' unless the file handle is created with the overlapped attribute.
' If the file handle is created for overlapped input and output
' (I/O), the application must adjust the position of the file pointer
' after the read operation.
'
#If VBA7 Then
    Private Declare PtrSafe Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, _
    lpOverlapped As OVERLAPPED) As Long
#Else
    Private Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, _
    lpOverlapped As OVERLAPPED) As Long
#End If

'
' Configures a communications device according to the specifications
' in a device-control block (a DCB structure). The function
' reinitializes all hardware and control settings, but it does not
' empty output or input queues.
'
#If VBA7 Then
    Private Declare PtrSafe Function SetCommState Lib "kernel32" _
    (ByVal hCommDev As Long, lpDCB As DCB) As Long
#Else
    Private Declare Function SetCommState Lib "kernel32" _
    (ByVal hCommDev As Long, lpDCB As DCB) As Long
#End If

'
' Sets the time-out parameters for all read and write operations on a
' specified communications device.
'
#If VBA7 Then
    Private Declare PtrSafe Function SetCommTimeouts Lib "kernel32" _
    (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
#Else
    Private Declare Function SetCommTimeouts Lib "kernel32" _
    (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
#End If

'
' Initializes the communications parameters for a specified
' communications device.
'
#If VBA7 Then
    Private Declare PtrSafe Function SetupComm Lib "kernel32" _
    (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
#Else
    Private Declare Function SetupComm Lib "kernel32" _
    (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
#End If

'
' Writes data to a file and is designed for both synchronous and
' asynchronous operation. The function starts writing data to the file
' at the position indicated by the file pointer. After the write
' operation has been completed, the file pointer is adjusted by the
' number of bytes actually written, except when the file is opened with
' FILE_FLAG_OVERLAPPED. If the file handle was created for overlapped
' input and output (I/O), the application must adjust the position of
' the file pointer after the write operation is finished.
'
#If VBA7 Then
    Private Declare PtrSafe Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, _
    lpOverlapped As OVERLAPPED) As Long
#Else
    Private Declare Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, _
    lpOverlapped As OVERLAPPED) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Sub AppSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub AppSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If



' Need CopyMemory to copy BSTR pointers around
#If VBA7 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (ByVal lpDest As Long, _
    ByVal lpSource As Long, ByVal cbCopy As Long)
#Else
    Public Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (ByVal lpDest As Long, _
    ByVal lpSource As Long, ByVal cbCopy As Long)
#End If


' SysAllocString creates a UNICODE BSTR string based on a UNICODE string
#If VBA7 Then
    Public Declare PtrSafe Function SysAllocString Lib "oleaut32" (ByVal pwsz As Long) As Long
#Else
    Public Declare Function SysAllocString Lib "oleaut32" (ByVal pwsz As Long) As Long
#End If


' Need GlobalFree to free the pointers in the CURRENT_USER_IE_PROXY_CONFIG
' structure returned from WinHttpGetIEProxyConfigForCurrentUser,
' per the documentation
#If VBA7 Then
    Public Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal p As Long) As Long
#Else
    Public Declare Function GlobalFree Lib "kernel32" (ByVal p As Long) As Long
#End If


' WinHttpGetIEProxyConfigForCurrentUser declaration
#If VBA7 Then
    Public Declare PtrSafe Function WinHttpGetIEProxyConfigForCurrentUser Lib "WinHTTP.dll" _
    (ByRef proxyConfig As WINHTTP_CURRENT_USER_IE_PROXY_CONFIG) As Long
#Else
    Public Declare Function WinHttpGetIEProxyConfigForCurrentUser Lib "WinHTTP.dll" _
    (ByRef proxyConfig As WINHTTP_CURRENT_USER_IE_PROXY_CONFIG) As Long
#End If

#If VBA7 Then
    Public Declare PtrSafe Function WinHttpGetProxyForUrl Lib "WinHTTP.dll" _
    (ByVal hSession As Long, _
    ByVal pszUrl As Long, _
    ByRef pAutoProxyOptions As WINHTTP_AUTOPROXY_OPTIONS, _
    ByRef pProxyInfo As WINHTTP_PROXY_INFO) As Long
#Else
    Public Declare Function WinHttpGetProxyForUrl Lib "WinHTTP.dll" _
    (ByVal hSession As Long, _
    ByVal pszUrl As Long, _
    ByRef pAutoProxyOptions As WINHTTP_AUTOPROXY_OPTIONS, _
    ByRef pProxyInfo As WINHTTP_PROXY_INFO) As Long
#End If

#If VBA7 Then
    Public Declare PtrSafe Function WinHttpOpen Lib "WinHTTP.dll" _
    (ByVal pszUserAgent As Long, _
    ByVal dwAccessType As Long, _
    ByVal pszProxyName As Long, _
    ByVal pszProxyBypass As Long, _
    ByVal dwFlags As Long) As Long
#Else
    Public Declare Function WinHttpOpen Lib "WinHTTP.dll" _
    (ByVal pszUserAgent As Long, _
    ByVal dwAccessType As Long, _
    ByVal pszProxyName As Long, _
    ByVal pszProxyBypass As Long, _
    ByVal dwFlags As Long) As Long
#End If

#If VBA7 Then
    Public Declare PtrSafe Function WinHttpCloseHandle Lib "WinHTTP.dll" _
    (ByVal hInternet As Long) As Long
#Else
    Public Declare Function WinHttpCloseHandle Lib "WinHTTP.dll" _
    (ByVal hInternet As Long) As Long
#End If


'-------------------------------------------------------------------------------
' Program Constants
'-------------------------------------------------------------------------------

Private Const MAX_PORTS = 16

'-------------------------------------------------------------------------------
' Program Structures
'-------------------------------------------------------------------------------

Private Type COMM_ERROR
    lngErrorCode As Long
    strFunction As String
    strErrorMessage As String
End Type

Private Type COMM_PORT
    lngHandle As Long
    blnPortOpen As Boolean
    udtDCB As DCB
End Type

 

'-------------------------------------------------------------------------------
' Program Storage
'-------------------------------------------------------------------------------

Private udtCommOverlap As OVERLAPPED
Private udtCommError As COMM_ERROR
Private udtPorts(1 To MAX_PORTS) As COMM_PORT
'-------------------------------------------------------------------------------
' GetSystemMessage - Gets system error text for the specified error code.
'-------------------------------------------------------------------------------
Public Function GetSystemMessage(lngErrorCode As Long) As String
Dim intPos As Integer
Dim strMessage As String, strMsgBuff As String * 256

    Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, lngErrorCode, 0, strMsgBuff, 255, 0)

    intPos = InStr(1, strMsgBuff, vbNullChar)
    If intPos > 0 Then
        strMessage = Trim$(Left$(strMsgBuff, intPos - 1))
    Else
        strMessage = Trim$(strMsgBuff)
    End If
    
    GetSystemMessage = strMessage
    
End Function
Public Function PauseApp(PauseInSeconds As Long)
    
    Call AppSleep(PauseInSeconds * 1000)
    
End Function

'-------------------------------------------------------------------------------
' CommOpen - Opens/Initializes serial port.
'
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   strPort     - COM port name. (COM1, COM2, COM3, COM4)
'   strSettings - Communication settings.
'                 Example: "baud=9600 parity=N data=8 stop=1"
'
' Returns:
'   Error Code  - 0 = No Error.
'
'-------------------------------------------------------------------------------
Public Function CommOpen(intPortID As Integer, strPort As String, _
    strSettings As String) As Long
    
Dim lngStatus       As Long
Dim udtCommTimeOuts As COMMTIMEOUTS

    On Error GoTo Routine_Error
    
    ' See if port already in use.
    If udtPorts(intPortID).blnPortOpen Then
        lngStatus = -1
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommOpen"
            .strErrorMessage = "Port in use."
        End With
        
        GoTo Routine_Exit
    End If

    ' Open serial port.
    udtPorts(intPortID).lngHandle = CreateFile(strPort, GENERIC_READ Or _
        GENERIC_WRITE, 0, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    If udtPorts(intPortID).lngHandle = -1 Then
        lngStatus = SetCommError("CommOpen (CreateFile)")
        GoTo Routine_Exit
    End If

    udtPorts(intPortID).blnPortOpen = True

    ' Setup device buffers (1K for input buffer, 1K for output).
    lngStatus = SetupComm(udtPorts(intPortID).lngHandle, 1024, 1024)
    
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetupComm)")
        GoTo Routine_Exit
    End If

    ' Purge buffers.
    lngStatus = PurgeComm(udtPorts(intPortID).lngHandle, PURGE_TXABORT Or _
        PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (PurgeComm)")
        GoTo Routine_Exit
    End If

    ' Set serial port timeouts.
    With udtCommTimeOuts
        .ReadIntervalTimeout = -1
        .ReadTotalTimeoutMultiplier = 0
        .ReadTotalTimeoutConstant = 1000
        .WriteTotalTimeoutMultiplier = 0
        .WriteTotalTimeoutMultiplier = 1000
    End With

    lngStatus = SetCommTimeouts(udtPorts(intPortID).lngHandle, udtCommTimeOuts)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetCommTimeouts)")
        GoTo Routine_Exit
    End If

    ' Get the current state (DCB).
    lngStatus = GetCommState(udtPorts(intPortID).lngHandle, _
        udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (GetCommState)")
        GoTo Routine_Exit
    End If

    ' Modify the DCB to reflect the desired settings.
    lngStatus = BuildCommDCB(strSettings, udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (BuildCommDCB)")
        GoTo Routine_Exit
    End If

    ' Set the new state.
    lngStatus = SetCommState(udtPorts(intPortID).lngHandle, _
        udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
    CommOpen = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommOpen"
        .strErrorMessage = Err.description
    End With
    Resume Routine_Exit
End Function

Private Function SetCommError(strFunction As String) As Long
    
    With udtCommError
        .lngErrorCode = Err.LastDllError
        .strFunction = strFunction
        .strErrorMessage = GetSystemMessage(.lngErrorCode)
        SetCommError = .lngErrorCode
    End With
End Function

Private Function SetCommErrorEx(strFunction As String, lngHnd As Long) As Long
Dim lngErrorFlags As Long
Dim udtCommStat As COMSTAT
    
    With udtCommError
        .lngErrorCode = GetLastError
        .strFunction = strFunction
        .strErrorMessage = GetSystemMessage(.lngErrorCode)
    
        Call ClearCommError(lngHnd, lngErrorFlags, udtCommStat)
    
        .strErrorMessage = .strErrorMessage & "  COMM Error Flags = " & _
                Hex$(lngErrorFlags)
        
        SetCommErrorEx = .lngErrorCode
    End With
    
End Function

'-------------------------------------------------------------------------------
' CommSet - Modifies the serial port settings.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   strSettings - Communication settings.
'                 Example: "baud=9600 parity=N data=8 stop=1"
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommSet(intPortID As Integer, strSettings As String) As Long
    
Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = GetCommState(udtPorts(intPortID).lngHandle, _
        udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (GetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = BuildCommDCB(strSettings, udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (BuildCommDCB)")
        GoTo Routine_Exit
    End If

    lngStatus = SetCommState(udtPorts(intPortID).lngHandle, _
        udtPorts(intPortID).udtDCB)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (SetCommState)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
    CommSet = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommSet"
        .strErrorMessage = Err.description
    End With
    Resume Routine_Exit
End Function
'-------------------------------------------------------------------------------
' CommClose - Close the serial port.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommClose(intPortID As Integer) As Long
    
Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    If udtPorts(intPortID).blnPortOpen Then
        lngStatus = CloseHandle(udtPorts(intPortID).lngHandle)
    
        If lngStatus = 0 Then
            lngStatus = SetCommError("CommClose (CloseHandle)")
            GoTo Routine_Exit
        End If
    
        udtPorts(intPortID).blnPortOpen = False
    End If

    lngStatus = 0

Routine_Exit:
    CommClose = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommClose"
        .strErrorMessage = Err.description
    End With
    Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommFlush - Flush the send and receive serial port buffers.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommFlush(intPortID As Integer) As Long
    
Dim lngStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = PurgeComm(udtPorts(intPortID).lngHandle, PURGE_TXABORT Or _
        PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommFlush (PurgeComm)")
        GoTo Routine_Exit
    End If

    lngStatus = 0

Routine_Exit:
    CommFlush = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommFlush"
        .strErrorMessage = Err.description
    End With
    Resume Routine_Exit
End Function


'-------------------------------------------------------------------------------
' CommRead - Read serial port input buffer.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   strData     - Data buffer.
'   lngSize     - Maximum number of bytes to be read.
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommRead(intPortID As Integer, strData As String, _
    lngSize As Long) As Long

Dim lngStatus As Long
Dim lngRdSize As Long, lngBytesRead As Long
Dim lngRdStatus As Long, strRdBuffer As String * 1024
Dim lngErrorFlags As Long, udtCommStat As COMSTAT
    
    On Error GoTo Routine_Error

    strData = ""
    lngBytesRead = 0
    DoEvents
    
    ' Clear any previous errors and get current status.
    lngStatus = ClearCommError(udtPorts(intPortID).lngHandle, lngErrorFlags, _
        udtCommStat)

    If lngStatus = 0 Then
        lngBytesRead = -1
        lngStatus = SetCommError("CommRead (ClearCommError)")
        GoTo Routine_Exit
    End If
        
    If udtCommStat.cbInQue > 0 Then
        If udtCommStat.cbInQue > lngSize Then
            lngRdSize = udtCommStat.cbInQue
        Else
            lngRdSize = lngSize
        End If
    Else
        lngRdSize = 0
    End If

    If lngRdSize Then
        lngRdStatus = ReadFile(udtPorts(intPortID).lngHandle, strRdBuffer, _
            lngRdSize, lngBytesRead, udtCommOverlap)

        If lngRdStatus = 0 Then
            lngStatus = GetLastError
            If lngStatus = ERROR_IO_PENDING Then
                ' Wait for read to complete.
                ' This function will timeout according to the
                ' COMMTIMEOUTS.ReadTotalTimeoutConstant variable.
                ' Every time it times out, check for port errors.

                ' Loop until operation is complete.
                While GetOverlappedResult(udtPorts(intPortID).lngHandle, _
                    udtCommOverlap, lngBytesRead, True) = 0
                                    
                    lngStatus = GetLastError
                                        
                    If lngStatus = ERROR_IO_INCOMPLETE Then
                        lngBytesRead = -1
                        lngStatus = SetCommErrorEx( _
                            "CommRead (GetOverlappedResult)", _
                            udtPorts(intPortID).lngHandle)
                        GoTo Routine_Exit
                    End If
                Wend
            Else
                ' Some other error occurred.
                lngBytesRead = -1
                lngStatus = SetCommErrorEx("CommRead (ReadFile)", _
                    udtPorts(intPortID).lngHandle)
                GoTo Routine_Exit
            
            End If
        End If
    
        strData = Left$(strRdBuffer, lngBytesRead)
    End If

Routine_Exit:
    CommRead = lngBytesRead
    Exit Function

Routine_Error:
    lngBytesRead = -1
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommRead"
        .strErrorMessage = Err.description
    End With
    Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommWrite - Output data to the serial port.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   strData     - Data to be transmitted.
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommWrite(intPortID As Integer, strData As String) As Long
    
Dim i As Integer
Dim lngStatus As Long, lngSize As Long
Dim lngWrSize As Long, lngWrStatus As Long
    
    On Error GoTo Routine_Error
    
    ' Get the length of the data.
    lngSize = Len(strData)

    ' Output the data.
    lngWrStatus = WriteFile(udtPorts(intPortID).lngHandle, strData, lngSize, _
        lngWrSize, udtCommOverlap)

    ' Note that normally the following code will not execute because the driver
    ' caches write operations. Small I/O requests (up to several thousand bytes)
    ' will normally be accepted immediately and WriteFile will return true even
    ' though an overlapped operation was specified.
        
    DoEvents
    
    If lngWrStatus = 0 Then
        lngStatus = GetLastError
        If lngStatus = 0 Then
            GoTo Routine_Exit
        ElseIf lngStatus = ERROR_IO_PENDING Then
            ' We should wait for the completion of the write operation so we know
            ' if it worked or not.
            '
            ' This is only one way to do this. It might be beneficial to place the
            ' writing operation in a separate thread so that blocking on completion
            ' will not negatively affect the responsiveness of the UI.
            '
            ' If the write takes long enough to complete, this function will timeout
            ' according to the CommTimeOuts.WriteTotalTimeoutConstant variable.
            ' At that time we can check for errors and then wait some more.

            ' Loop until operation is complete.
            While GetOverlappedResult(udtPorts(intPortID).lngHandle, _
                udtCommOverlap, lngWrSize, True) = 0
                                
                lngStatus = GetLastError
                                    
                If lngStatus = ERROR_IO_INCOMPLETE Then
                    lngStatus = SetCommErrorEx( _
                        "CommWrite (GetOverlappedResult)", _
                        udtPorts(intPortID).lngHandle)
                    GoTo Routine_Exit
                End If
            Wend
        Else
            ' Some other error occurred.
            lngWrSize = -1
                    
            lngStatus = SetCommErrorEx("CommWrite (WriteFile)", _
                udtPorts(intPortID).lngHandle)
            GoTo Routine_Exit
        
        End If
    End If
    
    For i = 1 To 10
        DoEvents
    Next
    
Routine_Exit:
    CommWrite = lngWrSize
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommWrite"
        .strErrorMessage = Err.description
    End With
    Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommGetLine - Get the state of selected serial port control lines.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   intLine     - Serial port line. CTS, DSR, RING, RLSD (CD)
'   blnState    - Returns state of line (Cleared or Set).
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommGetLine(intPortID As Integer, intLine As Integer, _
   blnState As Boolean) As Long
    
Dim lngStatus As Long
Dim lngComStatus As Long, lngModemStatus As Long
    
    On Error GoTo Routine_Error

    lngStatus = GetCommModemStatus(udtPorts(intPortID).lngHandle, lngModemStatus)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommReadCD (GetCommModemStatus)")
        GoTo Routine_Exit
    End If

    If (lngModemStatus And intLine) Then
        blnState = True
    Else
        blnState = False
    End If
        
    lngStatus = 0
        
Routine_Exit:
    CommGetLine = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommReadCD"
        .strErrorMessage = Err.description
    End With
    Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommSetLine - Set the state of selected serial port control lines.
'
' Parameters:
'   intPortID   - Port ID used when port was opened.
'   intLine     - Serial port line. BREAK, DTR, RTS
'                 Note: BREAK actually sets or clears a "break" condition on
'                 the transmit data line.
'   blnState    - Sets the state of line (Cleared or Set).
'
' Returns:
'   Error Code  - 0 = No Error.
'-------------------------------------------------------------------------------
Public Function CommSetLine(intPortID As Integer, intLine As Integer, _
   blnState As Boolean) As Long
   
Dim lngStatus As Long
Dim lngNewState As Long
    
    On Error GoTo Routine_Error
    
    If intLine = LINE_BREAK Then
        If blnState Then
            lngNewState = SETBREAK
        Else
            lngNewState = CLRBREAK
        End If
    
    ElseIf intLine = LINE_DTR Then
        If blnState Then
            lngNewState = SETDTR
        Else
            lngNewState = CLRDTR
        End If
    
    ElseIf intLine = LINE_RTS Then
        If blnState Then
            lngNewState = SETRTS
        Else
            lngNewState = CLRRTS
        End If
    End If

    lngStatus = EscapeCommFunction(udtPorts(intPortID).lngHandle, lngNewState)

    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSetLine (EscapeCommFunction)")
        GoTo Routine_Exit
    End If

    lngStatus = 0
        
Routine_Exit:
    CommSetLine = lngStatus
    Exit Function

Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommSetLine"
        .strErrorMessage = Err.description
    End With
    Resume Routine_Exit
End Function

'-------------------------------------------------------------------------------
' CommGetError - Get the last serial port error message.
'
' Parameters:
'   strMessage  - Error message from last serial port error.
'
' Returns:
'   Error Code  - Last serial port error code.
'-------------------------------------------------------------------------------
Public Function CommGetError(strMessage As String) As Long
    
    With udtCommError
        CommGetError = .lngErrorCode
        strMessage = "Error (" & CStr(.lngErrorCode) & "): " & .strFunction & _
            " - " & .strErrorMessage
    End With
    
End Function

Public Function MB_LoopBack(mb_add As Integer, subFunction As Integer, writeBuffer() As Long, writeLen As Integer, Readbuffer() As Long, iError As Integer, intPortID As Integer) As Integer
    'loopback using Modbus function code 08
    'Send:
    '[Slave Add][Func 08][HB subFunction][LB subFunction][HB D1][Lb D1][HB D2][LB D2]...[HB Dn][LB Dn][LB CRC][HB CRC]
    'Response:
    '[Slave Add][Func 08 dec AND error code][HB subFunction][HB D1][Lb D1][HB D2][LB D2]...[HB Dn][LB Dn][LB CRC][HB CRC]
    
    Dim ModbusMsg(256) As Byte
    Dim ModbusReply(256) As Byte
    Dim CRC As Long
    Dim lTemp As Long
    Dim MsgStr As String
    Dim i As Integer, X As Integer
    Dim temp_string As String
    Dim byte_counter As Integer
    Dim lngStatus As Long
    Dim MBReply As String, MBtimeout As Integer, sReplyLen As Integer
    Dim MBData(128) As Integer, ret As Integer
    Dim sReq As String, sResp As String, recvDone As Integer, MBJabber As Integer
    Dim iException As Integer, iFlag As Integer, writeMsgLen As Integer
    
    'g_bCommActive = True    'comm active with this device (might be redundant, but include just to be sure)
    
    ModbusMsg(0) = mb_add And 255
    ModbusMsg(1) = 8
    ModbusMsg(2) = (subFunction \ 256) And 255
    ModbusMsg(3) = (subFunction - ModbusMsg(2) * 256) And 255
    For i = 4 To writeLen + 4 - 1
        ModbusMsg(i) = writeBuffer(i - 4) And 255
    Next
    
    'compute the CRC
    CRC = ModCRC(ModbusMsg(), 4 + writeLen)
    'must invert byte order HB->LB and LB->HB
    ModbusMsg(5 + writeLen) = (CRC \ 256) And &HFF
    lTemp = ModbusMsg(7)
    ModbusMsg(4 + writeLen) = CRC - lTemp * 256 And &HFF
    
'    ModbusMsg = ModbusMsg + CRC \ 256
'    ModbusMsg = ModbusMsg + CRC - ModbusMsg(7) * 256
    
    'convert byte array to hex string
    For i = 0 To 4 + writeLen + 1
        'temp_string = (Convert.ToString(ModbusMsg(i), 16).PadLeft(2, "0"c).PadRight(3, " "c))
        temp_string = Chr(ModbusMsg(i))
        'padded = StrDup(2 - Len(temp_string), "0") & temp_string
        'MsgStr = MsgStr + padded
        MsgStr = MsgStr + temp_string
    Next
    writeMsgLen = 6 + writeLen
    lngStatus = CommWrite(intPortID, MsgStr)
    'WriteByte(ModbusMsg, 8, True)
    
    Pause (200) 'pause 100 ms (defined above)
    recvDone = 0    'clear this flag
    'read what has been returned
    Do Until recvDone = 1
    lngStatus = CommRead(intPortID, MsgStr, 10)
    MBReply = MsgStr
        If lngStatus Then
            'something heard - increment max duration timer
            MBJabber = MBJabber + 1
            MBtimeout = 0   'reset timeout timer
            If MBJabber > MBJabberTimeout Then Exit Do
            Do
                Pause (gMBTimeout * 2) 'wait 40 ms
                lngStatus = CommRead(intPortID, MsgStr, 10)
                MBReply = MBReply + MsgStr
                If lngStatus = 0 Then
                    recvDone = 1
                    Exit Do
                End If
                'increment counter to prevent infinite loop
                MBtimeout = MBtimeout + 1
                If MBtimeout > 1 Then
                    recvDone = 1
                    MB_LoopBack = 0  'timed out waiting for response
                    Exit Do
                End If
            Loop
        Else
            'nothing heard, increment timeout counter
            MBtimeout = MBtimeout + 1
            Pause (gMBTimeout)
            If MBtimeout > MBTimeoutMAX Then
                recvDone = 1
                MB_LoopBack = 0  'timed out waiting for response
                Exit Do
            End If
        End If
    Loop
    
    'now parse the reply stored in global array data_buffer
    'first test is verify that the number of bytes returned is the number expected
    'ModbusMsg(5) is the number of 16 bit registers expected.
    'Since the number of bytes is two times the number of registers, then the
    'byte count of the reply includes (with number of of bytes):
    'Slave Address (1), Function Code (1), byte count (1), data (2 * ModbusMsg(5), CRC (2)
    'or a total of:
    '5 + (2 * ModbusMsg(5))
    
    'now convert the string into bytes and place in buffer()
    sReplyLen = Len(MBReply)
    For i = 0 To sReplyLen - 1
        ModbusReply(i) = Asc(Mid$(MBReply, i + 1, 1))
    Next
        
    If Len(MBReply) = (6 + writeLen) Then
        If VerifyCRC(ModbusReply, sReplyLen) Then    '**** create data_buffer
            'CRC on reply is good
            MB_LoopBack = 1
            
            ret = DecodeMBReply(8, ModbusReply(), Readbuffer(), i + 1)
        Else
            'CRC on reply failed
            MB_LoopBack = 0
            g_MB_errors = g_MB_errors + 1
        End If
        'need to check subfunction code to see if we got the right answer
        If subFunction <> 0 Then
            'check to see if first bit of the second word is set
            iException = ModbusReply(1) And 32767
            If ModbusReply(1) And 32768 Then
                MsgBox ("Modbus exception code " & iException)
            Else
                'expect different response
                iFlag = 0   'clear to start
                If ModbusMsg(6) = ModbusReply(6) And ModbusMsg(7) = ModbusReply(7) Then
                    iFlag = 1   'CRC match -- reply identical to request
                    MsgBox ("Echo of request received.  Possible problem with echo cancellation on 2-wire modem.")
                    MB_LoopBack = 0
                End If
            End If
            
        End If
    Else
        'byte count wrong on reply - exception code
        MB_LoopBack = 0
        g_MB_errors = g_MB_errors + 1
        ret = DecodeMBErrorReply(3, ModbusReply(), writeBuffer())
        'buffer() now contains the error bits
        If writeBuffer(2) <> 0 Then
            iError = writeBuffer(2) 'this is the Modbus Exception code returned
        Else
            iError = 255            '255: nothing returned
        End If
        
        'collect the first 10 bytes of the request and of the response
        sReq = ""
        sResp = ""
        If writeMsgLen > 10 Then
            For i = 0 To 9
                sReq = sReq & " " & ModbusMsg(i)
            Next
            sReq = sReq & "..."
        Else
            For i = 0 To writeMsgLen - 1
                sReq = sReq & " " & ModbusMsg(i)
            Next

        End If
        
        If sReplyLen > 10 Then
            For i = 0 To 9
                sResp = sResp & " " & ModbusReply(i)
            Next
            sResp = sResp & "..."
        Else
            For i = 0 To sReplyLen
                If sReplyLen <> 0 Then
                    sResp = sResp & " " & ModbusReply(i)
                End If
            Next
        End If

        If sReplyLen <> 0 Then
            MsgBox ("Modbus exception code (" & iError & ": " & MBExceptionDecode(iError) & ") received. [Request: " & sReq & " Response: " & sResp & "]")
        Else
            MsgBox ("No response received (Error code: " & iError & ") [Request: " & sReq & " Response: " & sResp & "]")
        End If
        
        
    End If
    
End Function
Function MBExceptionDecode(exceptionCode As Integer) As String

    Select Case exceptionCode
    
        Case 1
            MBExceptionDecode = "That function is not supported by this device"
        Case 2
            MBExceptionDecode = "Register address requested is not available"
        Case 3
            MBExceptionDecode = "Improperly formed message. Internal error"
        Case 4
            MBExceptionDecode = "Downstream device timeout. Check wiring, power, etc."
        Case 5
            MBExceptionDecode = "Internal error. Try again"
        Case 7
            MBExceptionDecode = "NAK error reported. Try again"
        Case 84
            MBExceptionDecode = "Only a partial register was requested. Internal error"
        Case Else
            MBExceptionDecode = "Unknown exception code received.  Verify modem settings (especially regarding echo cancellation and 2-wire operation.)"
    
    End Select
    
End Function

Public Function MB_Read(mb_add As Integer, start As Integer, Length As Integer, buffer() As Long, iError As Integer, intPortID As Integer) As Integer
    'read from slave using Modbus function code 03
    'Send:
    '[Slave Add][Func 03][HB start add][LB start add][HB # regs to write][LB of # regs to write][LB CRC][HB CRC]
    'Response:
    '[Slave Add][Func 03 dec AND error code][Byte Count][HB D1][Lb D1][HB D2][LB D2]...[HB Dn][LB Dn][LB CRC][HB CRC]
    
'    Dim ModbusMsg As String
    Dim ModbusMsg(256) As Byte
    Dim ModbusReply(256) As Byte
    Dim CRC As Long
    Dim lTemp As Long
    Dim MsgStr As String
    Dim i As Integer, X As Integer
    Dim temp_string As String
    Dim byte_counter As Integer
    Dim lngStatus As Long
    Dim MBReply As String, MBtimeout As Integer, sReplyLen As Integer
    Dim MBData(128) As Integer, ret As Integer
    Dim sReq As String, sResp As String, recvDone As Integer, MBJabber As Integer
    Dim udtCommStat As COMSTAT
    
    'gMBTimeout = 480000 / gretBPS
    
    'g_bCommActive = True    'comm active with this device (might be redundant, but include just to be sure)
    
    ModbusMsg(0) = mb_add And 255
    ModbusMsg(1) = 3
    ModbusMsg(2) = (start \ 256) And &HFF
    ModbusMsg(3) = (start - ModbusMsg(2) * 256) And &HFF
    ModbusMsg(4) = (Length \ 256) And &HFF
    ModbusMsg(5) = (Length - ModbusMsg(4) * 256) And &HFF
    
    'compute the CRC
    CRC = ModCRC(ModbusMsg(), 6)
    'must invert byte order HB->LB and LB->HB
    ModbusMsg(7) = (CRC \ 256) And &HFF
    lTemp = ModbusMsg(7)
    ModbusMsg(6) = CRC - lTemp * 256 And &HFF
    
'    ModbusMsg = ModbusMsg + CRC \ 256
'    ModbusMsg = ModbusMsg + CRC - ModbusMsg(7) * 256
    
    'convert byte array to hex string
    For i = 0 To (8) - 1
        'temp_string = (Convert.ToString(ModbusMsg(i), 16).PadLeft(2, "0"c).PadRight(3, " "c))
        temp_string = Chr(ModbusMsg(i))
        'padded = StrDup(2 - Len(temp_string), "0") & temp_string
        'MsgStr = MsgStr + padded
        MsgStr = MsgStr + temp_string
    Next
    lngStatus = CommWrite(intPortID, MsgStr)
    'WriteByte(ModbusMsg, 8, True)
    
    'Worksheets("Test").Range("a1").Value = ""
    'Worksheets("Test").Range("a2").Value = ""
    
    Pause (gMBTimeout)
    
    recvDone = 0    'clear this flag
    'read what has been returned
    Do Until recvDone = 1
        lngStatus = CommRead(intPortID, MsgStr, 10)
        MBReply = MsgStr
        If lngStatus Then
            'something heard - increment max duration timer
            MBJabber = MBJabber + 1
            MBtimeout = 0   'reset timeout timer
            If MBJabber > MBJabberTimeout Then Exit Do
            Do
                'Worksheets("Test").Range("a1").Value = "Collect time: " & MBJabber * 0.02
                Pause (gMBTimeout)
                
                lngStatus = CommRead(intPortID, MsgStr, 10)
                
                MBReply = MBReply + MsgStr
                If lngStatus = 0 Then
                    recvDone = 1
                    Exit Do
                End If
                
                'increment counter to prevent infinite loop
                MBtimeout = MBtimeout + 1
                'Worksheets("Test").Range("c1").Value = MBtimeout * 0.02
                If MBtimeout > MBTimeoutMAX Then
                    recvDone = 1
                    MB_Read = 0  'timed out waiting for response
                    Exit Do
                End If
            Loop
        Else
            'nothing heard, increment timeout counter
            'Worksheets("Test").Range("a2").Value = "Dead wait time: " & MBtimeout * 0.02
            MBtimeout = MBtimeout + 1
            Pause (gMBTimeout)
            If MBtimeout > MBTimeoutMAX Then
                recvDone = 1
                MB_Read = 0  'timed out waiting for response
                Exit Do
            End If
        End If
    Loop
    
    'now parse the reply stored in global array data_buffer
    'first test is verify that the number of bytes returned is the number expected
    'ModbusMsg(5) is the number of 16 bit registers expected.
    'Since the number of bytes is two times the number of registers, then the
    'byte count of the reply includes (with number of of bytes):
    'Slave Address (1), Function Code (1), byte count (1), data (2 * ModbusMsg(5), CRC (2)
    'or a total of:
    '5 + (2 * ModbusMsg(5))
    
    'now convert the string into bytes and place in buffer()
    sReplyLen = Len(MBReply)
    For i = 0 To sReplyLen - 1
        ModbusReply(i) = Asc(Mid$(MBReply, i + 1, 1))
    Next
        
    If Len(MBReply) = (5 + (2 * ModbusMsg(5))) Then
        
        If VerifyCRC(ModbusReply, sReplyLen) Then    '**** create data_buffer
            'CRC on reply is good
            MB_Read = 1
            
            ret = DecodeMBReply(3, ModbusReply(), buffer(), i + 1)
        Else
            'CRC on reply failed
            MB_Read = 0
            g_MB_errors = g_MB_errors + 1
        End If
    Else
        'byte count wrong on reply - exception code
        'first check to see if exact echo of message received
        MB_Read = 0
        g_MB_errors = g_MB_errors + 1
        ret = DecodeMBErrorReply(3, ModbusReply(), buffer())
        
        'buffer() now contains the error bits
        If buffer(2) <> 0 Then
            iError = buffer(2) 'this is the Modbus Exception code returned
        Else
            iError = 255            '255: nothing returned
        End If
        
        If Len(MBReply) = 8 Then
            'check to see if bytes sent match bytes received
            If ModbusMsg(4) = Asc(Mid$(MBReply, 5, 1)) Then
                If ModbusMsg(5) = Asc(Mid$(MBReply, 6, 1)) Then
                    If ModbusMsg(6) = Asc(Mid$(MBReply, 7, 1)) Then
                        MsgBox ("Exact echo of request received.  Make sure your modem/convert is set properly to echo cancellation in 2-wire mode.")
                    Else
                        'bytes 4 and 5 matched, but not 6
                        MsgBox ("Echo of request received.  Make sure your modem or 485 converter is set properly to echo cancellation in 2-wire mode.")
                    End If
                Else
                    'byte 4 matched, but 5 didn't
                        MsgBox ("Unexpected message reply")
                End If
            End If
        End If
        
        'collect the first 4 bytes of the request and of the response
        sReq = ""
        sResp = ""
        For i = 0 To 3
            sReq = sReq & " " & ModbusMsg(i)
            If sReplyLen <> 0 Then
                sResp = sResp & " " & ModbusReply(i)
            End If
        Next
        sReq = sReq & "..."
        sResp = sResp & "..."
        If sReplyLen <> 0 Then
            MsgBox ("Modbus exception code (" & iError & ": " & MBExceptionDecode(iError) & ") received. [Request: " & sReq & " Response: " & sResp & "]")
        Else
            MsgBox ("No response received (Error code: " & iError & ") [Request: " & sReq & " Response: " & sResp & "]")
        End If
        
    End If
    
End Function
Public Function MB_Write(mb_add As Byte, start As Integer, data_2_write() As Long, Length As Integer, iError As Integer, intPortID As Integer) As Byte
    'write to slave using Modbus function code 16 (10 hex)
    'Send:
    '[Slave Add][Func 16 dec, 10 hex][HB start add][LB start add][HB # regs to write][LB of # regs to write][Byte Count][HB D1][Lb D1][HB D2][LB D2]...[HB Dn][LB Dn][LB CRC][HB CRC]
    'Response:
    '[Slave Add][Func 16 dec AND error code][HB start add][LB start add][HB # regs to write][LB of # regs to write][LB CRC][HB CRC]

    Dim ModbusMsg(256) As Byte
    Dim ModbusReply(256) As Byte
    Dim CRC As Long
    Dim lTemp As Long
    Dim MsgStr As String
    Dim i As Integer, X As Integer
    Dim temp_string As String
    Dim byte_counter As Integer
    Dim lngStatus As Long
    Dim MBReply As String, MBtimeout As Integer, sReplyLen As Integer
    Dim MBData(128) As Integer, ret As Integer
    Dim buffer(256) As Long, temp As Long
    Dim sReq As String, sResp As String, recvDone As Integer, MBJabber As Integer
    
    'byte_counter = 0 'clear

    ModbusMsg(0) = mb_add
    ModbusMsg(1) = 16   '10 hex
    ModbusMsg(2) = start \ 256
    ModbusMsg(3) = start - ModbusMsg(2) * 256
    ModbusMsg(4) = Length \ 256
    ModbusMsg(5) = Length - ModbusMsg(4) * 256
    ModbusMsg(6) = Length * 2
    For i = 0 To Length - 1
        ModbusMsg(7 + 2 * i) = (data_2_write(i) \ 256) And 255
        ModbusMsg(8 + 2 * i) = data_2_write(i) - CLng(ModbusMsg(7 + 2 * i)) * 256
    Next

    'compute the CRC
    CRC = ModCRC(ModbusMsg, 7 + 2 * Length)
    'must invert byte order HB->LB and LB->HB
    ModbusMsg(8 + 2 * Length) = CRC \ 256
    temp = ModbusMsg(8 + 2 * Length)
    ModbusMsg(7 + 2 * Length) = (CRC - temp * 256) And 255

    'convert byte array to hex string
    For i = 0 To (9 + 2 * Length) - 1
        temp_string = Chr(ModbusMsg(i))
        'temp_string = (Cstr(ModbusMsg(i), 16).PadLeft(2, "0"c).PadRight(3, " "c))
        'padded = StrDup(2 - Len(temp_string), "0") & temp_string
        'MsgStr = MsgStr + padded
        MsgStr = MsgStr + temp_string
    Next
    lngStatus = CommWrite(intPortID, MsgStr)

    Pause (gMBTimeout) 'pause 200 ms (defined above)
    recvDone = 0    'clear this flag
    'read what has been returned
    Do Until recvDone = 1
    lngStatus = CommRead(intPortID, MsgStr, 10)
    MBReply = MsgStr
        If lngStatus Then
            'something heard - increment max duration timer
            MBJabber = MBJabber + 1
            MBtimeout = 0   'reset timeout timer
            If MBJabber > MBJabberTimeout Then Exit Do
            Do
                Pause (gMBTimeout)  'wait 20 ms
                lngStatus = CommRead(intPortID, MsgStr, 10)
                MBReply = MBReply + MsgStr
                If lngStatus = 0 Then
                    recvDone = 1
                    Exit Do
                End If
            Loop
        Else
            'nothing heard, increment timeout counter
            MBtimeout = MBtimeout + 1
            Pause (gMBTimeout)
            If MBtimeout > MBTimeoutMAX Then
                recvDone = 1
                MB_Write = 0  'timed out waiting for response
                Exit Do
            End If
        End If
    Loop
    
    'now parse the reply stored in global array data_buffer
    'first test is verify that the number of bytes returned is the number expected
    'ModbusMsg(5) is the number of 16 bit registers expected.
    'Since the number of bytes is two times the number of registers, then the
    'byte count of the reply includes (with number of of bytes):
    'Slave Address (1), Function Code (1), byte count (1), data (2 * ModbusMsg(5), CRC (2)
    'or a total of:
    '5 + (2 * ModbusMsg(5))
    
    'now convert the string into bytes and place in buffer()
    sReplyLen = Len(MBReply)
    For i = 0 To sReplyLen - 1
        ModbusReply(i) = Asc(Mid$(MBReply, i + 1, 1))
    Next
        
    If Len(MBReply) = 8 Then
        
        If VerifyCRC(ModbusReply, sReplyLen) Then    '**** create data_buffer
            'CRC on reply is good
            MB_Write = 1
            iError = 0
            ret = DecodeMBReply(16, ModbusReply(), buffer(), i + 1)
        Else
            'CRC on reply failed
            MB_Write = 0
            g_MB_errors = g_MB_errors + 1
        End If
    Else
        'byte count wrong on reply - exception code
        MB_Write = 0
        g_MB_errors = g_MB_errors + 1
        ret = DecodeMBErrorReply(3, ModbusReply(), buffer())
        'buffer() now contains the error bits
        If buffer(2) <> 0 Then
            iError = buffer(2) 'this is the Modbus Exception code returned
        Else
            iError = 255            '255: nothing returned
        End If
        
        'collect the first 6 bytes of the request and of the response
        sReq = ""
        sResp = ""
        For i = 0 To 5
            sReq = sReq & " " & ModbusMsg(i)
            If sReplyLen <> 0 Then
                sResp = sResp & " " & ModbusReply(i)
            End If
        Next
        sReq = sReq & "..."
        sResp = sResp & "..."
        If sReplyLen <> 0 Then
            MsgBox ("Modbus exception code (" & iError & ": " & MBExceptionDecode(iError) & ") received. [Request: " & sReq & " Response: " & sResp & "] " & sReplyLen + 1 & " bytes returned")
        Else
            MsgBox ("No response received (Error code: " & iError & ") [Request: " & sReq & " Response: " & sResp & "]")
        End If
    End If

End Function
Function DecodeMBErrorReply(opcode As Integer, Inputbuffer() As Byte, Outputbuffer() As Long) As Integer

    Dim i As Integer
    
    'Length not needed since all MB exception messages are the same length
    For i = 0 To 4
        Outputbuffer(i) = Inputbuffer(i)
    Next
    
End Function
Function DecodeMBReply(opcode As Integer, Inputbuffer() As Byte, Outputbuffer() As Long, Length As Integer) As Integer
    'opcode = 3 Read Registers
    '   byte 1  MB address
    '   byte 2  Function code (03) or 83 hex (131 dec) if error
    '   byte 3  Byte count or exception code if MSB of byte 2 = 1
    '   byte 4  Hi byte of first register
    '   byte 5  Lo byte of first register
    '   byte 6  Hi byte of 2nd register
    '   byte 7  Lo byte of 2nd register
    '   byte 2*(N+1)        Hi byte of Nth register
    '   byte 2*(N+1) + 1    Lo byte of Nth register
    '   byte 2*(N+2) - 1    Lo byte of CRC
    '   byte 2*(N+2) - 1    Hi byte of CRC
    
    'opcode = 16 dec Write Registers
    '   byte 1  MB address
    '   byte 2  Function code (10 hex or 16 dec) or 90 hex or 144 dec if error
    '   byte 3  Hi byte of starting address or exception code if MSB of byte 2 = 1
    '   byte 4  Lo byte of starting address
    '   byte 5  Hi byte of quantity of registers written
    '   byte 6  Lo byte of quantity of registers written
    '   byte 7  Lo byte of CRC
    '   byte 8  Hi byte of CRC
    
    Dim i As Integer, temp As Long, j As Integer
    
    'check opcode
    Select Case opcode
        Case 3
            '2(N+1) * 65536 + 2(N+1)-1
            j = 0
            For i = 1 To Inputbuffer(2) / 2
                temp = (256 * CLng(Inputbuffer(2 * i + 1))) + Inputbuffer(2 * i + 2)
                Outputbuffer(j) = temp
                j = j + 1
            Next
        
        Case 8
            j = 0
            For i = 1 To (Length - 5) / 2
                temp = 256 * CLng(Inputbuffer(2 * i)) + Inputbuffer(2 * i + 1)
                Outputbuffer(j) = temp And 65535
                j = j + 1
            Next
            
        Case 16
            j = 0
            For i = 1 To (Length - 5) / 2
                temp = 256 * CLng(Inputbuffer(2 * i)) + Inputbuffer(2 * i + 1)
                Outputbuffer(j) = temp And 65535
                j = j + 1
            Next
            
        Case Else
        
    End Select

End Function

Public Function MBGenericMsg(txBuffer() As Byte, rxBuffer As String, iLength As Integer, intPortID As Integer) As Integer
'accepts bytes, adds CRC, transmits message and receives response

    Dim ModbusMsg(256) As Byte
    Dim CRC As Long
    Dim CRC2_Hi As Byte
    Dim CRC1_Lo As Byte
    Dim lTemp As Long
    Dim MsgStr As String
    Dim i As Integer
    Dim lngStatus As Long, MBtimeout As Integer
    
    For i = 0 To iLength - 1
        MsgStr = MsgStr + Chr(txBuffer(i))
    Next
    
    'compute the CRC
    'CRC = ModCRC(ModbusMsg(), iLength)
    CRC = CRC16A(MsgStr, iLength)
    'CRC = SolveMBCRC(MsgStr, iLength)
    
    CRC2_Hi = (CRC \ 256) And &HFF
    lTemp = CRC2_Hi     'gets around a "bug" in VBA that prevents multiplying byte values?
    CRC1_Lo = CRC - lTemp * 256 And &HFF
    
    MsgStr = MsgStr + Chr(CRC1_Lo) + Chr(CRC2_Hi)
    
    lngStatus = CommWrite(intPortID, MsgStr)    'send byte string
    
    Pause (gMBTimeout)   'pause 200 ms (defined above)
    'read what has been returned
    lngStatus = CommRead(intPortID, MsgStr, 10)
    rxBuffer = MsgStr
    MBtimeout = 0
    Do Until lngStatus = 0
        Pause (20)  'wait 20 ms
        lngStatus = CommRead(intPortID, MsgStr, 10)
        rxBuffer = rxBuffer + MsgStr
        'increment timeout
        MBtimeout = MBtimeout + 1
        If MBtimeout > 10 Then
            Exit Do
            MBGenericMsg = 0  'timed out waiting for response
        End If
    Loop
    
End Function
Public Function CRC16A(sBuffer As String, iLength As Integer) As Long
'CRC calculations
'Copyright Richard L. Grier, 2006
'Modified Dave Loucks
    Dim i As Long
    Dim temp As Long
    Dim CRC As Long
    Dim j As Integer
        CRC = 0
      'For i = 0 To UBound(buffer) - 1
      For i = 0 To iLength - 1
        'Temp = buffer(i) * &H100&
        temp = Asc(Mid$(sBuffer, i + 1, 1)) * CLng(256)
        CRC = CRC Xor temp
          For j = 0 To 7
            If (CRC And 32768) Then
              CRC = CRC * 2
              CRC = CRC Xor 4129
              CRC = CRC And 65535
            Else
              CRC = (CRC * 2) And CLng(65535)
            End If
          Next j
      Next i
      CRC16A = CRC And &HFFFF
End Function
Public Function SolveMBCRC(sBuffer As String, iLength As Integer) As Long
    Dim CRC As Long
    Dim i As Integer
    Dim j As Integer
    Dim temp As Byte
    
    CRC = &HFFFF
    
    For i = 0 To iLength - 1
        temp = Asc(Mid$(sBuffer, i + 1, 1)) 'get next byte from MB message
        CRC = CRC Xor temp                  'XOR with LSB of CRC
        
        For j = 7 To 0 Step -1
            If (CRC And 1) Then
                CRC = CRC \ 2
                CRC = CRC Xor &HA001
            Else
                CRC = CRC \ 2
            End If
        Next
        
    Next
    
End Function

'-------------------------------------------------
Public Function ModCRC(buffer() As Byte, iLength As Integer) As Long
'-------------------------------------------------

' returns the MODBUS CRC of buffer
Dim CRC1 As Long
Dim i As Integer
Dim j As Integer
Dim K As Long
Dim temp As Long

    '*** test code
    Dim iLoopCount As Integer
    'initialize counter and timer
    iLoopCount = 0
    
  CRC1 = 65535 ' init CRC
  
  For i = 0 To iLength - 1 ' each byte
  CRC1 = CRC1 Xor buffer(i)
    For j = 0 To 7 ' for each bit in byte
    K = CRC1 And 1 ' bit 0 value
    CRC1 = ((CRC1 And 65534) / 2) And 32767 ' Shift right with 0 MSB
    If K > 0 Then CRC1 = CRC1 Xor 40961: iLoopCount = iLoopCount + 1
    iLoopCount = iLoopCount + 1
    Next j
  Next i
  
  ModCRC = CRC1 And 65535
  
  '*** write to CRCtiming sheet
  'Worksheets("CRCtiming").Range("a1").offset(giReportRow, 0) = iLength
  'Worksheets("CRCtiming").Range("a1").offset(giReportRow, 1) = iLoopCount
  'giReportRow = giReportRow + 1
  
End Function

Public Function VerifyCRC(buffer() As Byte, mb_count As Integer)
    'recalculate CRC over message and verify it matches CRC on message
    Dim CRC As Long
    Dim CRC1 As Byte
    Dim CRC2 As Byte
    Dim temp As Long

    If mb_count = 0 Then    'nothing received
        VerifyCRC = 0   'failed test
        'frmMain.rtb_Advisory.AppendText ("No response received to MB request" + vbCrLf)
        Exit Function
    End If

    CRC = ModCRC(buffer, mb_count - 2)
    temp = (CRC \ 256) And 255
    CRC2 = (CRC - temp * 256) And 255
    CRC1 = temp

    If buffer(mb_count - 1) <> CRC1 Then
        VerifyCRC = 0   'failed test
        Exit Function
    ElseIf buffer(mb_count - 2) <> CRC2 Then
        VerifyCRC = 0   'failed test
        Exit Function
    Else
        VerifyCRC = 1   'both bytes match.  passed test
    End If

End Function

Sub Pause(iTime As Integer)
'
' Delay specified number of milliseconds
'
    DoEvents
    AppSleep (iTime) ' delay iTime milliseconds

End Sub

Function openCOMPort(retCOMport As Integer, gretBPS As Integer) As Integer

    ' Open COM port
    openCOMPort = CommOpen(retCOMport, "COM" & CStr(retCOMport), _
        "baud=" & gretBPS & " parity=N data=8 stop=1")
        
    'Initialize timeout
    gMBTimeout = 480000 / gretBPS
                                '1200 -> 0.4s
                                '9600 -> 0.05s
                                '19.2 -> 0.025s
    MBTimeoutMAX = 1000 / gMBTimeout
        
End Function
Function closeCOMPort(retCOMport As Integer) As Integer

    closeCOMPort = retCOMport
    Call CommClose(closeCOMPort)
    
End Function

Public Function GetProxyInfoForUrl(Url As String) As yoasu20lsk

    Dim IEProxyConfig As WINHTTP_CURRENT_USER_IE_PROXY_CONFIG
    Dim AutoProxyOptions As WINHTTP_AUTOPROXY_OPTIONS
    Dim WinHttpProxyInfo As WINHTTP_PROXY_INFO
    Dim ProxyInfo As yoasu20lsk
    Dim fDoAutoProxy As Boolean
    Dim ProxyStringPtr As Long
    Dim ptr As Long
    Dim error As Long
    
    AutoProxyOptions.dwFlags = 0
    AutoProxyOptions.dwAutoDetectFlags = 0
    AutoProxyOptions.lpszAutoConfigUrl = 0
    AutoProxyOptions.dwReserved = 0
    AutoProxyOptions.lpvReserved = 0
    AutoProxyOptions.fAutoLogonIfChallenged = 1
    
    IEProxyConfig.fAutoDetect = 0
    IEProxyConfig.lpszAutoConfigUrl = 0
    IEProxyConfig.lpszProxy = 0
    IEProxyConfig.lpszProxyBypass = 0
    
    WinHttpProxyInfo.dwAccessType = 0
    WinHttpProxyInfo.lpszProxy = 0
    WinHttpProxyInfo.lpszProxyBypass = 0
    
    ProxyInfo.active = False
    ProxyInfo.proxy = vbNullString
    ProxyInfo.proxyBypass = vbNullString
    
    fDoAutoProxy = False
    ProxyStringPtr = 0
    ptr = 0
    
    ' Check IE's proxy configuration
    If (WinHttpGetIEProxyConfigForCurrentUser(IEProxyConfig) > 0) Then
    ' If IE is configured to auto-detect, then we will too.
    If (IEProxyConfig.fAutoDetect <> 0) Then
    AutoProxyOptions.dwFlags = WINHTTP_AUTOPROXY_AUTO_DETECT
    AutoProxyOptions.dwAutoDetectFlags = _
    WINHTTP_AUTO_DETECT_TYPE_DHCP + _
    WINHTTP_AUTO_DETECT_TYPE_DNS
    fDoAutoProxy = True
    End If
    
    ' If IE is configured to use an auto-config script, then
    ' we will use it too
    If (IEProxyConfig.lpszAutoConfigUrl <> 0) Then
    AutoProxyOptions.dwFlags = AutoProxyOptions.dwFlags + _
    WINHTTP_AUTOPROXY_CONFIG_URL
    AutoProxyOptions.lpszAutoConfigUrl = IEProxyConfig.lpszAutoConfigUrl
    fDoAutoProxy = True
    End If
    Else
    ' if the IE proxy config is not available, then
    ' we will try auto-detection
    AutoProxyOptions.dwFlags = WINHTTP_AUTOPROXY_AUTO_DETECT
    AutoProxyOptions.dwAutoDetectFlags = _
    WINHTTP_AUTO_DETECT_TYPE_DHCP + _
    WINHTTP_AUTO_DETECT_TYPE_DNS
    fDoAutoProxy = True
    End If
    
    If fDoAutoProxy Then
    Dim hSession As Long
    
    ' Need to create a temporary WinHttp session handle
    ' Note: performance of this GetProxyInfoForUrl function can be
    ' improved by saving this hSession handle across calls
    ' instead of creating a new handle each time
    hSession = WinHttpOpen(0, 1, 0, 0, 0)
    
    If (WinHttpGetProxyForUrl(hSession, StrPtr(Url), AutoProxyOptions, _
    WinHttpProxyInfo) > 0) Then
    ProxyStringPtr = WinHttpProxyInfo.lpszProxy
    ' ignore WinHttpProxyInfo.lpszProxyBypass, it will not be set
    Else
    error = Err.LastDllError
    ' some possibly autoproxy errors:
    ' 12166 - error in proxy auto-config script code
    ' 12167 - unable to download proxy auto-config script
    ' 12180 - WPAD detection failed
    End If
    
    WinHttpCloseHandle (hSession)
    End If
    
    ' If we don't have a proxy server from WinHttpGetProxyForUrl,
    ' then pick one up from the IE proxy config (if given)
    If (ProxyStringPtr = 0) Then
    ProxyStringPtr = IEProxyConfig.lpszProxy
    End If
    
    ' If there's a proxy string, convert it to a Basic string
    If (ProxyStringPtr <> 0) Then
    ptr = SysAllocString(ProxyStringPtr)
    CopyMemory VarPtr(ProxyInfo.proxy), VarPtr(ptr), 4
    ProxyInfo.active = True
    End If
    
    ' Pick up any bypass string from the IEProxyConfig
    If (IEProxyConfig.lpszProxyBypass <> 0) Then
    ptr = SysAllocString(IEProxyConfig.lpszProxyBypass)
    CopyMemory VarPtr(ProxyInfo.proxyBypass), VarPtr(ptr), 4
    End If
    
    ' Free any strings received from WinHttp APIs
    If (IEProxyConfig.lpszAutoConfigUrl <> 0) Then
    GlobalFree (IEProxyConfig.lpszAutoConfigUrl)
    End If
    If (IEProxyConfig.lpszProxy <> 0) Then
    GlobalFree (IEProxyConfig.lpszProxy)
    End If
    If (IEProxyConfig.lpszProxyBypass <> 0) Then
    GlobalFree (IEProxyConfig.lpszProxyBypass)
    End If
    If (WinHttpProxyInfo.lpszProxy <> 0) Then
    GlobalFree (WinHttpProxyInfo.lpszProxy)
    End If
    If (WinHttpProxyInfo.lpszProxyBypass <> 0) Then
    GlobalFree (WinHttpProxyInfo.lpszProxyBypass)
    End If
    
    ' return the ProxyInfo struct
    GetProxyInfoForUrl = ProxyInfo

End Function
