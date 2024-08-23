Attribute VB_Name = "CapCam"
'################################################################################################
' This module will initialize the camera and capture a picture from it.
'################################################################################################

Option Compare Database
Option Explicit

' Type CapStatus describes the camera specific settings
Public Type CapStatus
    uiImageWidth As Long                    '// Width of the image
    uiImageHeight As Long                   '// Height of the image
    fLiveWindow As Long                     '// Now Previewing video?
    fOverlayWindow As Long                  '// Now Overlaying video?
    fScale As Long                          '// Scale image to client?
    ptxScroll As Long                       '// Scroll position
    ptyScroll As Long                       '// Scroll position
    fUsingDefaultPalette As Long            '// Using default driver palette?
    fAudioHardware As Long                  '// Audio hardware present?
    fCapFileExists As Long                  '// Does capture file exist?
    dwCurrentVideoFrame As Long             '// # of video frames cap'td
    dwCurrentVideoFramesDropped As Long     '// # of video frames dropped
    dwCurrentWaveSamples As Long            '// # of wave samples cap'td
    dwCurrentTimeElapsedMS As Long          '// Elapsed capture duration
    hPalCurrent As Long                     '// Current palette in use
    fCapturingNow As Long                   '// Capture in progress?
    dwReturn As Long                        '// Error value after any operation
    wNumVideoAllocated As Long              '// Actual number of video buffers
    wNumAudioAllocated As Long              '// Actual number of audio buffers
End Type

' Constats to address the vdieo settings
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WM_USER As Long = &H400
Private Const WM_CAP_START As Long = WM_USER
Private Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP_START + 10
Private Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP_START + 11
Private Const WM_CAP_SET_PREVIEW As Long = WM_CAP_START + 50
Private Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP_START + 52
Private Const WM_CAP_SET_VIDEOFORMAT As Long = WM_CAP_START + 45
Private Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_CAP_START + 41
Private Const WM_CAP_FILE_SAVEDIB As Long = WM_CAP_START + 25
Private Const WM_CAP_GET_STATUS As Long = WM_CAP_START + 54

Private Declare PtrSafe Function capCreateCaptureWindow _
    Lib "avicap32.dll" Alias "capCreateCaptureWindowA" _
         (ByVal lpszWindowName As String, ByVal dwStyle As Long _
        , ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long _
        , ByVal nHeight As Long, ByVal hwndParent As LongPtr _
        , ByVal nID As Long) As Long

Private Declare PtrSafe Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long _
        , ByVal wParam As Long, ByRef lParam As Any) As Long

Dim hCap As LongPtr

Public Sub CreateCaptureWindow(ByRef picwin As Object, ByVal message As String)
    hCap = capCreateCaptureWindow(message, WS_CHILD Or WS_VISIBLE, 0, 0, picwin.Width, picwin.Height, picwin.Form.hWnd, 0)
    If hCap <> 0 Then
        Call SendMessage(hCap, WM_CAP_DRIVER_CONNECT, 0, 0)
        Call SendMessage(hCap, WM_CAP_SET_PREVIEWRATE, 66, 0&)
        Call SendMessage(hCap, WM_CAP_SET_PREVIEW, CLng(True), 0&)
    End If
End Sub

Public Function FormatPictureDlg() As Long
    FormatPictureDlg = SendMessage(hCap, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
End Function

Public Function Disconnect() As Long
    Disconnect = SendMessage(hCap, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
End Function

Public Function SavePicture(ByVal savePath As String) As Object
    Call SendMessage(hCap, WM_CAP_SET_PREVIEW, CLng(False), 0&)
    Call SendMessage(hCap, WM_CAP_FILE_SAVEDIB, 0&, ByVal CStr(savePath))
DoFinally:
    Call SendMessage(hCap, WM_CAP_SET_PREVIEW, CLng(True), 0&)
End Function

Public Function GetCapStatus() As CapStatus
    Dim s As CapStatus
    Call SendMessage(hCap, WM_CAP_GET_STATUS, Len(s), s)
    GetCapStatus = s
End Function


