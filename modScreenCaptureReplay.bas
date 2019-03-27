Attribute VB_Name = "modScreenCaptureReplay"
'====================================================
'Description:   Thread functions to capture screen and send captured data
'Author:        IceLolly
'File:          modScreenCaptureReplay.bas
'====================================================

Option Explicit

'Zlib compression functions
Public Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

Public Const SPLIT_COUNT            As Integer = 7

Public DC1()                        As New clsMemDC, _
       DC2()                        As New clsMemDC, _
       DC3()                        As New clsMemDC, _
       DC4()                        As New clsMemDC
Public ScreenW                      As Long, _
       ScreenH                      As Long                                             'Screen size (in pixels)
Public SplitCount                   As Long                                             'Split count of remote computer
Public SplittedSize                 As Long                                             'Screen splitted height
Public ScreenDC                     As Long                                             'Screen DC

'Description:   Thread to capture the screen
'Args:          Param: Unused
'Return:        Unused
Public Function ScreenCaptureThread(Param As Long) As Long
    CreateIExprSrvObj 0&, 4&, 0&                                                        'Initialize VB6 runtime library
    CoInitializeEx ByVal 0&, ByVal (COINIT_MULTITHREADED Or COINIT_SPEED_OVER_MEMORY)   'Initialize COM components
    InitVBdll                                                                           'Initialize VB6 multithreading environment
    '==============================
    On Error GoTo ErrHandler
    Debug.Print "ScreenCaptureThread (" & GetCurrentThreadId() & ") created"
    
    Dim tMsg            As MSG                                                          'Thread message
    Dim i               As Integer
    Dim CompData()      As Byte                                                         'The image data that is being compressed
    Dim CompDest()      As Byte                                                         'Compressed image data
    Dim szData          As Long                                                         'Size of the compressed data
    Dim szTotal         As Long                                                         'Total size of all compressed data
    
    '--------------------
    frmMain.ResultDC.CreateMemDC 1920, 1080
    '--------------------
    
    Do While tMsg.message <> WM_QUIT                                                    'Loop until the thread is told to quit
        'Capture
        szTotal = 0
        For i = 0 To SPLIT_COUNT - 1
            DC3(i).BitBltFrom ScreenDC, 0, SplittedSize * i, 0, 0, ScreenW, SplittedSize    'Get screen update
            DC3(i).BitBltTo DC1(i).hDC, 0, 0, 0, 0, ScreenW, SplittedSize
            DC4(i).BitBltTo DC2(i).hDC, 0, 0, 0, 0, ScreenW, SplittedSize
            If RtlCompareMemory(ByVal DC1(i).lpBitData, ByVal DC2(i).lpBitData, DC1(i).iImageSize) <> DC1(i).iImageSize Then
                DC1(i).BitBltTo DC4(i).hDC, 0, 0, 0, 0, ScreenW, SplittedSize
                DC2(i).BitBltTo DC1(i).hDC, 0, 0, 0, 0, ScreenW, SplittedSize, vbSrcInvert
                
                ReDim CompData(DC1(i).iImageSize + 1)                                       'Allocate buffer
                CopyMemory CompData(0), i, 2                                                'Copy current image index
                CopyMemory CompData(2), ByVal DC1(i).lpBitData, DC1(i).iImageSize           'Copy current image data
                szData = UBound(CompData) * 1.01 + 12                                       'Estimate max. size required
                ReDim Preserve CompDest(szData)                                             'Allocate compression buffer
                compress CompDest(0), szData, CompData(0), DC1(i).iImageSize + 2            'Compress
                ReDim Preserve CompDest(szData - 1)                                         'Shrink the compression buffer
                szTotal = szTotal + szData                                                  'Add up into total size
                
                'Check connection status
                If ServerConnectedSocket <> INVALID_SOCKET Then                             'Connection established
                    SendStreamData ServerConnectedSocket, CompDest
                Else                                                                            'Quit thread if not connected
                    Debug.Print "WARNING: Not connected!"
                    frmMain.Caption = "Streaming Stopped"
                    GoTo Exiting
                End If
            End If
        Next i
        
        frmMain.Caption = "Streaming: " & Format(szTotal / 1024, "0.00kb")
        Sleep 20
        
        'Handle message from thread message queue
        PeekMessageW tMsg, 0, 0, 0, 0
        If tMsg.message = WM_QUIT Then
            GoTo Exiting
        End If
    Loop

Exiting:
    Debug.Print "Capture thread (" & GetCurrentThreadId() & ") exited"

    '==============================
    CoUninitialize                                                                      'Unitializes COM components
    Exit Function

ErrHandler:
    If MessageBoxW(0, "Error!" & vbCrLf & "Thread ID = " & GetCurrentThreadId() & vbCrLf & "ScreenCaptureThread" & _
                    vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "Continue?", _
                    "Runtime Error", MB_ICONEXCLAMATION Or MB_YESNO) = IDYES Then
        Err.Clear
        Resume Next
    Else
        Err.Clear
        Exit Function
    End If
End Function

