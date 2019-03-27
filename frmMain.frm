VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Screen Streaming"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Thread"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'Description:   The main window
'Author:        IceLolly
'File:          frmMain.frm
'====================================================

Option Explicit

Dim hCaptureThread              As Long, CaptureThreadTID       As Long     'Screen capture thread handle & ID

Public ResultDC                 As New clsMemDC                             'Result image data

'Description:   Initialize memory DC and start screen capture thread
Public Sub StartStreaming()
    Dim i           As Integer
    
    'Initalize memory DC
    For i = 0 To SPLIT_COUNT - 1
        DC1(i).CreateMemDC ScreenW, SplittedSize
        DC2(i).CreateMemDC ScreenW, SplittedSize
        DC3(i).CreateMemDC ScreenW, SplittedSize
        DC4(i).CreateMemDC ScreenW, SplittedSize
    Next i
    
    'Start screen capture thread
    hCaptureThread = CreateThread(0, 0, AddressOf ScreenCaptureThread, 0, 0, CaptureThreadTID)
End Sub

Private Sub cmdConnect_Click()
    ClientSocket = AllocateSocket()
    If ClientSocket = INVALID_SOCKET Then
        MsgBox "Failed to socket()!"
        Exit Sub
    End If
    If SocketBind(ClientSocket) = SOCKET_ERROR Then
        MsgBox "Failed to bind()!"
        Exit Sub
    End If
    If SocketConnect(ClientSocket, "127.0.0.1", 18787) = 0 Then
        MsgBox "Failed to create client thread!"
        Exit Sub
    End If
End Sub

Private Sub cmdStart_Click()
    ServerSocket = AllocateSocket()
    If ServerSocket = INVALID_SOCKET Then
        MsgBox "Failed to socket()!"
        Exit Sub
    End If
    If SocketBind(ServerSocket, 18787) = SOCKET_ERROR Then
        MsgBox "Failed to bind()!"
        Exit Sub
    End If
    If SocketListen(ServerSocket) = SOCKET_ERROR Then
        MsgBox "Failed to listen()!"
        Exit Sub
    End If
    hServerSocketThread = CreateThread(0, 0, AddressOf ServerSocketThread, ByVal ServerSocket, 0, ServerSocketTID)
End Sub

Private Sub cmdStop_Click()
    closesocket ServerSocket
    closesocket ServerConnectedSocket
    closesocket ClientSocket
    TerminateThread hServerSocketThread, 0
    TerminateThread hClientSocketThread, 0
End Sub

Private Sub Form_Load()
    ScreenDC = GetDC(0)                                                                 'Get screen DC
    ScreenW = Screen.Width / Screen.TwipsPerPixelX                                      'Calculate screen size
    ScreenH = Screen.Height / Screen.TwipsPerPixelY
    SplittedSize = ScreenH / SPLIT_COUNT
    
    'We need to initalize all memory DCs here to create new instances of clsMemDC
    'Cuz "New clsMemDC" in the thread procedures will crash the program
    'The sizes of DCs are (0, 0) in order to save memory
    'We will re-create these DCs in other procedures later while we won't need to "New clsMemDC" in these procedures
    Dim i           As Integer
    
    ResultDC.CreateMemDC 0, 0                                                           'Result image DC
    tempDC.CreateMemDC 0, 0                                                             'Temporary image DC
    ReDim DC1(SPLIT_COUNT - 1)                                                          'Initalize DC arrays for scan purpose
    ReDim DC2(SPLIT_COUNT - 1)
    ReDim DC3(SPLIT_COUNT - 1)
    ReDim DC4(SPLIT_COUNT - 1)
    For i = 0 To SPLIT_COUNT - 1                                                        'Initalize all DC in the arrays
        DC1(i).CreateMemDC 0, 0
        DC2(i).CreateMemDC 0, 0
        DC3(i).CreateMemDC 0, 0
        DC4(i).CreateMemDC 0, 0
    Next i
    ReDim ReplayDcBuffer(RECV_DELAY * FPS - 1)                                          'Initalize replay DC array
    ReDim RecvDcBuffer(RECV_DELAY * FPS - 1)                                            'Initalize received DC array
    For i = 0 To RECV_DELAY * FPS - 1                                                   'Initalize all DC in the arrays
        ReplayDcBuffer(i).CreateMemDC 0, 0
        RecvDcBuffer(i).CreateMemDC 0, 0
    Next i
    
    Dim wData   As WSADATA
    If WSAStartup(&H202, wData) <> 0 Then                                               'Startup WSA, &H202 = MAKEWORD(2, 2)
        MsgBox "WSAStartup() failed!", vbCritical, "Error"
        Unload Me
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    TerminateThread hServerSocketThread, 0
    TerminateThread hClientSocketThread, 0
    WSACleanup                                                                          'Close WSA
    ReleaseDC Me.hWnd, ScreenDC                                                         'Release screen DC
End Sub

