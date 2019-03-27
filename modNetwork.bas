Attribute VB_Name = "modNetwork"
'====================================================
'Description:   Functions & Thread procedures to handle socket-related jobs
'Author:        IceLolly
'File:          modNetwork.bas
'Note:          This network module is designed for this program ONLY.
'               You may make some changes in order to put this module
'               into work for your program
'====================================================

Option Explicit

'Create a new socket
Public Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal lType As Long, _
    ByVal protocol As Long) As Long

'Bind socket
Public Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As SOCKADDR, _
    ByVal namelen As Long) As Long
    
'Port conversion
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer

'IP address conversion
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long

'Close socket
Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long

'Put the specified socket into listening status
Public Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long

'Wait and accept the connection
Public Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As SOCKADDR, _
    ByRef addrlen As Long) As Long

'Send packet for TCP sockets
Public Declare Function send Lib "ws2_32.dll" (ByVal s As Long, buf As Any, _
    ByVal lLen As Long, ByVal flags As Long) As Long

'Receive packet from TCP sockets
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, buf As Any, _
    ByVal lLen As Long, ByVal flags As Long) As Long

'Connect to remote host
Private Declare Function connect Lib "ws2_32.dll" (ByVal s As Long, _
    ByRef name As SOCKADDR, ByVal namelen As Long) As Long

'socket(), af
Public Const AF_INET = 2
'socket(), lType
Public Const TCP_NODELAY = &H1
'socket(), protocol
Public Const IPPROTO_TCP = 6

'bind(), addr, sin_addr
Public Const INADDR_ANY = &H0
'bind(), return value
Public Const SOCKET_ERROR = -1
'socket() or accept(), return value
Public Const INVALID_SOCKET = -1

'bind(), addr
Private Type SOCKADDR
    sin_family                  As Integer
    sin_port                    As Integer
    sin_addr                    As Long
    sin_zero                    As String * 8
End Type

Public Const BUFFER_SIZE = 1024 * 16                                                    'Buffer size = 16kb
Public Const RECV_DELAY = 2                                                             'Length of receive buffer (in seconds)
Public Const FPS = 24                                                                   'Replay FPS

Public ServerSocket             As Long, ServerConnectedSocket  As Long                 'Server socket handle
Public hServerSocketThread      As Long, ServerSocketTID        As Long                 'Server socket thread handle & ID
Public ClientSocket             As Long                                                 'Client socket handle
Public hClientSocketThread      As Long, ClientSocketTID        As Long                 'Client socket thread handle & ID
Public tempDC                   As New clsMemDC                                         'Temporary received image data
Public ReplayDcBuffer()         As New clsMemDC                                         'Replay image buffer
Public RecvDcBuffer()           As New clsMemDC                                         'Received image buffer

Dim RemoteAddr                  As SOCKADDR                                             'Remote address of the client socket

'Description:   Create a new socket
'Return:        A socket handle if succeed, INVALID_SOCKET if failed
Public Function AllocateSocket() As Long
    AllocateSocket = socket(AF_INET, TCP_NODELAY, IPPROTO_TCP)                          'Create a TCP socket
End Function

'Description:   Bind the socket
'Args:          TargetSocket: Specific a socket handle
'               Port: Port number
'Return:        SOCKET_ERROR if failed
Public Function SocketBind(TargetSocket As Long, Optional Port As Integer = 0) As Long
    Dim RemoteAddr              As SOCKADDR                                             'Socket address info
    
    'Set address info
    With RemoteAddr
        .sin_family = AF_INET
        .sin_port = htons(Port)
        .sin_addr = INADDR_ANY
    End With
    
    SocketBind = bind(TargetSocket, RemoteAddr, Len(RemoteAddr))                        'Bind socket with specified port
    If SocketBind = SOCKET_ERROR Then                                                   'Close the socket if bind() fails
        closesocket TargetSocket
    End If
End Function

'Description:   Put the socket into listening status
'Args:          TargetSocket: Specific a socket handle
'Return:        SOCKET_ERROR if failed
Public Function SocketListen(TargetSocket As Long) As Long
    SocketListen = listen(TargetSocket, 1)
    If SocketListen = SOCKET_ERROR Then
        closesocket TargetSocket
    End If
End Function

'Description:   Connect the specified socket to the remote host with provided IP and port number
'Args:          TargetSocket: Specific a socket handle
'               RemoteIP: IP of the remote host
'               RemotePort: Port number of the remote host
'Return:        Handle to the client socket thread
'Note:          Call SocketBind() to bind the socket before calling me
Public Function SocketConnect(TargetSocket As Long, RemoteIP As String, RemotePort As Integer) As Long
    'Set remote address info
    With RemoteAddr
        .sin_family = AF_INET
        .sin_addr = inet_addr(RemoteIP)
        .sin_port = htons(RemotePort)
    End With
    
    hClientSocketThread = CreateThread(0, 0, AddressOf ClientSocketThread, 0, 0, ClientSocketTID)
    SocketConnect = hClientSocketThread
End Function

'Description:   Send a "stream" packet with specified header from the specified socket
'Args:          SocketHandle: Socket handle
'               Data(): Byte array to be sent
'Return:        Unused
Public Function SendStreamData(TargetSocket As Long, Data() As Byte) As Long
    Dim SendBuffer()            As Byte                                                 'Temporary buffer
    Dim DataSize                As Long                                                 'Data size of Data()
    
    'Buffer graph:
    'Average value of first 4 bytes
    '        ¡ý
    '¡õ¡õ¡õ¡õ¡õ¡õ¡­
    '©¸©¤©¤©¼  ©¸©¤
    'Data Size  Data
    
    DataSize = UBound(Data)                                                             'Retrieve size of Data()
    If DataSize = -1 Then                                                               'Invalid array to be sent
        SendStreamData = -1
        Exit Function
    End If
    ReDim SendBuffer(DataSize + 5)                                                      'Allocate temp. buffer
    CopyMemory SendBuffer(0), DataSize, 4                                               'Add data size to the head of the buffer
    'Add average value of first 4 bytes
    SendBuffer(4) = CByte((CInt(SendBuffer(0)) + SendBuffer(1) + SendBuffer(2) + SendBuffer(3)) / 4)
    CopyMemory SendBuffer(5), Data(0), DataSize + 1                                     'Copy the data to the buffer
    send TargetSocket, SendBuffer(0), DataSize + 6, 0                                   'Send the packet
End Function

'Description:   Thread to handle server socket connections
'Args:          SocketHandle: Socket handle
'Return:        Unused
Public Function ServerSocketThread(ByVal SocketHandle As Long) As Long
    CreateIExprSrvObj 0&, 4&, 0&                                                        'Initialize VB6 runtime library
    CoInitializeEx ByVal 0&, ByVal (COINIT_MULTITHREADED Or COINIT_SPEED_OVER_MEMORY)   'Initialize COM components
    InitVBdll                                                                           'Initialize VB6 multithreading environment
    '==============================
    On Error GoTo ErrHandler
    Debug.Print "ServerSocketThread (" & GetCurrentThreadId() & ") created"
    
    Dim szSockAddr              As Long                                                 'Size of SOCKADDR
    Dim ConnectedSocket         As Long                                                 'Handle to the connected socket
    Dim sAddr                   As SOCKADDR                                             'Remote address
    Dim MainBuffer()            As Byte                                                 'Main data buffer, dynamic size
    Dim RecvBuffer()            As Byte                                                 'Buffer to receive data, fixed size
    Dim szData                  As Long                                                 'Real size of the received data
    Dim PrevPos                 As Long                                                 'Writing position of main data buffer
    Dim szBufDataSize           As Long                                                 'To store packet size info read from the main buffer
    Dim BufferTempData()        As Byte                                                 'Complete packet read from the main buffer
    
    szSockAddr = Len(sAddr)
    ConnectedSocket = accept(SocketHandle, sAddr, szSockAddr)                           'Wait and accept the connection
    'accept() blocks the thread until there is a connection request
    
    closesocket SocketHandle                                                            'Close the listening socket
    If ConnectedSocket = INVALID_SOCKET Then                                            'Check if the connection is successful
        GoTo Exiting
    End If
    ServerConnectedSocket = ConnectedSocket
    
    frmMain.Caption = "Connected!"
    'Send streaming info
    ReDim StreamingInfo(9) As Byte                                                      'We can't use fixed length arrays directly in VB6 thread procedures
    CopyMemory StreamingInfo(0), ScreenW, 4
    CopyMemory StreamingInfo(4), ScreenH, 4
    CopyMemory StreamingInfo(8), SPLIT_COUNT, 2
    SendStreamData ConnectedSocket, StreamingInfo
    
    ReDim RecvBuffer(BUFFER_SIZE)                                                       'Allocate receive buffer
    ReDim MainBuffer(0)                                                                 'Initalize main buffer
    Do
        FillMemory RecvBuffer(0), BUFFER_SIZE, 0                                            'Clear receive buffer
        szData = recv(ConnectedSocket, RecvBuffer(0), BUFFER_SIZE, 0)                       'Receive data from the socket
        'recv() blocks the thread until there is a packet arrives or error occurs
        
        If szData <> SOCKET_ERROR And szData > 0 Then                                       'Check if there's no error
            PrevPos = UBound(MainBuffer)                                                        'Mark writing position of main buffer
            
            'Allocate main buffer memory, then copy data to it
            If PrevPos = 0 Then                                                                 'The first packet in the buffer
                ReDim Preserve MainBuffer(szData)
                CopyMemory MainBuffer(0), RecvBuffer(0), szData
            Else
                ReDim Preserve MainBuffer(PrevPos + szData)
                CopyMemory MainBuffer(PrevPos + 1), RecvBuffer(0), szData
            End If
            
            'Check main buffer size
            If UBound(MainBuffer) > 5 Then
                'Check average value of first 4 bytes
                If MainBuffer(4) <> CByte((CInt(MainBuffer(0)) + MainBuffer(1) + MainBuffer(2) + MainBuffer(3)) / 4) Then
                    Err.Raise 60001, , "Invalid packet header!"
                End If
            Else
                Err.Raise 60002, , "Main buffer size is less than 5 bytes!"
            End If
            
            CopyMemory szBufDataSize, MainBuffer(0), 4                                          'Get the size of first packet in the buffer
            Do While UBound(MainBuffer) >= szBufDataSize + 5
                ReDim BufferTempData(szBufDataSize)                                                 'Allocate temp buffer
                CopyMemory BufferTempData(0), MainBuffer(5), szBufDataSize + 1                      'Read the packet data from the buffer
                '---------------------------------------------------
                
                If szBufDataSize = 1 Then                                                           'Response from the remote computer
                    If StrConv(BufferTempData, vbUnicode) = "OK" Then                                   'The remote computer is ready, start streaming
                        frmMain.StartStreaming
                    Else
                        Err.Raise 60003, , "Unknown command received!"
                    End If
                End If
                
                '---------------------------------------------------
                If UBound(MainBuffer) - szBufDataSize - 6 > -1 Then                                 'If this is not the last packet in the buffer
                    CopyMemory MainBuffer(0), MainBuffer(6 + szBufDataSize), _
                        UBound(MainBuffer) - szBufDataSize - 5                                          'Delete the handled packet from the buffer
                    ReDim Preserve MainBuffer(UBound(MainBuffer) - szBufDataSize - 6)                   'Shrink the buffer
                    CopyMemory szBufDataSize, MainBuffer(0), 4                                          'Get the size of next packet
                Else                                                                                'If this is the last packet in the buffer
                    ReDim MainBuffer(0)                                                                 'Clear the buffer
                    Exit Do                                                                             'Exit the loop
                End If
            Loop
        Else                                                                                'Socket closes or error occurs
            frmMain.Caption = "Disconnected!"
            Exit Do
        End If
    Loop
    
Exiting:
    closesocket ConnectedSocket                                                         'Close the connected socket
    ServerConnectedSocket = INVALID_SOCKET                                              'Mark the socket as INVALID_SOCKET
    Debug.Print "ServerSocket thread (" & GetCurrentThreadId() & ") exited"
    
    '==============================
    CoUninitialize                                                                      'Unitialize COM components
    Exit Function

ErrHandler:
    If MessageBoxW(0, "Error!" & vbCrLf & "Thread ID = " & GetCurrentThreadId() & vbCrLf & "ServerSocketThread" & _
                    vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "Continue?", _
                    "Runtime Error", MB_ICONEXCLAMATION Or MB_YESNO) = IDYES Then
        Err.Clear
        Resume Next
    Else
        Err.Clear
        Exit Function
    End If
End Function

'Description:   Thread to handle client socket connections
'Args:          SocketInfo: Connection info of the client socket
'Return:        Unused
Public Function ClientSocketThread(Param As Long) As Long
    CreateIExprSrvObj 0&, 4&, 0&                                                        'Initialize VB6 runtime library
    CoInitializeEx ByVal 0&, ByVal (COINIT_MULTITHREADED Or COINIT_SPEED_OVER_MEMORY)   'Initialize COM components
    InitVBdll                                                                           'Initialize VB6 multithreading environment
    '==============================
    'On Error GoTo ErrHandler
    Debug.Print "ClientSocketThread (" & GetCurrentThreadId() & ") created"
    
    Dim ret                     As Long                                                 'Return value of functions
    Dim RecvBuffer()            As Byte                                                 'Buffer to receive data, fixed size
    Dim PrevPos                 As Long                                                 'Writing position of main data buffer
    Dim MainBuffer()            As Byte                                                 'Main data buffer, dynamic size
    Dim szBufDataSize           As Long                                                 'To store packet size info read from the main buffer
    Dim BufferTempData()        As Byte                                                 'Complete packet read from the main buffer
    Dim ImageSize               As Long                                                 'Estimated required decompression size of per image (in bytes)
    Dim tempDecompBuffer()      As Byte                                                 'Temporary decompression buffer
    Dim szDecomp                As Long                                                 'Size of decompressed data
    Dim ImageIndex              As Integer                                              'Image index
    Dim CurrBufferIndex         As Integer                                              'Index of replay buffer
    Dim i                       As Integer
    
    ret = connect(ClientSocket, RemoteAddr, ByVal Len(RemoteAddr))
    If ret = SOCKET_ERROR Or ClientSocket = INVALID_SOCKET Then                         'Check if the operation is successful
        GoTo Exiting
    End If
    
    frmMain.Caption = "Connected!"
    ReDim RecvBuffer(BUFFER_SIZE)                                                       'Allocate receive buffer
    ReDim MainBuffer(0)                                                                 'Initalize main buffer
    Do
        FillMemory RecvBuffer(0), BUFFER_SIZE, 0                                            'Clear receive buffer
        ret = recv(ClientSocket, RecvBuffer(0), BUFFER_SIZE, 0)                             'Receive data from the socket
        'recv() blocks the thread until there is a packet arrives or error occurs
        
        If ret <> SOCKET_ERROR And ret > 0 Then                                             'Check if there's no error
            PrevPos = UBound(MainBuffer)                                                        'Mark writing position of main buffer
            
            'Allocate main buffer memory, then copy data to it
            If PrevPos = 0 Then                                                                 'The first packet in the buffer
                ReDim Preserve MainBuffer(ret - 1)
                CopyMemory MainBuffer(0), RecvBuffer(0), ret
            Else
                ReDim Preserve MainBuffer(PrevPos + ret)
                CopyMemory MainBuffer(PrevPos + 1), RecvBuffer(0), ret
            End If
            
            'Check main buffer size
            If UBound(MainBuffer) > 5 Then
                'Check average value of first 4 bytes
                If MainBuffer(4) <> CByte((CInt(MainBuffer(0)) + MainBuffer(1) + MainBuffer(2) + MainBuffer(3)) / 4) Then
                    Err.Raise 60001, , "Invalid packet header!"
                End If
            Else
                Err.Raise 60002, , "Main buffer size is less than 5 bytes!"
            End If
            
            CopyMemory szBufDataSize, MainBuffer(0), 4                                          'Get the size of first packet in the buffer
            Do While UBound(MainBuffer) >= szBufDataSize + 5
                ReDim BufferTempData(szBufDataSize)                                                 'Allocate temp buffer
                CopyMemory BufferTempData(0), MainBuffer(5), szBufDataSize + 1                      'Read the packet data from the buffer
                '---------------------------------------------------
                
                If szBufDataSize = 9 Then                                                           'Streaming info received
                    CopyMemory ScreenW, BufferTempData(0), 4                                            'Screen weight
                    CopyMemory ScreenH, BufferTempData(4), 4                                            'Screen height
                    CopyMemory SplitCount, BufferTempData(8), 2                                         'Split count
                    
                    SplittedSize = ScreenH / SplitCount
                    ImageSize = ScreenW * SplittedSize * 16 / 8 + 2                                     'Calc. est. required decompression size (2 extra bytes to store image index)
                    frmMain.ResultDC.CreateMemDC ScreenW, ScreenH                                       'Initialize memory DC
                    tempDC.CreateMemDC ScreenW, SplittedSize                                            'Initialize temporary memory DC
                    
                    For i = 0 To RECV_DELAY * FPS - 1                                                   'Initalize DC buffers
                        ReplayDcBuffer(i).CreateMemDC ScreenW, ScreenH
                        RecvDcBuffer(i).CreateMemDC ScreenW, ScreenH
                    Next i
                    
                    SendStreamData ClientSocket, StrConv("OK", vbFromUnicode)                           'Response
                Else                                                                                'Image data received
                    ReDim tempDecompBuffer((ImageSize + 2) * 1.01 + 12)                                 'Allocate decompression buffer
                    szDecomp = UBound(tempDecompBuffer) + 1
                    uncompress tempDecompBuffer(0), szDecomp, _
                        BufferTempData(0), UBound(BufferTempData) + 1                                   'Uncompress
                    ReDim Preserve tempDecompBuffer(szDecomp - 1)                                       'Shrink the decompression buffer
                    CopyMemory ImageIndex, tempDecompBuffer(0), 2                                       'Get image index
                    CopyMemory ByVal tempDC.lpBitData, tempDecompBuffer(2), tempDC.iImageSize           'Get image data
                    
                    tempDC.BitBltTo frmMain.ResultDC.hDC, 0, SplittedSize * ImageIndex, 0, 0, _
                        ScreenW, SplittedSize, vbSrcInvert                                              'Paint the scan image to the result image
                    'frmMain.ResultDC.BitBltTo frmMain.hDC, 0, 0, 0, 0, ScreenW, ScreenH                 'Display the result image
                    frmMain.ResultDC.BitBltTo RecvBuffer(CurrBufferIndex), 0, 0, 0, 0, ScreenW, ScreenH 'Paint the result image to received buffer
                    '[ToDo] Maybe don't use full screen buffer, use a struct that stores ImageIndex and splitted screen image, this saves more memory
                    CurrBufferIndex = CurrBufferIndex + 1
                End If
                
                '---------------------------------------------------
                If UBound(MainBuffer) - szBufDataSize - 6 > -1 Then                                 'If this is not the last packet in the buffer
                    CopyMemory MainBuffer(0), MainBuffer(6 + szBufDataSize), _
                        UBound(MainBuffer) - szBufDataSize - 5                                          'Delete the handled packet from the buffer
                    ReDim Preserve MainBuffer(UBound(MainBuffer) - szBufDataSize - 6)                   'Shrink the buffer
                    CopyMemory szBufDataSize, MainBuffer(0), 4                                          'Get the size of next packet
                Else                                                                                'If this is the last packet in the buffer
                    ReDim MainBuffer(0)                                                                 'Clear the buffer
                    Exit Do                                                                             'Exit the loop
                End If
            Loop
        Else                                                                                'Socket closes or error occurs
            frmMain.Caption = "Disconnected!"
            Exit Do
        End If
    Loop
    
Exiting:
    closesocket ClientSocket                                                            'Close the connected socket
    tempDC.DeleteMemDC                                                                  'Delete the temporary memory DC
    Debug.Print "ClientSocketThread thread (" & GetCurrentThreadId() & ") exited"
    
    '==============================
    CoUninitialize                                                                      'Unitialize COM components
    Exit Function

ErrHandler:
    If MessageBoxW(0, "Error!" & vbCrLf & "Thread ID = " & GetCurrentThreadId() & vbCrLf & "ClientSocketThread" & _
                    vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "Continue?", _
                    "Runtime Error", MB_ICONEXCLAMATION Or MB_YESNO) = IDYES Then
        Err.Clear
        Resume Next
    Else
        Err.Clear
        Exit Function
    End If
End Function
