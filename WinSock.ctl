VERSION 5.00
Begin VB.UserControl WinSock 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   390
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   375
   ScaleWidth      =   390
   Windowless      =   -1  'True
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "WinSock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Download by http://www.codefans.net
Private Declare Function api_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function api_htons Lib "ws2_32.dll" Alias "htons" (ByVal hostshort As Integer) As Integer
Private Declare Function api_ntohs Lib "ws2_32.dll" Alias "ntohs" (ByVal netshort As Integer) As Integer
Private Declare Function api_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function api_gethostname Lib "ws2_32.dll" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Private Declare Function api_gethostbyname Lib "ws2_32.dll" Alias "gethostbyname" (ByVal host_name As String) As Long
Private Declare Function api_bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function api_getsockname Lib "ws2_32.dll" Alias "getsockname" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Private Declare Function api_getpeername Lib "ws2_32.dll" Alias "getpeername" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Private Declare Function api_inet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function api_send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function api_sendto Lib "ws2_32.dll" Alias "sendto" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef toaddr As sockaddr_in, ByVal tolen As Long) As Long
Private Declare Function api_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function api_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function api_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function api_recvfrom Lib "ws2_32.dll" Alias "recvfrom" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef from As sockaddr_in, ByRef fromlen As Long) As Long
Private Declare Function api_WSACancelAsyncRequest Lib "ws2_32.dll" Alias "WSACancelAsyncRequest" (ByVal hAsyncTaskHandle As Long) As Long
Private Declare Function api_listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function api_accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As sockaddr_in, ByRef addrlen As Long) As Long
Private Declare Function api_inet_ntoa Lib "ws2_32.dll" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function api_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
Private Declare Function api_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal strHostName As String, buf As Any, ByVal buflen As Long) As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function LenA Lib "KERNEL32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function api_lstrcpy Lib "KERNEL32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "KERNEL32" (ByVal hLibModule As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetModuleHandle Lib "KERNEL32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "KERNEL32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Type tSubData
    hWnd                    As Long
    nAddrSub                As Long
    nAddrOrig               As Long
    nMsgCntA                As Long
    nMsgCntB                As Long
    aMsgTblA()              As Long
    aMsgTblB()              As Long
End Type
Private aBuf(1 To 200)      As Byte
Private sc_aSubData()       As tSubData
Private bTrack              As Boolean
Private bTrackUser32        As Boolean
Private bInCtrl             As Boolean
Public Enum SockState
    sckClosed = 0
    sckOpen
    sckListening
    sckConnectionPending
    sckResolvingHost
    sckHostResolved
    sckConnecting
    sckConnected
    sckClosing
    sckError
End Enum
Public Enum ProtocolConstants
    sckTCPProtocol = 0
    sckUDPProtocol = 1
End Enum
Private m_blnInitiated      As Boolean
Private m_lngSocksQuantity  As Long
Private m_lngWindowHandle   As Long
Private m_lngSocketHandle   As Long
Private m_enmState          As SockState
Private m_strTag            As String
Private m_strRemoteHost     As String
Private m_lngRemotePort     As Long
Private m_strRemoteHostIP   As String
Private m_lngLocalPort      As Long
Private m_lngLocalPortBind  As Long
Private m_strLocalIP        As String
Private m_enmProtocol       As ProtocolConstants
Private m_lngMemoryPointer  As Long
Private m_lngMemoryHandle   As Long
Private m_lngSendBufferLen  As Long
Private m_lngRecvBufferLen  As Long
Private m_strSendBuffer     As String
Private m_strRecvBuffer     As String
Private m_blnAcceptClass    As Boolean
Private m_colAcceptList     As Collection
Private m_colSocketsInst    As Collection
Private Type WSAData
    wVersion                As Integer
    wHighVersion            As Integer
    szDescription(256)      As Byte
    szSystemStatus(128)     As Byte
    iMaxSockets             As Integer
    iMaxUdpDg               As Integer
    lpVendorInfo            As Long
End Type
Private Type HOSTENT
    hName                   As Long
    hAliases                As Long
    hAddrType               As Integer
    hLength                 As Integer
    hAddrList               As Long
End Type
Private Type sockaddr_in
    sin_family              As Integer
    sin_port                As Integer
    sin_addr                As Long
    sin_zero(1 To 8)        As Byte
End Type
Public Event CloseSck()
Public Event Connect()
Public Event ConnectionRequest(ByVal requestID As Long)
Public Event DataArrival(ByVal bytesTotal As Long)
Public Event Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event SendComplete()
Public Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    Select Case uMsg
        Case 32768
            PostResolution wParam, HiWord(lParam)
        Case 32769
            PostSocket (lParam And &HFFFF&), HiWord(lParam)
    End Select
End Sub

Public Property Get RemotePort() As Long
    RemotePort = m_lngRemotePort
End Property

Public Property Let RemotePort(ByVal lngPort As Long)
    If Not (lngPort < 0 Or lngPort > 65535) Then
        m_lngRemotePort = lngPort
    End If
End Property

Public Property Get RemoteHost() As String
    RemoteHost = m_strRemoteHost
End Property

Public Property Let RemoteHost(ByVal strHost As String)
    m_strRemoteHost = strHost
End Property

Public Property Get RemoteHostIP() As String
    RemoteHostIP = m_strRemoteHostIP
End Property

Public Property Get LocalPort() As Long
    If m_lngLocalPortBind = 0 Then
        LocalPort = m_lngLocalPort
    Else
        LocalPort = m_lngLocalPortBind
    End If
End Property

Public Property Let LocalPort(ByVal lngPort As Long)
    If Not (lngPort < 0 Or lngPort > 65535) Then
        m_lngLocalPort = lngPort
    End If
End Property

Public Property Get State() As SockState
    State = m_enmState
End Property

Public Property Get LocalHostName() As String
    LocalHostName = GetLocalHostName
End Property

Public Property Get LocalIP() As String
    If m_enmState = sckConnected Then
        LocalIP = m_strLocalIP
    Else
        LocalIP = GetLocalIP
    End If
End Property

Public Property Get BytesReceived() As Long
    If m_enmProtocol = sckTCPProtocol Then
        BytesReceived = LenA(m_strRecvBuffer)
    Else
        BytesReceived = GetBufferLenUDP
    End If
End Property

Public Property Get SocketHandle() As Long
    SocketHandle = m_lngSocketHandle
End Property

Public Property Get Tag() As String
    Tag = m_strTag
End Property

Public Property Let Tag(ByVal strTag As String)
    m_strTag = strTag
End Property

Public Property Get Protocol() As ProtocolConstants
    Protocol = m_enmProtocol
End Property

Public Property Let Protocol(ByVal enmProtocol As ProtocolConstants)
    If m_enmState = sckClosed Then
        m_enmProtocol = enmProtocol
    End If
End Property

Private Sub DestroySocket()
    If Not m_lngSocketHandle = -1 Then
        Dim lngResult As Long
        lngResult = api_closesocket(m_lngSocketHandle)
        If lngResult = -1 Then
            m_enmState = sckError
            Dim lngErrorCode As Long
            lngErrorCode = Err.LastDllError
        Else
            m_lngSocketHandle = -1
        End If
    End If
End Sub

Public Sub CloseSck()
    If m_lngSocketHandle = -1 Then Exit Sub
    m_enmState = sckClosing
    CleanResolutionSystem
    DestroySocket
    m_lngLocalPortBind = 0
    m_strRemoteHostIP = ""
    m_strRecvBuffer = ""
    m_strSendBuffer = ""
    m_lngSendBufferLen = 0
    m_lngRecvBufferLen = 0
    m_enmState = sckClosed
End Sub

Private Function SocketExists() As Boolean
    SocketExists = True
    Dim lngResult As Long
    Dim lngErrorCode As Long
    If m_lngSocketHandle = -1 Then
        If m_enmProtocol = sckTCPProtocol Then
            lngResult = api_socket(2, 1, 6)
        Else
            lngResult = api_socket(2, 2, 17)
        End If
        If lngResult = -1 Then
            m_enmState = sckError
            SocketExists = False
            lngErrorCode = Err.LastDllError
            EventMsg lngErrorCode, "WinSock.SocketExists"
        Else
            m_lngSocketHandle = lngResult
            ProcessOptions
            SocketExists = RegisterSocket(m_lngSocketHandle, True)
        End If
    End If
End Function

Public Sub Connect(Optional RemoteHost As Variant, Optional RemotePort As Variant)
    If Not IsMissing(RemoteHost) Then
        m_strRemoteHost = CStr(RemoteHost)
    End If
    If m_strRemoteHost = vbNullString Then
        m_strRemoteHost = ""
    End If
    If Not IsMissing(RemotePort) Then
        If IsNumeric(RemotePort) Then
            If Not (CLng(RemotePort) > 65535 Or CLng(RemotePort) < 1) Then
                m_lngRemotePort = CLng(RemotePort)
            End If
        End If
    End If
    If Not SocketExists Then Exit Sub
    If Not BindInternal Then Exit Sub
    If m_enmProtocol = sckUDPProtocol Then
        m_enmState = sckOpen
        Exit Sub
    End If
    Dim lngAddress As Long
    lngAddress = ResolveIfHostname(m_strRemoteHost)
    If lngAddress <> vbNull Then
        ConnectToIP lngAddress, 0
    End If
End Sub

Private Sub PostResolution(ByVal lngAsynHandle As Long, ByVal lngErrorCode As Long)
    If m_enmState <> sckResolvingHost Then Exit Sub
    If lngErrorCode = 0 Then
        m_enmState = sckHostResolved
        Dim udtHostent As HOSTENT
        Dim lngPtrToIP As Long
        Dim arrIpAddress(1 To 4) As Byte
        Dim lngRemoteHostAddress As Long
        Dim Count As Integer
        Dim strIpAddress As String
        CopyMemory udtHostent, ByVal m_lngMemoryPointer, LenB(udtHostent)
        CopyMemory lngPtrToIP, ByVal udtHostent.hAddrList, 4
        CopyMemory arrIpAddress(1), ByVal lngPtrToIP, 4
        CopyMemory lngRemoteHostAddress, ByVal lngPtrToIP, 4
        FreeMemory
        For Count = 1 To 4
            strIpAddress = strIpAddress & arrIpAddress(Count) & "."
        Next
        strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)
        ConnectToIP lngRemoteHostAddress, 0
    Else
        FreeMemory
        ConnectToIP vbNull, lngErrorCode
    End If
End Sub

Private Sub PostSocket(ByVal lngEventID As Long, ByVal lngErrorCode As Long)
    If lngErrorCode <> 0 Then
        m_enmState = sckError
        EventMsg lngErrorCode, "WinSock.PostSocket"
        Exit Sub
    End If
    Dim udtSockAddr As sockaddr_in
    Dim lngResult As Long
    Dim lngBytesReceived As Long
    Select Case lngEventID
        Case 16
            If m_enmState <> sckConnecting Then
                Exit Sub
            End If
            GetLocalInfo m_lngSocketHandle, m_lngLocalPortBind, m_strLocalIP
            GetRemoteInfo m_lngSocketHandle, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
            m_enmState = sckConnected
            RaiseEvent Connect
        Case 2
            If m_enmState <> sckConnected Then
                Exit Sub
            End If
            If Len(m_strSendBuffer) > 0 Then
                SendBufferedData
            End If
        Case 1
            If m_enmProtocol = sckTCPProtocol Then
                If m_enmState <> sckConnected Then
                    Exit Sub
                End If
                lngBytesReceived = RecvDataToBuffer
                If lngBytesReceived > 0 Then
                    RaiseEvent DataArrival(LenA(m_strRecvBuffer))
                End If
            Else
                If m_enmState <> sckOpen Then
                    Exit Sub
                End If
                lngBytesReceived = GetBufferLenUDP
                If lngBytesReceived > 0 Then
                    RaiseEvent DataArrival(lngBytesReceived)
                End If
                EmptyBuffer
            End If
        Case 8
            If m_enmState <> sckListening Then
                Exit Sub
            End If
            lngResult = api_accept(m_lngSocketHandle, udtSockAddr, LenB(udtSockAddr))
            If lngResult = -1 Then
                lngErrorCode = Err.LastDllError
                m_enmState = sckError
                EventMsg lngErrorCode, "WinSock.PostSocket"
            Else
                RegisterAccept lngResult
                Dim lngTempRP As Long
                Dim strTempRHIP As String
                Dim strTempRH As String
                lngTempRP = m_lngRemotePort
                strTempRHIP = m_strRemoteHostIP
                strTempRH = m_strRemoteHost
                m_strRemoteHost = ""
                GetRemoteInfo lngResult, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
                RaiseEvent ConnectionRequest(lngResult)
                If m_enmState = sckListening Then
                    m_lngRemotePort = lngTempRP
                    m_strRemoteHostIP = strTempRHIP
                    m_strRemoteHost = strTempRH
                End If
                If IsAcceptRegistered(lngResult) Then
                    api_closesocket lngResult
                End If
            End If
        Case &H20
            If m_enmState <> sckConnected Then
                Exit Sub
            End If
            m_enmState = sckClosing
            RaiseEvent CloseSck
    End Select
End Sub

Private Sub ConnectToIP(ByVal lngRemoteHostAddress As Long, ByVal lngErrorCode As Long)
    If lngErrorCode <> 0 Then
        m_enmState = sckError
        EventMsg lngErrorCode, "WinSock.ConnectToIP"
        Exit Sub
    End If
    m_enmState = sckConnecting
    Dim udtSockAddr As sockaddr_in
    Dim lngResult As Long
    With udtSockAddr
        .sin_addr = lngRemoteHostAddress
        .sin_family = 2
        .sin_port = api_htons(UnsignedToInteger(m_lngRemotePort))
    End With
    lngResult = api_connect(m_lngSocketHandle, udtSockAddr, LenB(udtSockAddr))
    If lngResult = -1 Then
        lngErrorCode = Err.LastDllError
        If lngErrorCode <> 10035 Then
            If lngErrorCode <> 10049 Then
                m_enmState = sckError
                EventMsg lngErrorCode, "WinSock.ConnectToIP"
            End If
        End If
    End If
End Sub

Public Sub Bind(Optional LocalPort As Variant, Optional LocalIP As Variant)
    If BindInternal(LocalPort, LocalIP) Then
        m_enmState = sckOpen
    End If
End Sub

Private Function BindInternal(Optional ByVal varLocalPort As Variant, Optional ByVal varLocalIP As Variant) As Boolean
    If m_enmState = sckOpen Then
        BindInternal = True
        Exit Function
    End If
    Dim lngLocalPortInternal As Long
    Dim strLocalHostInternal As String
    Dim strIP As String
    Dim lngAddressInternal As Long
    Dim lngResult As Long
    Dim lngErrorCode As Long
    BindInternal = False
    If Not IsMissing(varLocalPort) Then
        If IsNumeric(varLocalPort) Then
            If varLocalPort < 0 Or varLocalPort > 65535 Then
                BindInternal = False
            Else
                lngLocalPortInternal = CLng(varLocalPort)
            End If
        Else
            BindInternal = False
        End If
    Else
        lngLocalPortInternal = m_lngLocalPort
    End If
    If Not IsMissing(varLocalIP) Then
        If varLocalIP <> vbNullString Then
            strLocalHostInternal = CStr(varLocalIP)
        Else
            strLocalHostInternal = ""
        End If
    Else
        strLocalHostInternal = ""
    End If
    lngAddressInternal = ResolveIfHostnameSync(strLocalHostInternal, strIP, lngResult)
    If Not SocketExists Then Exit Function
    Dim udtSockAddr As sockaddr_in
    With udtSockAddr
        .sin_addr = lngAddressInternal
        .sin_family = 2
        .sin_port = api_htons(UnsignedToInteger(lngLocalPortInternal))
    End With
    lngResult = api_bind(m_lngSocketHandle, udtSockAddr, LenB(udtSockAddr))
    If lngResult = -1 Then
        lngErrorCode = Err.LastDllError
    Else
        If lngLocalPortInternal <> 0 Then
            m_lngLocalPort = lngLocalPortInternal
        Else
            lngResult = GetLocalPort(m_lngSocketHandle)
            If lngResult = -1 Then
                lngErrorCode = Err.LastDllError
            Else
                m_lngLocalPortBind = lngResult
            End If
        End If
        BindInternal = True
    End If
End Function

Private Function AllocateMemory() As Long
    m_lngMemoryHandle = GlobalAlloc(0&, 1024)
    If m_lngMemoryHandle <> 0 Then
        m_lngMemoryPointer = GlobalLock(m_lngMemoryHandle)
        If m_lngMemoryPointer <> 0 Then
            GlobalUnlock (m_lngMemoryHandle)
            AllocateMemory = m_lngMemoryPointer
        Else
            GlobalFree (m_lngMemoryHandle)
            AllocateMemory = m_lngMemoryPointer
        End If
    Else
        AllocateMemory = m_lngMemoryHandle
    End If
End Function

Private Sub FreeMemory()
    If m_lngMemoryHandle <> 0 Then
        m_lngMemoryPointer = 0
        GlobalFree m_lngMemoryHandle
        m_lngMemoryHandle = 0
    End If
End Sub

Private Function GetLocalHostName() As String
    Dim strHostNameBuf As String * 256
    Dim lngResult As Long
    lngResult = api_gethostname(strHostNameBuf, 256)
    If lngResult = -1 Then
        GetLocalHostName = vbNullString
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
    Else
        GetLocalHostName = Left(strHostNameBuf, InStr(1, strHostNameBuf, vbNullChar) - 1)
    End If
End Function

Private Function GetLocalIP() As String
    Dim lngResult As Long
    Dim lngPtrToIP As Long
    Dim strLocalHost As String
    Dim arrIpAddress(1 To 4) As Byte
    Dim Count As Integer
    Dim udtHostent As HOSTENT
    Dim strIpAddress As String
    strLocalHost = GetLocalHostName
    lngResult = api_gethostbyname(strLocalHost)
    If lngResult = 0 Then
        GetLocalIP = vbNullString
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
    Else
        CopyMemory udtHostent, ByVal lngResult, LenB(udtHostent)
        CopyMemory lngPtrToIP, ByVal udtHostent.hAddrList, 4
        CopyMemory arrIpAddress(1), ByVal lngPtrToIP, 4
        For Count = 1 To 4
            strIpAddress = strIpAddress & arrIpAddress(Count) & "."
        Next
        strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)
        GetLocalIP = strIpAddress
    End If
End Function

Private Function ResolveIfHostname(ByVal Host As String) As Long
    Dim lngAddress As Long
    lngAddress = api_inet_addr(Host)
    If lngAddress = &HFFFF Then
        ResolveIfHostname = vbNull
        m_enmState = sckResolvingHost
        If AllocateMemory Then
            Dim lngAsynHandle As Long
            lngAsynHandle = WSAAsyncGetHostByName(m_lngWindowHandle, 32768, Host, ByVal m_lngMemoryPointer, 1024)
            If lngAsynHandle = 0 Then
                FreeMemory
                m_enmState = sckError
                Dim lngErrorCode As Long
                lngErrorCode = Err.LastDllError
                EventMsg lngErrorCode, "WinSock.ResolveIfHostname"
            End If
        Else
            m_enmState = sckError
        End If
    Else
        ResolveIfHostname = lngAddress
    End If
End Function

Private Function ResolveIfHostnameSync(ByVal Host As String, ByRef strHostIP As String, ByRef lngErrorCode As Long) As Long
    Dim lngPtrToHOSTENT As Long
    Dim udtHostent As HOSTENT
    Dim lngAddress As Long
    Dim lngPtrToIP As Long
    Dim arrIpAddress(1 To 4) As Byte
    Dim Count As Integer
    lngAddress = api_inet_addr(Host)
    If lngAddress = &HFFFF Then
        lngPtrToHOSTENT = api_gethostbyname(Host)
        If lngPtrToHOSTENT = 0 Then
            lngErrorCode = Err.LastDllError
            strHostIP = vbNullString
            ResolveIfHostnameSync = vbNull
        Else
            CopyMemory udtHostent, ByVal lngPtrToHOSTENT, LenB(udtHostent)
            CopyMemory lngPtrToIP, ByVal udtHostent.hAddrList, 4
            CopyMemory arrIpAddress(1), ByVal lngPtrToIP, 4
            CopyMemory lngAddress, ByVal lngPtrToIP, 4
            For Count = 1 To 4
                strHostIP = strHostIP & arrIpAddress(Count) & "."
            Next
            strHostIP = Left$(strHostIP, Len(strHostIP) - 1)
            lngErrorCode = 0
            ResolveIfHostnameSync = lngAddress
        End If
    Else
        lngErrorCode = 0
        strHostIP = Host
        ResolveIfHostnameSync = lngAddress
    End If
End Function

Private Function GetLocalPort(ByVal lngSocket As Long) As Long
    Dim udtSockAddr As sockaddr_in
    Dim lngResult As Long
    lngResult = api_getsockname(lngSocket, udtSockAddr, LenB(udtSockAddr))
    If lngResult = -1 Then
        GetLocalPort = -1
    Else
        GetLocalPort = IntegerToUnsigned(api_ntohs(udtSockAddr.sin_port))
    End If
End Function

Public Sub SendData(Data As Variant)
    Dim arrData() As Byte
    If m_enmProtocol = sckTCPProtocol Then
        If m_enmState <> sckConnected Then
            Exit Sub
        End If
    Else
        If Not SocketExists Then Exit Sub
        If Not BindInternal Then Exit Sub
        m_enmState = sckOpen
    End If
    Select Case varType(Data)
        Case vbString
            Dim strData As String
            strData = CStr(Data)
            If Len(strData) = 0 Then Exit Sub
            ReDim arrData(LenA(strData) - 1)
            arrData() = StrConv(strData, vbFromUnicode)
        Case vbArray + vbByte
            Dim strArray As String
            strArray = StrConv(Data, vbUnicode)
            If LenB(strArray) = 0 Then Exit Sub
            arrData() = StrConv(strArray, vbFromUnicode)
        Case vbBoolean
            Dim blnData As Boolean
            blnData = CBool(Data)
            ReDim arrData(LenB(blnData) - 1)
            CopyMemory arrData(0), blnData, LenB(blnData)
        Case vbByte
            Dim bytData As Byte
            bytData = CByte(Data)
            ReDim arrData(LenB(bytData) - 1)
            CopyMemory arrData(0), bytData, LenB(bytData)
        Case vbCurrency
            Dim curData As Currency
            curData = CCur(Data)
            ReDim arrData(LenB(curData) - 1)
            CopyMemory arrData(0), curData, LenB(curData)
        Case vbDate
            Dim datData As Date
            datData = CDate(Data)
            ReDim arrData(LenB(datData) - 1)
            CopyMemory arrData(0), datData, LenB(datData)
        Case vbDouble
            Dim dblData As Double
            dblData = CDbl(Data)
            ReDim arrData(LenB(dblData) - 1)
            CopyMemory arrData(0), dblData, LenB(dblData)
        Case vbInteger
            Dim intData As Integer
            intData = CInt(Data)
            ReDim arrData(LenB(intData) - 1)
            CopyMemory arrData(0), intData, LenB(intData)
        Case vbLong
            Dim lngData As Long
            lngData = CLng(Data)
            ReDim arrData(LenB(lngData) - 1)
            CopyMemory arrData(0), lngData, LenB(lngData)
        Case vbSingle
            Dim sngData As Single
            sngData = CSng(Data)
            ReDim arrData(LenB(sngData) - 1)
            CopyMemory arrData(0), sngData, LenB(sngData)
        Case Else
    End Select
    If Len(m_strSendBuffer) > 0 Then
        m_strSendBuffer = m_strSendBuffer + StrConv(arrData(), vbUnicode)
        Exit Sub
    Else
        m_strSendBuffer = m_strSendBuffer + StrConv(arrData(), vbUnicode)
    End If
    SendBufferedData
End Sub

Private Sub SendBufferedData()
    If m_enmProtocol = sckTCPProtocol Then
        SendBufferedDataTCP
    Else
        SendBufferedDataUDP
    End If
End Sub

Private Sub SendBufferedDataUDP()
    Dim lngAddress As Long
    Dim udtSockAddr As sockaddr_in
    Dim arrData() As Byte
    Dim lngBufferLength As Long
    Dim lngResult As Long
    Dim lngErrorCode As Long
    Dim strTemp As String
    lngAddress = ResolveIfHostnameSync(m_strRemoteHost, strTemp, lngErrorCode)
    If lngErrorCode <> 0 Then
        m_strSendBuffer = ""
    End If
    With udtSockAddr
        .sin_addr = lngAddress
        .sin_family = 2
        .sin_port = api_htons(UnsignedToInteger(m_lngRemotePort))
    End With
    lngBufferLength = LenA(m_strSendBuffer)
    arrData() = StrConv(m_strSendBuffer, vbFromUnicode)
    m_strSendBuffer = ""
    lngResult = api_sendto(m_lngSocketHandle, arrData(0), lngBufferLength, 0&, udtSockAddr, LenB(udtSockAddr))
    If lngResult = -1 Then
        lngErrorCode = Err.LastDllError
        If lngErrorCode <> 10035 Then
            m_enmState = sckError
            EventMsg lngErrorCode, "WinSock.SendBufferedDataUDP"
        End If
    End If
End Sub

Private Sub SendBufferedDataTCP()
    Dim arrData()       As Byte
    Dim lngBufferLength As Long
    Dim lngResult    As Long
    Dim lngTotalSent As Long
    Do Until lngResult = -1 Or Len(m_strSendBuffer) = 0
        lngBufferLength = Len(m_strSendBuffer)
        If lngBufferLength > m_lngSendBufferLen Then
            lngBufferLength = m_lngSendBufferLen
            arrData() = StrConv(Left$(m_strSendBuffer, m_lngSendBufferLen), vbFromUnicode)
        Else
            arrData() = StrConv(m_strSendBuffer, vbFromUnicode)
            lngBufferLength = UBound(arrData) + 1
        End If
        lngResult = api_send(m_lngSocketHandle, arrData(0), lngBufferLength, 0&)
        If lngResult = -1 Then
            Dim lngErrorCode As Long
            lngErrorCode = Err.LastDllError
            If lngErrorCode = 10035 Then
                If lngTotalSent > 0 Then RaiseEvent SendProgress(lngTotalSent, Len(m_strSendBuffer))
            Else
                m_enmState = sckError
                EventMsg lngErrorCode, "WinSock.SendBufferedData"
            End If
        Else
            lngTotalSent = lngTotalSent + lngResult
            If Len(m_strSendBuffer) > lngResult Then
                m_strSendBuffer = Mid$(m_strSendBuffer, lngResult + 1)
            Else
                m_strSendBuffer = ""
                Dim lngTemp As Long
                lngTemp = lngTotalSent
                lngTotalSent = 0
                RaiseEvent SendProgress(lngTemp, 0)
                RaiseEvent SendComplete
            End If
        End If
    Loop
End Sub

Private Function RecvDataToBuffer() As Long
    Dim arrBuffer() As Byte
    Dim lngBytesReceived As Long
    Dim strBuffTemporal As String
    ReDim arrBuffer(m_lngRecvBufferLen - 1)
    lngBytesReceived = api_recv(m_lngSocketHandle, arrBuffer(0), m_lngRecvBufferLen, 0&)
    If lngBytesReceived = -1 Then
        m_enmState = sckError
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
    ElseIf lngBytesReceived > 0 Then
        strBuffTemporal = StrConv(arrBuffer(), vbUnicode)
        m_strRecvBuffer = m_strRecvBuffer & Left$(strBuffTemporal, lngBytesReceived)
        RecvDataToBuffer = lngBytesReceived
    End If
End Function

Private Sub ProcessOptions()
    Dim lngResult As Long
    Dim lngBuffer As Long
    Dim lngErrorCode As Long
    If m_enmProtocol = sckTCPProtocol Then
        lngResult = api_getsockopt(m_lngSocketHandle, 65535, &H1002, lngBuffer, LenB(lngBuffer))
        If lngResult = -1 Then
            lngErrorCode = Err.LastDllError
        Else
            m_lngRecvBufferLen = lngBuffer
        End If
        lngResult = api_getsockopt(m_lngSocketHandle, 65535, &H1001, lngBuffer, LenB(lngBuffer))
        If lngResult = -1 Then
            lngErrorCode = Err.LastDllError
        Else
            m_lngSendBufferLen = lngBuffer
        End If
    Else
        lngBuffer = 1
        lngResult = api_setsockopt(m_lngSocketHandle, 65535, 32, lngBuffer, LenB(lngBuffer))
        lngResult = api_getsockopt(m_lngSocketHandle, 65535, 8195, lngBuffer, LenB(lngBuffer))
        If lngResult = -1 Then
            lngErrorCode = Err.LastDllError
        Else
            m_lngRecvBufferLen = lngBuffer
            m_lngSendBufferLen = lngBuffer
        End If
    End If
End Sub

Public Sub GetData(ByRef Data As Variant, Optional varType As Variant, Optional maxLen As Variant)
    If m_enmProtocol = sckTCPProtocol Then
        If m_enmState <> sckConnected And Not m_blnAcceptClass Then
            Exit Sub
        End If
    Else
        If m_enmState <> sckOpen Then
            Exit Sub
        End If
        If GetBufferLenUDP = 0 Then Exit Sub
    End If
    If Not IsMissing(maxLen) Then
        If Not IsNumeric(maxLen) Then
            If m_enmProtocol = sckTCPProtocol Then
                maxLen = LenA(m_strRecvBuffer)
            Else
                maxLen = GetBufferLenUDP
            End If
        End If
    End If
    Dim lngBytesRecibidos  As Long
    lngBytesRecibidos = RecvData(Data, False, varType, maxLen)
End Sub

Public Sub PeekData(ByRef Data As Variant, Optional varType As Variant, Optional maxLen As Variant)
    If m_enmProtocol = sckTCPProtocol Then
        If m_enmState <> sckConnected Then
            Exit Sub
        End If
    Else
        If m_enmState <> sckOpen Then
            Exit Sub
        End If
        If GetBufferLenUDP = 0 Then Exit Sub
    End If
    If Not IsMissing(maxLen) Then
        If IsNumeric(maxLen) Then
            If CLng(maxLen) < 0 Then
            End If
        Else
            If m_enmProtocol = sckTCPProtocol Then
                maxLen = LenA(m_strRecvBuffer)
            Else
                maxLen = GetBufferLenUDP
            End If
        End If
    End If
    Dim lngBytesRecibidos  As Long
    lngBytesRecibidos = RecvData(Data, True, varType, maxLen)
End Sub

Private Function RecvData(ByRef Data As Variant, ByVal blnPeek As Boolean, Optional varClass As Variant, Optional maxLen As Variant) As Long
    Dim blnMaxLenMiss   As Boolean
    Dim blnClassMiss    As Boolean
    Dim strRecvData     As String
    Dim lngBufferLen    As Long
    Dim arrBuffer()     As Byte
    Dim lngErrorCode    As Long
    If m_enmProtocol = sckTCPProtocol Then
        lngBufferLen = LenB(m_strRecvBuffer)
    Else
        lngBufferLen = GetBufferLenUDP
    End If
    blnMaxLenMiss = IsMissing(maxLen)
    blnClassMiss = IsMissing(varClass)
    If varType(Data) = vbEmpty Then
        If blnClassMiss Then varClass = vbArray + vbByte
    Else
        varClass = varType(Data)
    End If
    If varClass = vbString Or varClass = vbArray + vbByte Then
        If blnMaxLenMiss Then
            If lngBufferLen = 0 Then
                RecvData = 0
                arrBuffer = StrConv("", vbFromUnicode)
                Data = arrBuffer
                Exit Function
            Else
                RecvData = lngBufferLen
                BuildArray lngBufferLen, blnPeek, lngErrorCode, arrBuffer
            End If
        Else
            If maxLen = 0 Or lngBufferLen = 0 Then
                RecvData = 0
                arrBuffer = StrConv("", vbFromUnicode)
                Data = arrBuffer
                If m_enmProtocol = sckUDPProtocol Then
                    EmptyBuffer
                End If
                Exit Function
            ElseIf maxLen > lngBufferLen Then
                RecvData = lngBufferLen
                BuildArray lngBufferLen, blnPeek, lngErrorCode, arrBuffer
            Else
                RecvData = CLng(maxLen)
                BuildArray CLng(maxLen), blnPeek, lngErrorCode, arrBuffer
            End If
        End If
    End If
    Select Case varClass
        Case vbString
            Dim strData As String
            strData = StrConv(arrBuffer(), vbUnicode)
            Data = strData
        Case vbArray + vbByte
            Data = arrBuffer
        Case vbBoolean
            Dim blnData As Boolean
            If LenB(blnData) > lngBufferLen Then Exit Function
            BuildArray LenB(blnData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(blnData)
            CopyMemory blnData, arrBuffer(0), LenB(blnData)
            Data = blnData
        Case vbByte
            Dim bytData As Byte
            If LenB(bytData) > lngBufferLen Then Exit Function
            BuildArray LenB(bytData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(bytData)
            CopyMemory bytData, arrBuffer(0), LenB(bytData)
            Data = bytData
        Case vbCurrency
            Dim curData As Currency
            If LenB(curData) > lngBufferLen Then Exit Function
            BuildArray LenB(curData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(curData)
            CopyMemory curData, arrBuffer(0), LenB(curData)
            Data = curData
        Case vbDate
            Dim datData As Date
            If LenB(datData) > lngBufferLen Then Exit Function
            BuildArray LenB(datData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(datData)
            CopyMemory datData, arrBuffer(0), LenB(datData)
            Data = datData
        Case vbDouble
            Dim dblData As Double
            If LenB(dblData) > lngBufferLen Then Exit Function
            BuildArray LenB(dblData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(dblData)
            CopyMemory dblData, arrBuffer(0), LenB(dblData)
            Data = dblData
        Case vbInteger
            Dim intData As Integer
            If LenB(intData) > lngBufferLen Then Exit Function
            BuildArray LenB(intData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(intData)
            CopyMemory intData, arrBuffer(0), LenB(intData)
            Data = intData
        Case vbLong
            Dim lngData As Long
            If LenB(lngData) > lngBufferLen Then Exit Function
            BuildArray LenB(lngData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(lngData)
            CopyMemory lngData, arrBuffer(0), LenB(lngData)
            Data = lngData
        Case vbSingle
            Dim sngData As Single
            If LenB(sngData) > lngBufferLen Then Exit Function
            BuildArray LenB(sngData), blnPeek, lngErrorCode, arrBuffer
            RecvData = LenB(sngData)
            CopyMemory sngData, arrBuffer(0), LenB(sngData)
            Data = sngData
        Case Else
    End Select
End Function

Private Sub BuildArray(ByVal Size As Long, ByVal blnPeek As Boolean, ByRef lngErrorCode As Long, ByRef bytArray() As Byte)
    Dim strData As String
    If m_enmProtocol = sckTCPProtocol Then
        strData = Left$(m_strRecvBuffer, CLng(Size))
        bytArray = StrConv(strData, vbFromUnicode)
        If Not blnPeek Then
            m_strRecvBuffer = Mid$(m_strRecvBuffer, Size + 1)
        End If
    Else
        Dim arrBuffer() As Byte
        Dim lngResult As Long
        Dim udtSockAddr As sockaddr_in
        Dim lngFlags As Long
        If blnPeek Then lngFlags = 2&
        ReDim arrBuffer(Size - 1)
        lngResult = api_recvfrom(m_lngSocketHandle, arrBuffer(0), Size, lngFlags, udtSockAddr, LenB(udtSockAddr))
        If lngResult = -1 Then
            lngErrorCode = Err.LastDllError
        End If
        bytArray = arrBuffer
        GetRemoteInfoFromSI udtSockAddr, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
    End If
End Sub

Private Sub CleanResolutionSystem()
    Dim varAsynHandle As Variant
    Dim lngResult As Long
    lngResult = api_WSACancelAsyncRequest(varAsynHandle)
    If lngResult = 0 Then
        FreeMemory
    End If
End Sub

Public Sub Listen()
    If Not SocketExists Then Exit Sub
    If Not BindInternal Then Exit Sub
    Dim lngResult As Long
    lngResult = api_listen(m_lngSocketHandle, 5)
    If lngResult = -1 Then
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
    Else
        m_enmState = sckListening
    End If
End Sub

Public Sub Accept(requestID As Long)
    If m_enmState <> sckClosed Then
    End If
    m_lngSocketHandle = requestID
    m_enmProtocol = sckTCPProtocol
    ProcessOptions
    If Not IsAcceptRegistered(requestID) Then
        If IsSocketRegistered(requestID) Then
            m_lngSocketHandle = -1
            m_lngRecvBufferLen = 0
            m_lngSendBufferLen = 0
        Else
            m_blnAcceptClass = True
            m_enmState = sckConnected
            GetLocalInfo m_lngSocketHandle, m_lngLocalPortBind, m_strLocalIP
            RegisterSocket m_lngSocketHandle, False
            Exit Sub
        End If
    End If
    UnregisterAccept requestID
    GetLocalInfo m_lngSocketHandle, m_lngLocalPortBind, m_strLocalIP
    GetRemoteInfo m_lngSocketHandle, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
    m_enmState = sckConnected
End Sub

Private Sub UnregisterAccept(ByVal lngSocket As Long)
    m_colAcceptList.Remove "S" & lngSocket
    If m_colAcceptList.Count = 0 Then
        Set m_colAcceptList = Nothing
    End If
End Sub

Private Function GetLocalInfo(ByVal lngSocket As Long, ByRef lngLocalPort As Long, ByRef strLocalIP As String) As Boolean
    GetLocalInfo = False
    Dim lngResult As Long
    Dim udtSockAddr As sockaddr_in
    lngResult = api_getsockname(lngSocket, udtSockAddr, LenB(udtSockAddr))
    If lngResult = -1 Then
        lngLocalPort = 0
        strLocalIP = ""
    Else
        GetLocalInfo = True
        lngLocalPort = IntegerToUnsigned(api_ntohs(udtSockAddr.sin_port))
        strLocalIP = StringFromPointer(api_inet_ntoa(udtSockAddr.sin_addr))
    End If
End Function

Private Function GetRemoteInfo(ByVal lngSocket As Long, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String) As Boolean
    GetRemoteInfo = False
    Dim lngResult As Long
    Dim udtSockAddr As sockaddr_in
    lngResult = api_getpeername(lngSocket, udtSockAddr, LenB(udtSockAddr))
    If lngResult = 0 Then
        GetRemoteInfo = True
        GetRemoteInfoFromSI udtSockAddr, lngRemotePort, strRemoteHostIP, strRemoteHost
    Else
        lngRemotePort = 0
        strRemoteHostIP = ""
        strRemoteHost = ""
    End If
End Function

Private Sub GetRemoteInfoFromSI(ByRef udtSockAddr As sockaddr_in, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String)
    lngRemotePort = IntegerToUnsigned(api_ntohs(udtSockAddr.sin_port))
    strRemoteHostIP = StringFromPointer(api_inet_ntoa(udtSockAddr.sin_addr))
End Sub

Private Function GetBufferLenUDP() As Long
    Dim lngResult As Long
    Dim lngBuffer As Long
    lngResult = api_ioctlsocket(m_lngSocketHandle, &H4004667F, lngBuffer)
    If lngResult = -1 Then
        GetBufferLenUDP = 0
    Else
        GetBufferLenUDP = lngBuffer
    End If
End Function

Private Sub EmptyBuffer()
    Dim B As Byte
    api_recv m_lngSocketHandle, B, Len(B), 0&
End Sub

Private Function InitiateProcesses() As Long
    Dim udtWSAData As WSAData
    InitiateProcesses = 0
    m_lngSocksQuantity = m_lngSocksQuantity + 1
    If Not m_blnInitiated Then
        If m_lngWindowHandle = 0 Then
            m_lngWindowHandle = CreateWindowEx(0&, "STATIC", "SOCKET_WINDOW", 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
        End If
        m_blnInitiated = True
        Dim lngResult As Long
        lngResult = WSAStartup(&H101, udtWSAData)
    End If
End Function

Private Function FinalizeProcesses() As Long
    FinalizeProcesses = 0
    m_lngSocksQuantity = m_lngSocksQuantity - 1
    If m_blnInitiated And m_lngSocksQuantity = 0 Then
        m_blnInitiated = False
    End If
End Function

Private Sub EventMsg(lngErrCode As Long, Description As String)
    Dim blnCancelDisplay As Boolean
    blnCancelDisplay = True
    RaiseEvent Error(lngErrCode, GetErr(lngErrCode), 0, Description, "", 0, blnCancelDisplay)
    If blnCancelDisplay = False Then MsgBox GetErr(lngErrCode), vbOKOnly, Description
End Sub

Private Function GetErr(ByVal lngErrorCode As Long) As String
    Select Case lngErrorCode
        Case 10013
            GetErr = "权限被拒绝."
        Case 10048
            GetErr = "地址在使用中."
        Case 10049
            GetErr = "不能分配请求地址."
        Case 10047
            GetErr = "地址家族不支持的请求操作."
        Case 10037
            GetErr = "操作已经在进行."
        Case 10053
            GetErr = "软件造成连接中止."
        Case 10061
            GetErr = "连接被拒绝."
        Case 10054
            GetErr = "连接被对方重置."
        Case 10039
            GetErr = "需要目标地址."
        Case 10014
            GetErr = "错误地址."
        Case 10065
            GetErr = "没有到主机的路由."
        Case 10036
            GetErr = "操作现在在进行."
        Case 10004
            GetErr = "函数调用中断."
        Case 10022
            GetErr = "无效的参数."
        Case 10056
            GetErr = "套接字已经连接."
        Case 10024
            GetErr = "打开文件过多."
        Case 10040
            GetErr = "信息太长."
        Case 10050
            GetErr = "网络断开."
        Case 10052
            GetErr = "网络在重置时断开."
        Case 10051
            GetErr = "网络不可访问."
        Case 10055
            GetErr = "无缓存空间可用."
        Case 10042
            GetErr = "协议选项错误."
        Case 10057
            GetErr = "套接字没有连接."
        Case 10038
            GetErr = "无效套接字操作."
        Case 10045
            GetErr = "不支持的操作."
        Case 10046
            GetErr = "协议家族不支持."
        Case 10067
            GetErr = "进程过多."
        Case 10043
            GetErr = "协议不支持."
        Case 10041
            GetErr = "套接字协议类型错误."
        Case 10058
            GetErr = "套接字关闭后无法发送."
        Case 10044
            GetErr = "套接字类型不支持."
        Case 10060
            GetErr = "连接超时."
        Case 10035
            GetErr = "资源暂时不可用."
        Case 11001
            GetErr = "授权应答：未找到主机。"
        Case 10093
            GetErr = "套接字没有初始化."
        Case 11004
            GetErr = "无效名，对所请求的类型无数据记录 ."
        Case 11003
            GetErr = "不可恢复的错误."
        Case 10091
            GetErr = "网络子系统不可用."
        Case 11002
            GetErr = "非授权主机没有找到."
        Case 10092
            GetErr = "Winsock.dll 版本错误."
        Case 40006
            GetErr = "所请求的事务或请求本身的错误协议或者错误连接状态。"
        Case 40014
            GetErr = "传递给函数的参数格式不确定，或者不在指定范围内。"
        Case 40018
            GetErr = "不受支持的变量类型。"
        Case 40020
            GetErr = "在当前状态下的无效操作"
        Case Else
            GetErr = "未知的错误.错误号：" & CStr(lngErrorCode)
    End Select
End Function

Private Function HiWord(lngValue As Long) As Long
    If (lngValue And &H80000000) = &H80000000 Then
        HiWord = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        HiWord = (lngValue And &HFFFF0000) \ &H10000
    End If
End Function

Private Function StringFromPointer(ByVal lPointer As Long) As String
    Dim strTemp As String
    Dim lRetVal As Long
    strTemp = String$(LenA(ByVal lPointer), 0)
    lRetVal = api_lstrcpy(ByVal strTemp, ByVal lPointer)
    If lRetVal Then StringFromPointer = strTemp
End Function

Private Function UnsignedToInteger(Value As Long) As Integer
    If Value < 0 Or Value >= 65536 Then Error 6
    If Value <= 32769 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - 65536
    End If
End Function

Private Function IntegerToUnsigned(Value As Integer) As Long
    If Value < 0 Then
        IntegerToUnsigned = Value + 65536
    Else
        IntegerToUnsigned = Value
    End If
End Function

Private Function RegisterSocket(ByVal lngSocket As Long, ByVal blnEvents As Boolean) As Boolean
    Dim lngEvents As Long
    Dim lngResult As Long
    lngEvents = &H3B
    lngResult = WSAAsyncSelect(lngSocket, m_lngWindowHandle, 32769, lngEvents)
    If lngResult = -1 Then
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
    Else
        RegisterSocket = True
    End If
End Function

Private Function IsSocketRegistered(ByVal lngSocket As Long) As Boolean
    On Error GoTo Error_Handler
    m_colSocketsInst.Item ("S" & lngSocket)
    IsSocketRegistered = True
    Exit Function
Error_Handler:
    IsSocketRegistered = False
End Function

Private Sub RegisterAccept(ByVal lngSocket As Long)
    If m_colAcceptList Is Nothing Then
        Set m_colAcceptList = New Collection
    End If
    m_colAcceptList.Add "k" & lngSocket, "S" & lngSocket
End Sub

Private Function IsAcceptRegistered(ByVal lngSocket As Long) As Boolean
    On Error GoTo Error_Handler
    m_colAcceptList.Item ("S" & lngSocket)
    IsAcceptRegistered = True
    Exit Function
Error_Handler:
    IsAcceptRegistered = False
End Function

Private Sub UserControl_Initialize()
    m_lngSocketHandle = -1
    InitiateProcesses
End Sub

Private Sub UserControl_Terminate()
    CleanResolutionSystem
    If Not m_blnAcceptClass Then DestroySocket
    FinalizeProcesses
    Subclass_StopAll
    DestroyWindow m_lngWindowHandle
    m_lngWindowHandle = 0
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 420
    UserControl.Height = 420
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.LocalPort = PropBag.ReadProperty("LocalPort", 0)
    Me.Protocol = PropBag.ReadProperty("Protocol", 0)
    Me.RemoteHost = PropBag.ReadProperty("RemoteHost", "")
    Me.RemotePort = PropBag.ReadProperty("RemotePort", 0)
    Me.Tag = PropBag.ReadProperty("Tag", "")
    InitializeSubClassing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "LocalPort", Me.LocalPort, 0
    PropBag.WriteProperty "Protocol", Me.Protocol, 0
    PropBag.WriteProperty "RemoteHost", Me.RemoteHost, ""
    PropBag.WriteProperty "RemotePort", Me.RemotePort, 0
    PropBag.WriteProperty "Tag", Me.Tag, ""
End Sub

Private Sub InitializeSubClassing()
    On Error GoTo handle
    If Ambient.UserMode Then
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
        If Not bTrack Then Exit Sub
        Call Subclass_Start(m_lngWindowHandle)
        Call Subclass_AddMsg(m_lngWindowHandle, 32768, 1)
        Call Subclass_AddMsg(m_lngWindowHandle, 32769, 1)
    End If
handle:
End Sub

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call CopyMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    Dim i                       As Long
    Dim j                       As Long
    Dim nSubIdx                 As Long
    Dim sHex                    As String
    If aBuf(1) = 0 Then
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
               "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
               "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
               "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        i = 1
        Do While j < 200
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
            i = i + 2
        Loop
        If Subclass_InIDE Then
            aBuf(16) = &H90
            aBuf(17) = &H90
        End If
        ReDim sc_aSubData(0 To 0) As tSubData
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If
        Subclass_Start = nSubIdx
    End If
    With sc_aSubData(nSubIdx)
        .hWnd = lng_hWnd
        .nAddrSub = GlobalAlloc(0&, 200)
        .nAddrOrig = SetWindowLong(.hWnd, -4, .nAddrSub)
        Call CopyMemory(ByVal .nAddrSub, aBuf(1), 200)
        Call zPatchRel(.nAddrSub, 18, zAddrFunc("vba6", "EbMode"))
        Call zPatchVal(.nAddrSub, 68, .nAddrOrig)
        Call zPatchRel(.nAddrSub, 78, zAddrFunc("user32", "SetWindowLongA"))
        Call zPatchVal(.nAddrSub, 116, .nAddrOrig)
        Call zPatchRel(.nAddrSub, 121, zAddrFunc("user32", "CallWindowProcA"))
        Call zPatchVal(.nAddrSub, 186, ObjPtr(Me))
    End With
End Function

Private Sub Subclass_StopAll()
    Dim i As Long
    On Error GoTo er
    i = UBound(sc_aSubData())
    Do While i >= 0
        With sc_aSubData(i)
            If .hWnd <> 0 Then
                Call Subclass_Stop(.hWnd)
            End If
        End With
        i = i - 1
    Loop
er:
End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLong(.hWnd, -4, .nAddrOrig)
        Call zPatchVal(.nAddrSub, 93, 0)
        Call zPatchVal(.nAddrSub, 137, 0)
        Call GlobalFree(.nAddrSub)
        .hWnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB
        Erase .aMsgTblA
    End With
End Sub

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal eMsgWhen As Long, ByVal nAddr As Long)
    Dim nEntry  As Long
    Dim nOff1   As Long
    Dim nOff2   As Long
    If uMsg = -1 Then
        nMsgCnt = -1
    Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = 0 Then
                aMsgTbl(nEntry) = uMsg
                Exit Sub
            ElseIf aMsgTbl(nEntry) = uMsg Then
                Exit Sub
            End If
        Loop
        nMsgCnt = nMsgCnt + 1
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        aMsgTbl(nMsgCnt) = uMsg
    End If
    If eMsgWhen = 2 Then
        nOff1 = 88
        nOff2 = 93
    Else
        nOff1 = 132
        nOff2 = 137
    End If
    If uMsg <> -1 Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
End Function

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0
        With sc_aSubData(zIdx)
            If .hWnd = lng_hWnd Then
                If Not bAdd Then
                    Exit Function
                End If
            ElseIf .hWnd = 0 Then
                If bAdd Then
                    Exit Function
                End If
            End If
        End With
        zIdx = zIdx - 1
    Loop
    If Not bAdd Then
    End If
End Function

Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim hMod        As Long
    Dim bLibLoaded  As Boolean
    hMod = GetModuleHandle(sModule)
    If hMod = 0 Then
        hMod = LoadLibrary(sModule)
        If hMod Then
            bLibLoaded = True
        End If
    End If
    If hMod Then
        If GetProcAddress(hMod, sFunction) Then
            IsFunctionExported = True
        End If
    End If
    If bLibLoaded Then
        Call FreeLibrary(hMod)
    End If
End Function

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal eMsgWhen As Long)
    With sc_aSubData(zIdx(lng_hWnd))
        If eMsgWhen And 2 Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, 2, .nAddrSub)
        End If
        If eMsgWhen And 1 Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, 1, .nAddrSub)
        End If
    End With
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call CopyMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function
'Download by http://www.codefans.net
Private Function zSetTrue(T As Boolean) As Boolean
    zSetTrue = True
    T = True
End Function




