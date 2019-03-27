VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SocketStreamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'SocketStreamer for Winsock control by Icelolly
'15 May 2018    Project created
'17 May 2018    Finished
'18 May 2018    Fixed a serious bug in Socket_DataArrival(). An error occurs when multiple packet is received.

'Notes:
'1. You are supposed to handle connection events such as ConnectionRequest for the Winsock control.
'2. Do NOT handle DataArrival event for the Winsock control or the program may not run properly.        (IMPORTANT)
'3. Avoid sending data via Winsock control directly since it may 'confuse' the buffer.
'   Use SendStreamData function of this class instead.

'How to declare:
'Private WithEvents ss As SocketStreamer    'Put this in the General region
'Set ss = New SocketStreamer
'Set ss.Socket = Me.wsMain                  'wsMain is a Winsock control

Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Event PacketArrival(Data() As Byte)                  'Packet arrival event

Dim Buffer()                    As Byte                     'Data buffer

'Clear the buffer
Public Sub ClearBuffer()
    ReDim Buffer(0)
End Sub

Private Sub Class_Initialize()
    ClearBuffer
End Sub

Public Sub SendStreamData(Data() As Byte)
    Dim SendBuf()   As Byte
    Dim DataSize    As Long
    
    'Buffer graph:
    'Average value of first 4 bytes
    '        ¡ý
    '¡õ¡õ¡õ¡õ¡õ¡õ¡­
    '©¸©¤©¤©¼  ©¸©¤
    'Data Size  Data
    
    DataSize = UBound(Data)
    If DataSize = -1 Then                                                           'Avoid sending empty data
        Exit Sub
    End If
    ReDim SendBuf(DataSize + 5)                                                     'Allocate temp buffer
    CopyMemory SendBuf(0), DataSize, 4                                              'Copy the data size (Long type) to the head of the buffer
    SendBuf(4) = (CInt(SendBuf(0)) + SendBuf(1) + SendBuf(2) + SendBuf(3)) / 4      'Calculate the average value of first 4 byte
    CopyMemory SendBuf(5), Data(0), DataSize + 1                                    'Copy the data to the buffer
    Socket.SendData SendBuf                                                         'Send the buffer
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim tempData()      As Byte                                                                 'Data received
    Dim tempBuffer()    As Byte                                                                 'Temp buffer to store the main buffer data when it's reallocated
    Dim DataSize        As Long                                                                 'Data size of the packet in the buffer
    Dim PrevBufSz       As Long                                                                 'Writing position of the buffer
    
    Me.Socket.GetData tempData                                                                  'Receive data
    bytesTotal = UBound(tempData) + 1
    PrevBufSz = UBound(Buffer)                                                                  'Record the data writing position of buffer
    If PrevBufSz = 0 Then                                                                       'If it's the first packet in the buffer
        ReDim Preserve Buffer(PrevBufSz + bytesTotal - 1)                                           'Allocate buffer memory
        CopyMemory Buffer(PrevBufSz), tempData(0), bytesTotal                                       'Copy data to buffer
    Else
        ReDim Preserve Buffer(PrevBufSz + bytesTotal)                                               'Allocate buffer memory
        CopyMemory Buffer(PrevBufSz + 1), tempData(0), bytesTotal                                   'Copy data to buffer
    End If
    
    If UBound(Buffer) > 5 Then                                                                      'Check the size of the packet
        If Buffer(4) <> CByte((CInt(Buffer(0)) + Buffer(1) + Buffer(2) + Buffer(3)) / 4) Then           'Check buffer header
            Err.Raise 60001, , "Invalid packet header, ClearBuffer() suggested"                             'Packet check failure
        End If
    Else                                                                                        'Abnormal packet size
        Err.Raise 60002, , "Buffer size is less than 5 bytes, ClearBuffer() suggested"
    End If
    
    CopyMemory DataSize, Buffer(0), 4                                                           'Get the size of the first packet in the buffer
    Do While UBound(Buffer) >= DataSize + 5                                                     'The first packet in the buffer is received completely
        ReDim tempData(DataSize)                                                                    'Allocate the buffer to store the single packet
        CopyMemory tempData(0), Buffer(5), DataSize + 1                                             'Read the packet data from the buffer
        RaiseEvent PacketArrival(tempData)                                                          'Raise the packet arrival event
        If UBound(Buffer) - DataSize - 6 > -1 Then                                                  'If it's not the last packet in the buffer
            ReDim tempBuffer(UBound(Buffer) - DataSize - 6)                                             'Allocate the temp buffer
            CopyMemory tempBuffer(0), Buffer(6 + DataSize), UBound(tempBuffer) + 1                      'Read the unhandled data from the buffer
            ReDim Buffer(UBound(tempBuffer))                                                            'Reallocate the buffer
            CopyMemory Buffer(0), tempBuffer(0), UBound(tempBuffer) + 1                                 'Copy the unhandled data back to the buffer
            CopyMemory DataSize, Buffer(0), 4                                                           'Get the size of the first packet in the buffer
        Else                                                                                        'If it's the last packet in the buffer
            ReDim Buffer(0)                                                                             'Clear buffer
            Exit Do                                                                                     'Quit the loop
        End If
    Loop
End Sub

