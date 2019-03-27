Attribute VB_Name = "modThreadSupport"
'====================================================
'Description:   Implements multithreading in VB6
'Source:        https://tieba.baidu.com/p/3616346086?pid=65162934490
'               I pay the highest respect to the author and all people who
'               made effort to this code! Thank you for your hard work!
'Author:        Combined by IceLolly
'Note:          Code in this module is not written by me, I just combined
'               the code I found in Baidu Tieba and made some minor modifications
'File:          modThreadSupport.bas
'====================================================

Option Explicit

'Create threads
Public Declare Function CreateThread Lib "kernel32" (ByVal lpSecurityAttributes As Long, ByVal dwStackSize As Long, _
    ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long

'Get module base address
Public Declare Function VBGetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModName As Long) As Long

Public AvoidReentrant   As Boolean                  'Prevents the main thread from creating for the second time
Public IsIDEorEXE       As Long                     'In IDE or not
Public VBHeader         As Long                     'VB6 header address

'Description:   To retrieve VB6 data header address
'Return:        Succeed or not
Public Function GETVBHeader() As Boolean
    Dim BaseAds     As Long, GetOffSet          As Long
    Dim VBHdChar(3) As Byte, MemData(&H1FDA&)   As Byte
    
    IsIDEorEXE = App.LogMode
    
    If IsIDEorEXE <> 0 Then
        VBHdChar(0) = &H56          'V
        VBHdChar(1) = &H42          'B
        VBHdChar(2) = &H35          '5
        VBHdChar(3) = &H21          '!
        BaseAds = VBGetModuleHandle(ByVal 0&) + &H1024&
        CopyMemory MemData(0), ByVal (BaseAds), &H1FDB&
        GetOffSet = InStrB(1, MemData, VBHdChar, vbBinaryCompare)
        If GetOffSet > 0 Then
            VBHeader = GetOffSet + BaseAds - 1
            GETVBHeader = True
        Else
            MessageBoxW ByVal 0&, "Failed to locate VB data header! Program may have instabilities.", _
                "Unable to Find VB Header", MB_ICONEXCLAMATION
            VBHeader = 0
            GETVBHeader = False
        End If
    Else
        VBHeader = 0
        GETVBHeader = False
    End If
End Function

'Description:   To initialize other components for VB6 runtime library
'Return:        Succeed or not
Public Function InitVBdll() As Boolean
    Dim pIID    As IID, pDummy     As Long
    Dim u1      As Long, u2         As Long, u3     As Long
    
    If VBHeader > 0 Then
        'Set pIID = IID of IClassFactory = {00000001-0000-0000-C000-000000000046}
        pIID.Data1 = &H1&
        pIID.Data4(0) = &HC0
        pIID.Data4(7) = &H46
        
        'Get u1, u2 for VBDllGetClassObject
        u3 = VBGetModuleHandle(ByVal 0&)
        UserDllMain u1, u2, u3, 1&, 0&
        VBDllGetClassObject u1, u2, VBHeader, pDummy, pIID, pDummy
        InitVBdll = True
    Else
        InitVBdll = False
    End If
End Function

'Description:   The entry point of the program
'Note:          We need to initalize VB6 threading environment here
Sub Main()
    If AvoidReentrant = False Then                      'Prevent the main thread from createing for the second time
        AvoidReentrant = True
        If App.PrevInstance Then                            'Allow one instance only
            MessageBoxW ByVal 0&, "There's already one running instance.", "Only One Instance Allowed", MB_ICONERROR
            Exit Sub
        Else
            InitCommonControls                                  'Initalizes common controls
            GETVBHeader                                         'Get VB6 data header address
            
            frmMain.Show                                        'Show the main window
        End If
    End If
End Sub

'The following snippet is a template of thread procedure
'All thread functions in this program will follow this template
'
'Public Function ThreadProc(Param As Long) As Long
'    CreateIExprSrvObj 0&, 4&, 0&                                                        'Initialize VB6 runtime library
'    CoInitializeEx ByVal 0&, ByVal (COINIT_MULTITHREADED Or COINIT_SPEED_OVER_MEMORY)   'Initialize COM components
'    InitVBdll                                                                           'Initialize VB6 multithreading environment
'    '==============================
'
'    'Do things
'
'    '==============================
'    CoUninitialize                                                                      'Unitialize COM components
'End Function

