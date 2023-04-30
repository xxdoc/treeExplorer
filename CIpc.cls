VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event Message(m As String, ByRef retval As Long)

Private f As Form

Private WithEvents sc As CSubclass2
Attribute sc.VB_VarHelpID = -1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageByVal Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 
Private Type COPYDATASTRUCT
    dwFlag As Long
    cbSize As Long
    lpData As Long
End Type

Private Const WM_COPYDATA = &H4A

Public RemoteHWND As Long
Private subclassed_hwnd As Long
Public serverName As String
         
Public timeout As Long
Private DataIn As String

Private isListening As Boolean


Property Get IamListening() As Boolean
    IamListening = isListening
End Property

Function Listen(mf As Object, server As String) As Boolean

    If isListening Then Exit Function
     
    Set f = mf
    subclassed_hwnd = mf.hWnd
    
    If Not sc.AttachMessage(subclassed_hwnd, WM_COPYDATA) Then Exit Function
    
    RegisterServer server
    isListening = True
    Listen = True
     
End Function

 
Sub RegisterServer(server As String)
    On Error Resume Next
    SaveSetting "IPC", "Handles", server, f.hWnd
End Sub


'
' Private Function ValidateActiveIDAWindows()
'     On Error Resume Next
'     Dim X
'     For Each X In Servers 'remove any that arent still valid..
'        If IsWindow(X) = 0 Then
'            Servers.Remove "hwnd:" & X
'        End If
'     Next
' End Function
 
Function SendCmdRecvText(msg As String) As String
    Send msg
    SendCmdRecvText = DataIn
End Function
 
Function SendCmdRecvInt(msg As String) As Long
    SendCmdRecvInt = Send(msg)
End Function

Function Send(msg As String) As Long
    On Error Resume Next
    
    DataIn = Empty
    Dim buffer(1 To 2048) As Byte
    Dim CopyData As COPYDATASTRUCT
    
    If IsWindow(RemoteHWND) = 0 Then
        If Not FindClient(serverName) Then
            f.List1.AddItem "SendData Invalid hwnd: " & RemoteHWND
            Exit Function
        End If
    End If
    
    f.List1.AddItem "SendData: " & msg
    
    CopyMemory buffer(1), ByVal msg, Len(msg)
    CopyData.dwFlag = 3
    CopyData.cbSize = Len(msg) + 1
    CopyData.lpData = VarPtr(buffer(1))
    Send = SendMessage(RemoteHWND, WM_COPYDATA, RemoteHWND, CopyData)
    
End Function

'finds a client by reg name lookup, sets it as remoteHWND if found
Function FindClient(client As String) As Boolean
    Dim hWnd As Long
    On Error Resume Next
    hWnd = CLng(GetSetting("IPC", "Handles", client, 0))
    If hWnd <> 0 Then
        If IsWindow(hWnd) = 1 Then
            FindClient = True
            RemoteHWND = hWnd
            serverName = client
        Else
            SaveSetting "IPC", "Handles", client, 0
            RemoteHWND = 0
            FindClient = False
        End If
    End If
End Function

Private Sub Class_Initialize()
    Set sc = New CSubclass2
End Sub

Function ShutDown()
    isListening = False
    sc.DetatchMessage subclassed_hwnd, WM_COPYDATA
End Function



Private Sub sc_MessageReceived(hWnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
      
      Dim msg As String, retval As Long
    
      If wMsg = WM_COPYDATA Then
            If RecieveTextMessage(lParam, msg) Then
                RaiseEvent Message(msg, retval)
                sc.OverRideRetVal retval
            End If
      End If
            
End Sub

Private Function RecieveTextMessage(lParam As Long, ByRef msg As String) As Boolean
   
    Dim CopyData As COPYDATASTRUCT
    Dim buffer(1 To 2048) As Byte
    Dim temp As String
    
    msg = Empty
    
    CopyMemory CopyData, ByVal lParam, Len(CopyData)
    
    If CopyData.dwFlag = 3 Then
        CopyMemory buffer(1), ByVal CopyData.lpData, CopyData.cbSize
        temp = StrConv(buffer, vbUnicode)
        temp = Left$(temp, InStr(1, temp, Chr$(0)) - 1)
        'heres where we work with the intercepted message
        msg = temp
        RecieveTextMessage = True
        f.List1.AddItem "RecvData: " & msg
    End If
    
End Function

 

