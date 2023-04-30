VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVBComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Typ  As Long
Public name As String
Public path As String
Public raw As String

Public icon As String
Public defFolder As String
Public n As Node

Function loadFromList(raw)
    Dim yy() As String, icn As String
    
    Me.raw = raw
    yy = Split(raw, "|") 'type|name|path  path may not exist yet if not saved...
    Typ = yy(0)
    name = yy(1)
    path = yy(2)
    
    defFolder = DefaultFolderForType(Typ, icn)
    icon = icn
        
End Function

Function loadFromEvent(e As CComponentEvent)
    Dim icn As String
    
    Typ = e.ComponentType
    name = e.ComponentName
    path = e.FileName
    
    defFolder = DefaultFolderForType(Typ, icn)
    icon = icn
    
End Function
