VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   15615
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   28560
   LinkTopic       =   "Form1"
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrQue 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8340
      Top             =   120
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   63900
      TabIndex        =   0
      Top             =   810
      Width           =   2.45745e5
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim que As New Collection

'this is kind of bullshit but it is what it is...
Function queCommand(cmd)
    tmrQue.Enabled = False
    que.Add cmd
    tmrQue.Enabled = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not unloading Then
        Cancel = True
        Me.Visible = False
    End If
End Sub

Private Sub tmrQue_Timer()
    Dim i As Long
    tmrQue.Enabled = False
    For i = 1 To que.Count
        handleCmd que(i)
    Next
    Set que = New Collection
End Sub

Function handleCmd(m)

    On Error Resume Next
    Dim c As VBComponent, i As Long, tmp As String, p As VBProject, fn As String, j As Long
    Dim x() As String
    
    If Left(m, 8) = "addfile:" Then
        m = Mid(m, 9)
        If FileExists(CStr(m)) Then
            VBInstance.ActiveVBProject.VBComponents.AddFile CStr(m)
        End If
    End If
    
    If Left(m, 5) = "show:" Then
        m = LCase(Mid(m, 6))
        For Each c In VBInstance.ActiveVBProject.VBComponents
            If LCase(c.Name) = m Then
                c.CodeModule.CodePane.Show
                Exit For
            End If
        Next
        Exit Function
    End If
    
    If m = "list" Then
        For Each c In VBInstance.ActiveVBProject.VBComponents
            'name may be different than file name, file name may not exist if not yet saved...
            tmp = c.Type & "|" & c.Name & "|" & c.FileNames(1)
            push x, tmp
        Next
        tmp = Join(x, vbCrLf)
        memfile.WriteFile tmp, , True 'maybe > than our 2048 ipc send buffer...
        ipc.Send Len(tmp)
        Exit Function
    End If
    
    If m = "projects" Then
        tmp = VBInstance.VBProjects.Count & "|"
        For Each p In VBInstance.VBProjects
            tmp = tmp & p.Name & "|"
        Next
        tmp = Mid(tmp, 1, Len(tmp) - 1)
        ipc.Send tmp
        Exit Function
    End If
    
    If Err.Number <> 0 Then
        List1.AddItem "Error in handleCmd: " & Err.Description
    End If
    
End Function


'VBComponents.Remove c does not work...
'we could select the component..and then find the Project->remove xxx menu item and click it?

'    If Left(m, 8) = "remfile:" Then
'        tmp = Mid(m, 9)
'        For i = 1 To VBInstance.ActiveVBProject.VBComponents.Count
'            fn = LCase(VBInstance.ActiveVBProject.VBComponents(i).FileNames(1))
'            If Len(fn) > 0 Then
'                If tmp = fn Then
'                    VBInstance.ActiveVBProject.VBComponents.Remove VBInstance.ActiveVBProject.VBComponents(i)
'                    DoEvents
'                    Exit Function
'                End If
'            End If
'        Next
'
'
''       For Each c In VBInstance.ActiveVBProject.VBComponents
''           'name may be different than file name, file name may not exist if not yet saved...
''           For i = 1 To c.FileCount
''               If tmp = LCase(c.FileNames(i)) Then
''
''                    VBInstance.ActiveVBProject.VBComponents.Remove c
''
''                    Exit Function
''               End If
''           Next
''       Next
'       Exit Function
'    End If


