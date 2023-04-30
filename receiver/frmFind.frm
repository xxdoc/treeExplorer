VERSION 5.00
Begin VB.Form frmFind 
   Caption         =   "Find"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5970
   LinkTopic       =   "Form2"
   ScaleHeight     =   6300
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin TreeExplorer.ucFilterList lv 
      Height          =   5835
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10292
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim fso As New CFileSystem2
Dim m_tv As TreeView

Sub init(tv As TreeView)

    On Error Resume Next
    Dim n As Node, fn As String, li As ListItem
    
    lv.Clear
    
    For Each n In tv.Nodes
        If n.Image <> "folder" And Len(n.key) > 0 Then
            fn = fso.FileNameFromPath(n.key)
            If Len(fn) > 0 Then
                Set li = lv.AddItem(fn)
                Set li.tag = n
            End If
        End If
    Next
    
    Me.Visible = True
    
End Sub

Private Sub Form_Load()
    lv.SetColumnHeaders "file*"
    lv.SetFont "Courier", 12
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lv.Width = Me.ScaleWidth - 200
    lv.Height = Me.ScaleHeight - 200
End Sub

Private Sub lv_ItemClick(ByVal item As MSComctlLib.ListItem)
    On Error Resume Next
    Dim n As Node
    Set n = item.tag
    n.EnsureVisible
    n.selected = True
End Sub
