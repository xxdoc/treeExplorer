VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClearList 
      Caption         =   "Clear"
      Height          =   435
      Left            =   6960
      TabIndex        =   10
      Top             =   6600
      Width           =   1155
   End
   Begin VB.CommandButton cmdSaveTree 
      Caption         =   "Save Tree"
      Height          =   315
      Left            =   9960
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdStartup 
      Caption         =   "Startup"
      Height          =   435
      Left            =   9900
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdFreshMirror 
      Caption         =   "Mirror Fresh"
      Height          =   315
      Left            =   9900
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "list"
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reconnect"
      Height          =   495
      Left            =   10560
      TabIndex        =   5
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   2655
      Left            =   4320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   3060
      Width           =   7455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   5040
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   4260
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1800
      Width           =   1995
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   180
      TabIndex        =   0
      Top             =   6420
      Width           =   6315
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   5340
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":059A
            Key             =   "quest"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08EE
            Key             =   "frm"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14B2
            Key             =   "unk"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1804
            Key             =   "mdi"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B58
            Key             =   "bas"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":20F2
            Key             =   "cls"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":268C
            Key             =   "pag"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C26
            Key             =   "ctl"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":31C0
            Key             =   "txt"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":331A
            Key             =   "func"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":366C
            Key             =   "dob"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3C06
            Key             =   "dsr"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":41A0
            Key             =   "proj"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   6135
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   10821
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "img1"
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
         Begin VB.Menu mnuAddGroup 
            Caption         =   "Group"
         End
         Begin VB.Menu mnuAddFolder 
            Caption         =   "Folder"
         End
         Begin VB.Menu mnuAddFile 
            Caption         =   "Files"
         End
      End
      Begin VB.Menu mnuRemoveItem 
         Caption         =   "Remove Item"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ipc As CIpc
Attribute ipc.VB_VarHelpID = -1
Private WithEvents saveTree As CSaveTree
Attribute saveTree.VB_VarHelpID = -1

'template for dragging nodes around: https://www.developerfusion.com/article/77/treeview-control/8/


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private blnDragging As Boolean
Private selNode As Node
Dim memfile As New CMemMapFile
Dim projPath As String

Private Sub cmdClearList_Click()
    List1.Clear
End Sub

Private Sub cmdSaveTree_Click()
    Dim cfg As String
    projPath = ipc.SendCmdRecvText("projpath")
    If Not fso.FileExists(projPath) Then Exit Sub 'not connected?
    cfg = fso.GetParentFolder(projPath) & "\tree.cfg"
    saveTree.saveTree tv, cfg
End Sub

Private Sub saveTree_DeSerialize(n As MSComctlLib.Node, ByVal appendTag As String, ByVal index As Long)
    'On Error Resume Next
    Dim c As New CVBComponent
    
    If InStr(appendTag, "|") = 0 Then
        If InStr(appendTag, "- expanded") > 0 Then
            n.Expanded = True
            appendTag = Replace(appendTag, "- expanded", Empty)
        End If
        n.key = n.Text
        n.Image = appendTag
    Else
        c.loadFromList appendTag
        n.key = c.name
        n.Image = c.icon
        Set c.n = n
        Set n.tag = c
    End If
    
End Sub

Private Sub saveTree_Serialize(n As MSComctlLib.Node, appendTag As String, ByVal index As Long)
    On Error Resume Next
    Dim c As CVBComponent
    
    If n.tag Is Nothing Then 'its a top level entry just save icon
        appendTag = n.Image
        If n.Expanded Then appendTag = appendTag & "- expanded"
    Else
        Set c = n.tag
        appendTag = c.raw
    End If
    
End Sub

 
Private Sub cmdStartup_Click()
    
    Dim cfg As String, vbc As CVBComponent, x, xx() As String, n As Node, nn As Node, foundNode As Node
    Dim freshElements As New Collection
    
    projPath = ipc.SendCmdRecvText("projpath")
    If Not fso.FileExists(projPath) Then Exit Sub 'not connected?
    
    cfg = fso.GetParentFolder(projPath) & "\tree.cfg"
    
    'easy we just build new nothing saved...
    If Not fso.FileExists(cfg) Then
        cmdFreshMirror_Click
        Exit Sub
    End If
    
    saveTree.RestoreTree tv, cfg
    'now we need to diff and see if were missing anything (added or lost)
    
    x = listRemote()
    If Len(x) = 0 Then Exit Sub
    
    xx = Split(x, vbCrLf)
    For Each x In xx
    
        Set vbc = New CVBComponent
        vbc.loadFromList x
        freshElements.Add vbc, vbc.name
        
        If NodeExists(vbc.name, foundNode) Then
            'great this didnt change
        Else
        
            List1.AddItem vbc.name & " - a new entry has been added in the IDE we dont know about"
            
            If Not NodeExists(vbc.defFolder, n) Then 'use default folder name since we dont know where to place it
                Set n = tv.Nodes.Add(tv.Nodes(1), tvwChild, vbc.defFolder, vbc.defFolder, "folder")
            End If

            Set nn = tv.Nodes.Add(n, tvwChild, vbc.name, vbc.name, vbc.icon)
            Set nn.tag = vbc
            Set vbc.n = nn
            
            n.Expanded = True
        End If
    Next
    
    'now we need to look for nodes we had in our tree, but which are no longer in the IDE
    For Each n In tv.Nodes
        If n.Image <> "folder" And n.Image <> "proj" Then
            If Not keyExistsInCollection(n.key, freshElements) Then
                n.Image = "quest"
            End If
        End If
    Next

End Sub

Private Sub Command1_Click()
   Text2 = ipc.SendCmdRecvText(Text1)
End Sub

Private Sub Command2_Click()
    Reconnect
End Sub

Function listRemote() As String

    On Error Resume Next
    Dim x As String, sz As Long, i As Long
    
    ipc.Send "list"
    
    For i = 0 To 100
        Sleep 10
        DoEvents
        If Len(ipc.LastRecv) > 0 Then Exit For
    Next
    
    sz = CLng(ipc.LastRecv)
    If sz < 1 Then Exit Function
    
    If Not memfile.ReadLength(x, sz) Then Exit Function
    
    listRemote = x
     
    
End Function

Private Sub Command3_Click()
    listRemote
End Sub


Private Sub cmdFreshMirror_Click()
    'Dim hwnd As Long
    'we cant sync this way no file paths..
    'hwnd = ipc.SendCmdRecvInt("treehwnd")
    'CopyRemoteTreeView hwnd, tv
    'List1.AddItem "treehwnd: " & hwnd
    
    On Error Resume Next
    
    Dim x, xx() As String, p As Node, n As Node, nn As Node
    Dim vbc As CVBComponent
    
    projPath = ipc.SendCmdRecvText("projpath")
    
    tv.Nodes.Clear

    Set p = tv.Nodes.Add(, , projPath, fso.FileNameFromPath(projPath), "proj")
    p.Expanded = True
    
    x = listRemote()
    If Len(x) = 0 Then Exit Sub
    
    xx = Split(x, vbCrLf)
    For Each x In xx
        Set vbc = New CVBComponent
        vbc.loadFromList x
        
        If NodeExists(vbc.name) Then 'this shouldnt happen..its from the ide and clean build
            List1.AddItem "Node exists name: " & vbc.name
        Else
            If Not NodeExists(vbc.defFolder, n) Then
                Set n = tv.Nodes.Add(tv.Nodes(1), tvwChild, vbc.defFolder, vbc.defFolder, "folder")
            End If

            Set nn = tv.Nodes.Add(n, tvwChild, vbc.name, vbc.name, vbc.icon)
            Set nn.tag = vbc
            Set vbc.n = nn
            
            n.Expanded = True
        End If
    Next
    
End Sub


 

Private Sub Form_Load()

    Set saveTree = New CSaveTree
    Set ipc = New CIpc
    ipc.Listen Me, "Treeview"
    Reconnect
    
    If Not memfile.CreateMemMapFile("ProjectExplorer", 200000, True) Then
        List1.AddItem "CreateMemMapFile error:" & memfile.ErrorMessage
        If Not memfile.OpenMemMapFile("ProjectExplorer", 200000) Then 'may already exist shared all instances..
            List1.AddItem "OpenMemMapFile error:" & memfile.ErrorMessage
        End If
    End If
    
End Sub

Sub Reconnect()

    If Not ipc.FindClient("ProjectExplorer") Then
        List1.AddItem "Could not find ProjectExplorer"
    Else
        ipc.Send "hwnd:" & Me.hwnd
    End If

End Sub

Private Sub ipc_Message(m As String)
    
    Dim ce As CComponentEvent
    Dim c As CVBComponent
    
    If Left(m, 10) = "Component|" Then
        Set ce = LoadComponentEvent(m)
        Set c = HandleComponentEvent(ce) 'this will handle renames in tree, and add new components
    End If

End Sub

Function LoadComponentEvent(raw As String) As CComponentEvent
    Dim e As New CComponentEvent
    e.init raw
    Set LoadComponentEvent = e
End Function

Private Sub mnuAddFolder_Click()
        
        On Error Resume Next
        Dim f As String, p As Node, fn As String
        
        If tv.Nodes.Count = 0 Then Exit Sub
        If selNode Is Nothing Then selNode = tv.Nodes(1)
        
        f = dlg.FolderDialog2()
        If Len(f) = 0 Then Exit Sub
        
        fn = fso.FolderName(f)
        Set p = tv.Nodes.Add(selNode, tvwChild, f, fn, "folder")
        AddFolderToTree f, p
            
End Sub

Private Sub mnuAddFile_Click()

        On Error Resume Next
        Dim x, p As Node, fn As String, c As Collection
        
        If tv.Nodes.Count = 0 Then Exit Sub
        If selNode Is Nothing Then selNode = tv.Nodes(1)
        
        Set c = dlg.OpenMulti()
        If c.Count = 0 Then Exit Sub

        For Each x In c
            AddNodeFromFile selNode, x
        Next
        
End Sub

Function AddFolderToTree(ByVal folder As String, p As Node, Optional recursive As Boolean = True)
    
        Dim ff() As String, x, bn As String, pp As Node
        
        If Not fso.FolderExists(folder) Then Exit Function
        
        'fn = fso.FolderName(f)
        ff = fso.GetFolderFiles(folder)
        p.Expanded = True
        
        For Each x In ff
            AddNodeFromFile p, x
        Next
        
        If recursive Then
            ff = fso.GetSubFolders(folder)
            If Not AryIsEmpty(ff) Then
                For Each x In ff
                    bn = fso.FolderName(CStr(x))
                    Set pp = tv.Nodes.Add(p, tvwChild, x, bn, "folder")
                    pp.Expanded = True
                    AddFolderToTree x, pp
                Next
            End If
        End If
        
        
End Function
    
Private Sub mnuAddGroup_Click()
    On Error Resume Next
    Dim nn As String, f As String
    
    If selNode Is Nothing Then Set selNode = tv.Nodes(1)
    
    nn = selNode.Text
    
    f = InputBox("Enter name of new folder to add under " & nn)
    If Len(f) = 0 Then Exit Sub
    
    tv.Nodes.Add selNode, tvwChild, f, f, "folder"
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuFind_Click()
    frmFind.init tv
End Sub

Private Sub mnuRemoveItem_Click()

    Dim n As Node
    Dim c As New Collection
    Dim i As Long
    
    If selNode Is Nothing Then Exit Sub
    
    'MsgBox selNode.Image
   
    If selNode.Children > 0 Then
        If MsgBox("Are you sure you want to delete " & selNode.Children & " nodes?", vbYesNo) = vbNo Then Exit Sub
        AllNodesUnder selNode, c
        For i = c.Count To 1 Step -1
            Set n = c(i)
            tv.Nodes.Remove n.key
        Next
        Exit Sub
    End If

    tv.Nodes.Remove selNode.key
    Set selNode = Nothing
    
    ' If selNode.Image = "folder" Then
    
    
End Sub

Private Sub tv_DblClick()
    If selNode Is Nothing Then Exit Sub
    If selNode.Image = "folder" Or selNode.Image = "proj" Then Exit Sub
    ipc.Send "show:" & selNode.Text
End Sub

Private Sub tv_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
    Dim nodNode As Node
    '// get the node we are over
    Set nodNode = tv.HitTest(x, y)
    If nodNode Is Nothing Then Exit Sub '// no node
    '// ensure node is actually selected, just incase we start dragging.
    nodNode.selected = True
End Sub

Private Sub tv_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    Set selNode = Node
End Sub

'// occurs when the user starts dragging
'// this is where you assign the effect and the data.
Private Sub tv_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    AllowedEffects = vbDropEffectMove '// Set the effect to move
    Data.SetData tv.SelectedItem.key '// Assign the selected item's key to the DataObject
    blnDragging = True '// we are dragging from this control internally
End Sub



'Text = 1 (vbCFText)
'Bitmap = 2 (vbCFBitmap)
'Metafile = 3
'Emetafile = 14
'DIB = 8
'Palette = 9
'Files = 15 (vbCFFiles)
'RTF = -16639

'// occurs when the object is dragged over the control.
'// this is where you check to see if the mouse is over
'// a valid drop object
Private Sub tv_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single, state As Integer)
    Dim nodNode As Node

    Effect = vbDropEffectMove
    Set nodNode = tv.HitTest(x, y)
    
    If nodNode Is Nothing Or blnDragging = False Then
        '// the dragged object is not over a node, invalid drop target
        '// or the object is not from this control.
                
        If Not Data.GetFormat(vbCFFiles) Then 'we also accept files from the desktop
            Effect = vbDropEffectNone 'setting this will block the transfer further..
        End If
        
    End If
    
    
End Sub

Function AddNodeFromFile(p As Node, ByVal fpath As String) As Boolean
    
    Dim vbc As New CVBComponent
    Dim n As Node
    
    If Not vbc.loadFromFile(fpath) Then
        List1.AddItem "AddNodeFromFile failed: " & fpath
        Exit Function
    End If
    
    If NodeExists(vbc.name) Then
        List1.AddItem vbc.name & " already exists in tree: " & fpath
        Exit Function
    End If
    
    Set n = tv.Nodes.Add(p, tvwChild, vbc.name, vbc.name, vbc.icon)
    Set vbc.n = n
    Set n.tag = vbc
    
    'this will trigger an IPC Component|Added message, but our node will already exist so it will be ignored with warning in list1
    ipc.Send "addfile:" & fpath
    
    AddNodeFromFile = True
            
End Function

'// occurs when the user drops the object
'// this is where you move the node and its children.
'// this will not occur if Effect = vbDropEffectNone
Private Sub tv_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single)
    
    On Error Resume Next
    
    Dim strSourceKey As String
    Dim nodTarget    As Node
    Dim f As String, fn As String, p As Node, icon As String
    Dim cn As String, vbc As CVBComponent
     
    
    Set nodTarget = tv.HitTest(x, y)
    
    '// if the target node is not a folder or the root item
    '// then get it's parent (that is a folder or the root item)
    'If nodTarget.Image <> "FolderClosed" And nodTarget.Key <> "Root" Then
    '    Set nodTarget = nodTarget.Parent
    'End If
    
    If Data.GetFormat(vbCFText) Then
        
        'internal drag to rearrange nodes
        strSourceKey = Data.GetData(vbCFText)
        Set tv.Nodes(strSourceKey).Parent = nodTarget
        
    ElseIf Data.GetFormat(vbCFFiles) Then
        
        f = Data.Files(1)
        
        If fso.FolderExists(f) Then
            fn = fso.FolderName(f)
            Set p = tv.Nodes.Add(nodTarget, tvwChild, f, fn, "folder")
            AddFolderToTree f, p
        Else
            AddNodeFromFile nodTarget, f
        End If
    End If
    
    '// NOTE: You will also need to update the key to reflect the changes if you are using it
    blnDragging = False
    Effect = 0 '// cancel effect so that VB doesn't muck up your transfer
    
End Sub



