VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   11145
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   19080
   _ExtentX        =   33655
   _ExtentY        =   19659
   _Version        =   393216
   Description     =   "IPC Server for external  VB6 TreeView Explorer"
   DisplayName     =   "TreeView Explorer"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mForm As Form1
Private mainProjPath As String
Private mProjectLoaded As Boolean

Public WithEvents ipc As CIpc
Attribute ipc.VB_VarHelpID = -1

Private mProjectTreeHwnd As Long
Private mcbMenuCommandBar     As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
Private WithEvents mFileEvents As FileControlEvents
Attribute mFileEvents.VB_VarHelpID = -1
Private WithEvents mProjectEvents As VBProjectsEvents
Attribute mProjectEvents.VB_VarHelpID = -1
Private WithEvents mComponentEvents As VBComponentsEvents
Attribute mComponentEvents.VB_VarHelpID = -1

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
'    On Error Resume Next
'    mProjectLoaded = True
'    mainProjPath = VBInstance.ActiveVBProject.FileName
'    ipc.Send "AddinInstance_OnStartupComplete:" & mainProjPath
'End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    'right now just showing internal debug message form, todo: show external tree explorer
    Me.Show
End Sub

Private Sub ipc_Message(m As String, retval As Long)
    
    On Error Resume Next
     
    If Len(m) = 0 Then Exit Sub
      
    If m = "projpath" Then
        ipc.Send mainProjPath 'cached so we can send it immediatly..
        Exit Sub
    End If
      
    If Left(m, 5) = "hwnd:" Then
        ipc.RemoteHWND = Mid(m, 6)
        Exit Sub
    End If
    
    If m = "treehwnd" Then
        retval = CStr(mProjectTreeHwnd) 'received as return from SendMessage directly..
        Exit Sub
    End If
    
    'we need to avoid error: An outgoing call cannot be made since the application is dispatching an input-synchronous call.
    mForm.queCommand m
    
End Sub


''Relevant Project events
Private Sub mComponentEvents_ItemAdded(ByVal c As VBIDE.VBComponent)

    If Not mProjectLoaded Then Exit Sub
    ipc.Send "Component|Added|" & c.Collection.Parent.Name & "|" & c.Type & "|" & c.Name & "|" & c.FileNames(1)

   'If mProjectLoading Then Exit Sub 'we'll process all the components when the project has finished loading
   'AddNewItem vbNullString, KeyForComponent(VBComponent), CaptionForComponent(VBComponent), VBComponent.Type, True, False
End Sub

Private Sub mComponentEvents_ItemRemoved(ByVal c As VBIDE.VBComponent)

    If Not mProjectLoaded Then Exit Sub
    ipc.Send "Component|Removed|" & c.Collection.Parent.Name & "|" & c.Type & "|" & c.Name & "|" & c.FileNames(1)

End Sub

Private Sub mComponentEvents_ItemRenamed(ByVal c As VBIDE.VBComponent, ByVal OldName As String)

    If Not mProjectLoaded Then Exit Sub
    ipc.Send "Component|Renamed|" & c.Collection.Parent.Name & "|" & c.Type & "|" & c.Name & "|" & c.FileNames(1) & "|" & OldName

End Sub

Private Sub mComponentEvents_ItemSelected(ByVal c As VBIDE.VBComponent)

    If Not mProjectLoaded Then Exit Sub
    ipc.Send "Component|Selected|" & c.Collection.Parent.Name & "|" & c.Type & "|" & c.Name & "|" & c.FileNames(1)

'Dim i As Long
'   If mProjectLoading Then Exit Sub
'
'   If Not (cmbProject.ListIndex = 0 Or VBComponent.Collection.Parent.Name = cmbProject.Text) Then
'      For i = 1 To cmbProject.ListCount - 1
'         If cmbProject.List(i) = VBComponent.Collection.Parent.Name Then
'            cmbProject.ListIndex = i
'            Exit For
'         End If
'      Next i
'   End If
'   ucProjEx.SelectItem KeyForComponent(VBComponent), True
End Sub


'these are not triggered just from doing Add Form or remove Form..
Private Sub mFileEvents_AfterAddFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
    ipc.Send "FileEvents_AddFile: " & FileName
End Sub

'they did a save as, changing form name in properties does not trigger this...
Private Sub mFileEvents_AfterChangeFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal NewName As String, ByVal OldName As String)
    ipc.Send "FileEvents_ChangeFileName: " & OldName & "|" & NewName
End Sub

Private Sub mFileEvents_AfterRemoveFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String)
    ipc.Send "FileEvents_RemoveFile: " & FileName
End Sub



'(there'll be lots of mComponentEvents_ItemAdded events between these two, but we're supressing them via our mProjectLoading flag)
Private Sub mProjectEvents_ItemAdded(ByVal VBProject As VBIDE.VBProject) '... and this marks the end of a project loading
'   mProjectLoading = False
'   If Len(VBProject.FileName) Then InitProject VBProject ' If len=0 it's not a saved project & we'll catch the component-add events
'   RefreshProjectList vbNullString
End Sub

Private Sub mProjectEvents_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
'    Dim Group As clsGroup, Item As clsItem, s() As String
'
'   If VBProject.Name = vbNullString Then
'      mProjectLoading = False 'a project without a name that is removed is one that failed to load
'   Else
'      'If Len(VBProject.FileName) Then WritePexFile VBProject
'
'      For Each Group In ucProjEx.Groups
'         For Each Item In Group.Items
'            s = Split(Item.Key, "|")
'            If s(1) = VBProject.Name Then
'               ucProjEx.RemoveItem Item.Key, Group.Key, True
'            End If
'         Next Item
'      Next Group
'   End If
'
'   If gVBInstance.VBProjects.Count = 1 Then ucProjEx.Clear Else ucProjEx.Init
'   RefreshProjectList VBProject.Name
End Sub


Sub Hide()
    On Error Resume Next
    mForm.Visible = False
End Sub

Sub Show()
   On Error Resume Next
   mForm.Visible = True
End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    Dim e1 As String
    Dim hWnd As Long
     
   If ConnectMode = ext_cm_AfterStartup Then
        mProjectLoaded = True
        mainProjPath = VBInstance.ActiveVBProject.FileName
        ipc.Send "AddinInstance_OnStartupComplete:" & mainProjPath
        Exit Sub
   End If
   
   If ConnectMode = ext_cm_Startup Then
         
        Set VBInstance = Application
        
        Set mForm = New Form1
        Set ipc = New CIpc
        Set Module1.ipc = Me.ipc 'we need a copy for the form, but must sink events here...
        
        Set mcbMenuCommandBar = AddToAddInCommandBar("VB Proj Exp")
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)

        hWnd = FindWindowEx(VBInstance.MainWindow.hWnd, 0, "PROJECT", vbNullString)
        mProjectTreeHwnd = FindWindowEx(hWnd, 0, "SysTreeView32", vbNullString)
       
        If Not memfile.CreateMemMapFile("ProjectExplorer", 200000, True) Then
            e1 = memfile.ErrorMessage
            If Not memfile.OpenMemMapFile("ProjectExplorer", 200000) Then 'may already exist shared all instances..
                mForm.List1.AddItem "CreateMemMapFile error:" & e1 & ", OpenExisting: " & memfile.ErrorMessage
            End If
        End If
        
        If Not ipc.Listen(mForm, "ProjectExplorer") Then
            mForm.List1.AddItem "Ipc.Listen failed"
        Else
            If Not ipc.FindClient("Treeview") Then
                 mForm.List1.AddItem "ipc.FindClient(Treeview) failed"
            Else
                mForm.List1.AddItem "found server: " & ipc.RemoteHWND
            End If
        End If
        
        Set mFileEvents = VBInstance.Events.FileControlEvents(Nothing)
        Set mComponentEvents = VBInstance.Events.VBComponentsEvents(Nothing)
        Set mProjectEvents = VBInstance.Events.VBProjectsEvents
        
    '    For Each thisProj In gVBInstance.VBProjects
    '      InitProject thisProj
    '    Next thisProj
  
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    unloading = True
    mcbMenuCommandBar.Delete
    Unload mForm
    
    Set VBInstance = Nothing
    Set mFileEvents = Nothing
    Set mComponentEvents = Nothing
    Set mProjectEvents = Nothing
    
End Sub



Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        Exit Function
    End If
    
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function




