Attribute VB_Name = "Module1"
Option Explicit

Global fso As New CFileSystem2
Global dlg As New CCmnDlg

Function PreloadComponentName(fpath As String) As String
    'Attribute VB_Name = "Connect"
    On Error GoTo hell

    Dim fs As New clsFileStream, x As String, i As Long
    Const marker = "Attribute VB_Name = "
    
    fs.fOpen fpath, otreading
    
    While Not fs.EndOfFile
        x = fs.ReadLine
        If InStr(x, marker) > 0 Then
            x = Replace(x, marker, Empty)
            x = Replace(x, vbCr, Empty)
            x = Replace(x, vbLf, Empty)
            x = Trim(Replace(x, """", Empty))
            PreloadComponentName = x
            fs.fClose
            Exit Function
        End If
    Wend
    
hell:
           
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
 

Sub AllNodesUnder(ByVal n As Node, c As Collection)
    
    Dim nn As Node
    c.Add n
    
    For Each nn In Form1.tv.Nodes
        If Not nn.Parent Is Nothing Then
            If nn.Parent = n Then
                c.Add nn
                If nn.Children > 0 Then AllNodesUnder nn, c
            End If
        End If
    Next
    
End Sub

Function keyExistsInCollection(key As String, c As Collection, Optional isObj As Boolean = True) As Boolean
    On Error Resume Next
    Dim o As Object, x
    If isObj Then
        Set o = c(key)
    Else
        x = c(key)
    End If
    keyExistsInCollection = (Err.Number = 0)
End Function

Function ComponentExists(name As String, Optional ByRef c As CVBComponent) As Boolean
    On Error Resume Next
    Dim n As Node
    Set c = Nothing
    If NodeExists(name, n) Then
        Set c = n.tag
        ComponentExists = (Not c Is Nothing)
    End If
End Function

Function NodeExists(key As String, Optional ByRef n As Node) As Boolean
    On Error Resume Next
    Set n = Form1.tv.Nodes(key)
    NodeExists = (Err.Number = 0)
End Function

Function HandleComponentEvent(e As CComponentEvent, Optional createMissing As Boolean = True) As CVBComponent
    
    'On Error Resume Next
    
    Dim c As CVBComponent
    Dim n As String
    Dim p As Node, nn As Node
    
    n = e.ComponentName
    If e.EventType = ec_Rename Then n = e.OldName
    
    If ComponentExists(n, c) Then
    
        Set HandleComponentEvent = c
        
        If e.EventType = ec_Remove Then
            Form1.tv.Nodes.Remove c.n.key
            Set c = Nothing
            Set HandleComponentEvent = Nothing
            Exit Function
        End If
            
        If e.EventType = ec_Rename Then
            c.name = e.ComponentName
            If Not c.n Is Nothing Then
                Set p = c.n.Parent
                Form1.tv.Nodes.Remove c.n.key 'we need to reset its key and new name text
                Set nn = Form1.tv.Nodes.Add(p, tvwChild, c.name, c.name, c.icon)
                Set c.n = nn
                Set nn.tag = c
            End If
            Exit Function
        End If
        
    Else
        If createMissing Then
            Set c = New CVBComponent
            c.loadFromEvent e
            
            If Not NodeExists(c.defFolder, p) Then
                Set p = Form1.tv.Nodes.Add(Form1.tv.Nodes(1), tvwChild, c.defFolder, c.defFolder, "folder")
            End If
            
            If NodeExists(c.name) Then
                Form1.List1.AddItem "HandleComponentEvent Node exists: " & e.raw
            Else
                Set c.n = Form1.tv.Nodes.Add(p, tvwChild, c.name, c.name, c.icon)
                Set c.n.tag = c
            End If
        End If
    End If

End Function


'Public Enum vbext_ComponentType
'    vbext_ct_StdModule = 1
'    vbext_ct_ClassModule = 2
'    vbext_ct_MSForm = 3
'    vbext_ct_ResFile = 4
'    vbext_ct_VBForm = 5
'    vbext_ct_VBMDIForm = 6
'    vbext_ct_PropPage = 7
'    vbext_ct_UserControl = 8
'    vbext_ct_DocObject = 9
'    vbext_ct_RelatedDocument = &HA
'    vbext_ct_ActiveXDesigner = &HB
'End Enum

Function typeFromPath(fpath As String) As Long

    On Error Resume Next
    Dim ext As String, i As Long
    
    ext = LCase(fso.GetExtension(fpath))
    If Left(ext, 1) = "." Then ext = Mid(ext, 2)
    
    Select Case ext
        Case "bas": i = 1
        Case "cls": i = 2
        Case "frm": i = 3
        Case "res": i = 4
        Case "frm": i = 5
        Case "mdi": i = 6
        Case "pag": i = 7
        Case "ctl": i = 8
        Case "dob": i = 9
        Case "txt": i = 10
        Case "dsr": i = 11
    End Select
      
    typeFromPath = i
    
End Function

Function DefaultFolderForType(t As Long, Optional ByRef icon As String) As String
    On Error Resume Next
    Dim tn  As String
    
    Select Case t
        Case 1: tn = "Modules"
                icon = "bas"
                
        Case 2: tn = "Classes"
                icon = "cls"

        Case 3: tn = "Forms"
                icon = "frm"

        Case 4: tn = "Resources"
                icon = "res"

        Case 5: tn = "Forms"
                icon = "frm"

        Case 6: tn = "Forms"
                icon = "mdi"

        Case 7: tn = "Property Pages"
                icon = "pag"

        Case 8: tn = "User Controls"
                icon = "ctl"

        Case 9: tn = "ActiveX Documents"
                icon = "dob"

        Case 10: tn = "Related Documents"
                icon = "txt"

        Case 11: tn = "Designers"
                icon = "dsr"

        Case Default:
                tn = "Unknown"
                icon = "unk"
                
    End Select
      
    DefaultFolderForType = tn

End Function
