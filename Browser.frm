VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "VB Code Browser"
   ClientHeight    =   6315
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2595
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6315
   ScaleWidth      =   2595
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.TreeView ctlProjectParts 
      Height          =   3345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   5900
      _Version        =   327680
      Indentation     =   529
      LabelEdit       =   1
      Style           =   6
      Appearance      =   1
      MouseIcon       =   "Browser.frx":0000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   225
      Top             =   3645
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      Filter          =   "VB Projects(*.vbp)|*.vbp"
   End
   Begin VB.Menu Menu_File 
      Caption         =   "&File"
      Begin VB.Menu Menu_File_Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu Menu_File_Reopen 
         Caption         =   "&Reopen"
      End
   End
   Begin VB.Menu Menu_Help 
      Caption         =   "&Help"
      Begin VB.Menu Menu_Help_About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'---- Private Properties --------------------------------

Private sProjectPath As String
Private sProjectFile As String
Private sProjectContents As String
Private oPartNames As Collection
Private oPartTypes As Collection
Private oPartForms As Collection



'---- Public Properties --------------------------------


'---- Public Methods --------------------------------


'---- Private Methods --------------------------------
Public Sub LoadProject()
    Dim oNode As Node
    Dim sContents As String, sLine As String, nPos As Integer
    Dim sKey As String, sValue As String
    
    Call ResetCodeViews
    
    sProjectContents = ReadFile(sProjectFile)
    
    'Set up the basic structure
    ctlProjectParts.Nodes.Clear
    Set oNode = ctlProjectParts.Nodes.Add(, , "Project", "Project")
    Set oNode = ctlProjectParts.Nodes.Add("Project", tvwChild, "Project/Form", "Forms")
    oNode.EnsureVisible
    Set oNode = ctlProjectParts.Nodes.Add("Project", tvwChild, "Project/Class", "Classes")
    oNode.EnsureVisible
    Set oNode = ctlProjectParts.Nodes.Add("Project", tvwChild, "Project/Control", "Controls")
    oNode.EnsureVisible
    Set oNode = ctlProjectParts.Nodes.Add("Project", tvwChild, "Project/Module", "Modules")
    oNode.EnsureVisible
    
    'Parse the project file
    sContents = sProjectContents
    Do
        nPos = InStr(1, sContents, vbCrLf)
        If nPos = 0 Then
            If sContents <> "" Then
                sLine = sContents
            Else
                Exit Do
            End If
        Else
            sLine = Left(sContents, nPos - 1)
            sContents = Right(sContents, Len(sContents) - nPos - 1)
        End If
        
        nPos = InStr(1, sLine, "=")
        If nPos <> 0 Then
            sKey = Left(sLine, nPos - 1)
            sValue = Right(sLine, Len(sLine) - nPos)
            If sKey = "Form" Then
                Set oNode = ctlProjectParts.Nodes.Add("Project/Form", tvwChild, "Project/Form/" & sValue, sValue)
                oNode.EnsureVisible
            ElseIf sKey = "Class" Then
                nPos = InStr(1, sValue, "; ")
                Set oNode = ctlProjectParts.Nodes.Add("Project/Class", tvwChild, "Project/Class/" & Right(sValue, Len(sValue) - nPos - 1), Right(sValue, Len(sValue) - nPos - 1))
                oNode.EnsureVisible
            ElseIf sKey = "Module" Then
                nPos = InStr(1, sValue, "; ")
                Set oNode = ctlProjectParts.Nodes.Add("Project/Module", tvwChild, "Project/Module/" & Right(sValue, Len(sValue) - nPos - 1), Right(sValue, Len(sValue) - nPos - 1))
                oNode.EnsureVisible
            ElseIf sKey = "UserControl" Then
                Set oNode = ctlProjectParts.Nodes.Add("Project/Control", tvwChild, "Project/Control/" & sValue, sValue)
                oNode.EnsureVisible
            ElseIf sKey = "Name" Then
                sValue = Left(sValue, Len(sValue) - 1)
                sValue = Right(sValue, Len(sValue) - 1)
                Me.Caption = sValue & " - VB Code Browser"
            End If
        End If
    Loop
    
    'Select the root item
    Set ctlProjectParts.SelectedItem = ctlProjectParts.Nodes(1)
End Sub

Private Sub ResetCodeViews()
    If Not oPartForms Is Nothing Then
        Dim i As Integer
        For i = 1 To oPartForms.Count
            oPartForms.Item(i).AllowClose = True
            Unload oPartForms.Item(i)
        Next i
    End If
    Set oPartNames = New Collection
    Set oPartTypes = New Collection
    Set oPartForms = New Collection
End Sub

Public Sub OpenPart(sType As String, sFilename As String)
    Dim sFileKey As String, oForm As Form
    sFileKey = sType & "/" & sFilename
    On Error Resume Next
    Call oPartNames.Item(sFileKey)
    If Err Then 'Not already opened
        On Error GoTo 0
        Set oForm = New frmCodeView
        Call oPartNames.Add(sFileKey, sFileKey)
        Call oPartTypes.Add(sType, sFileKey)
        Call oPartForms.Add(oForm, sFileKey)
        
        oForm.ctlCode.FileType = sType
        oForm.ctlCode.FilePath = CurDir & "\"
        oForm.ctlCode.FileName = sFilename
        oForm.Caption = sType & " - " & sFilename
        oForm.Left = Me.Left + Me.Width 'Docking
        oForm.Top = Me.Top 'Docking
        oForm.Show
        Call oForm.ctlCode.LoadFile
        On Error Resume Next
        oForm.ctlCode.SetFocus
        On Error GoTo 0
    
    Else 'Already open
        On Error GoTo 0
        Set oForm = oPartForms.Item(sFileKey)
        If oForm.Visible Then
            Call oForm.Hide
        Else
            Call oForm.Show
        End If
    End If
End Sub


'---- Private Event Handlers --------------------------------

Private Sub Form_Load()
    Call ResetCodeViews
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ResetCodeViews
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Height < 1000 Then
        Me.Height = 1000
        Exit Sub
    End If
    If Me.Width < 400 Then
        Me.Width = 400
        Exit Sub
    End If
    ctlProjectParts.Width = Me.ScaleWidth - 2 * ctlProjectParts.Left
    ctlProjectParts.Height = Me.ScaleHeight - 2 * ctlProjectParts.Left
End Sub

Private Sub ctlProjectParts_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sFilename As String
    If Data.Files.Count > 0 Then
        On Error GoTo Error
        sFilename = Data.Files.Item(1)
        If sFilename = "" Then Exit Sub
        sProjectPath = PathFromFile(sFilename, sProjectFile)
        ChDrive sProjectPath
        ChDir sProjectPath
        Call Menu_File_Reopen_Click
    End If
    Exit Sub
    
Error:
    MsgBox Err.Description
End Sub

Private Sub Menu_File_Open_Click()
    Dim sFilename As String
    CommonDialog1.ShowOpen
    sFilename = CommonDialog1.FileName
    If sFilename = "" Then Exit Sub
    sProjectPath = PathFromFile(sFilename, sProjectFile)
    ChDrive sProjectPath
    ChDir sProjectPath
    Call Menu_File_Reopen_Click
End Sub

Private Sub Menu_File_Reopen_Click()
    Call LoadProject
End Sub

Private Sub ctlProjectParts_Click()
    If ctlProjectParts.Nodes.Count = 0 Then Exit Sub
    Dim oNode As Node, sKey As String, sFileType As String
    Dim nPos As Integer
    Set oNode = ctlProjectParts.SelectedItem
    sKey = oNode.Key
    If sKey = "Project" Then
        'Call Shell("Notepad """ & sProjectPath & sProjectFile & """", vbNormalFocus)
    Else
        sKey = Mid(sKey, InStr(1, sKey, "/") + 1)
        nPos = InStr(1, sKey, "/")
        If nPos = 0 Then
            sFileType = sKey
        Else
'            sFileType = Left(sKey, nPos - 1)
'            sKey = Mid(sKey, nPos + 1)
    
            Dim sFileKey As String, oForm As Form
            On Error Resume Next
            Set oForm = oPartForms.Item(sKey)
            If Err Then
            Else 'Already opened
                On Error GoTo 0
                If oForm.Visible Then
                    oForm.SetFocus
                End If
            End If
            On Error GoTo 0
                
        End If
    End If
End Sub

Private Sub ctlProjectParts_DblClick()
    If ctlProjectParts.Nodes.Count = 0 Then Exit Sub
    Dim oNode As Node, sKey As String, sFileType As String
    Dim nPos As Integer
    Set oNode = ctlProjectParts.SelectedItem
    sKey = oNode.Key
    If sKey = "Project" Then
        'Call Shell("Notepad """ & sProjectPath & sProjectFile & """", vbNormalFocus)
    Else
        sKey = Mid(sKey, InStr(1, sKey, "/") + 1)
        nPos = InStr(1, sKey, "/")
        If nPos = 0 Then
            sFileType = sKey
        Else
            sFileType = Left(sKey, nPos - 1)
            sKey = Mid(sKey, nPos + 1)
            Call OpenPart(sFileType, sKey)
        End If
    End If
End Sub

Private Sub Menu_Help_About_Click()
    Dim sMessage As String
    sMessage = sMessage & "Visual Basic Code Browser" & vbCrLf
    sMessage = sMessage & "by James Vincent Carnicelli, II" & vbCrLf
    sMessage = sMessage & "20 March 1998" & vbCrLf
    MsgBox sMessage
End Sub
