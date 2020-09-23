VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.UserControl ctlCodeView 
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   ScaleHeight     =   3810
   ScaleWidth      =   7800
   Begin ComctlLib.ProgressBar ctlProgress 
      Height          =   240
      Left            =   2340
      TabIndex        =   7
      Top             =   1710
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   423
      _Version        =   327680
      Appearance      =   0
   End
   Begin ComctlLib.TreeView ctlAspects 
      Height          =   3795
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   6694
      _Version        =   327680
      Indentation     =   529
      LabelEdit       =   1
      Style           =   6
      Appearance      =   1
      MouseIcon       =   "ProjectView.ctx":0000
   End
   Begin VB.CheckBox ctlWrap 
      Caption         =   "Wrap"
      Height          =   240
      Left            =   5535
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   735
   End
   Begin VB.CommandButton ctlCopy 
      Caption         =   "Copy"
      Height          =   300
      Left            =   6300
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   600
   End
   Begin RichTextLib.RichTextBox ctlDetail 
      Height          =   1005
      Left            =   2340
      TabIndex        =   1
      Top             =   300
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   1773
      _Version        =   327680
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"ProjectView.ctx":001C
   End
   Begin VB.Label ctlDetailStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2340
      TabIndex        =   5
      Top             =   0
      Width           =   3120
   End
   Begin VB.Label ctlProgressStatus 
      Height          =   240
      Left            =   2430
      TabIndex        =   4
      Top             =   1350
      Visible         =   0   'False
      Width           =   5235
   End
   Begin VB.Label ctlSizerMiddle 
      Height          =   3795
      Left            =   2205
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "ctlCodeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VB Code Member Viewer
'by James Vincent Carnicelli, II
'begun 18 March 1998

Option Explicit

'---- Private Constants --------------------------------
Const LineWidth = 70



'---- Private Properties --------------------------------
Private dProportionMiddle As Double
Private bSizingMiddle As Boolean
Private sFileContents As String
Private sFileContentsRtf As String
Private sWrappedFileContentsRtf As String
Private lFileLines As Long
Private sCurrentlyViewing As String

Private nTokenNumber As Integer
Private oTokens As Collection
Private oTokenTypes As Collection
Private sCommentsAccum As String

Private oMemberNames As Collection
Private oMemberVisibility As Collection
Private oMemberTypes As Collection
Private oMemberReturnTypes As Collection
Private oMemberParamLists As Collection
Private oMemberReadable As Collection
Private oMemberWritable As Collection
Private oMemberComments As Collection


'---- Public Properties --------------------------------

Public FileType As String

Public FilePath As String

Public FileName As String


'---- Public Methods --------------------------------

Public Sub LoadFile()
    Call ShowProgress("Processing...")
    Call ResetDetail
    sFileContents = ReadFile(FilePath & FileName)
    ctlDetail.Text = sFileContents
    lFileLines = CountLines(sFileContents)
    Call LoadCode
    Call ParseCode
    sFileContentsRtf = ctlDetail.TextRTF
    '    call ctlAspects_Click
    Call ctlAspects_NodeClick(ctlAspects.Nodes(1))
    If ctlWrap = 1 Then
        Call WordWrap
    End If
    Call HideProgress
End Sub


'---- Private Methods --------------------------------

Private Sub LoadCode()
    Dim oNode As Node
    ctlAspects.Nodes.Clear
    Set oNode = ctlAspects.Nodes.Add(, , "Root", FileType)
    Set oNode = ctlAspects.Nodes.Add("Root", tvwChild, "Root/Public", "Public")
    oNode.EnsureVisible
    Set oNode = ctlAspects.Nodes.Add("Root", tvwChild, "Root/Private", "Private")
    oNode.EnsureVisible
    Set oNode = ctlAspects.Nodes.Add("Root/Public", tvwChild, "Root/Public/Property", "Properties")
    oNode.EnsureVisible
    Set oNode = ctlAspects.Nodes.Add("Root/Public", tvwChild, "Root/Public/Method", "Methods")
    oNode.EnsureVisible
    Set oNode = ctlAspects.Nodes.Add("Root/Public", tvwChild, "Root/Public/Event", "Events")
    oNode.EnsureVisible
    Set oNode = ctlAspects.Nodes.Add("Root/Private", tvwChild, "Root/Private/Property", "Properties")
    oNode.EnsureVisible
    Set oNode = ctlAspects.Nodes.Add("Root/Private", tvwChild, "Root/Private/Method", "Methods")
    oNode.EnsureVisible
    Set ctlAspects.SelectedItem = ctlAspects.Nodes(1)
End Sub


Private Sub ShowProgress(sCaption As String)
    ctlDetail.Visible = False
    ctlAspects.Enabled = False
    ctlCopy.Enabled = False
    ctlWrap.Enabled = False
    ctlProgressStatus.Visible = True
    ctlProgress.Visible = True
    Call SetProgress(sCaption, 0)
End Sub

Private Sub HideProgress()
    ctlProgressStatus.Visible = False
    ctlProgress.Visible = False
    ctlAspects.Enabled = True
    ctlCopy.Enabled = True
    ctlWrap.Enabled = True
    ctlDetail.Visible = True
End Sub

Private Sub SetProgress(sCaption As String, dProgress As Double)
    If Int(dProgress) <> ctlProgress.Value Then
        ctlProgressStatus.Caption = sCaption
        ctlProgressStatus.Refresh
        ctlProgress.Value = Int(dProgress)
        DoEvents
    End If
End Sub

Private Sub ResetDetail()
    ctlDetailStatus.Caption = ""
    ctlDetail.Text = ""
    ctlDetail.SelFontName = "Courier New"
    ctlDetail.SelFontSize = 10
    ctlDetail.SelColor = RGB(0, 0, 0)
    ctlDetail.SelBold = False
End Sub

Public Sub UpdateLineNumber()
    ctlDetailStatus.Caption = "Line " & CountLines(ctlDetail.Text, ctlDetail.SelStart + 1) & " of " & CountLines(ctlDetail.Text)
End Sub

Private Function CountLines(sContents As String, Optional lCharPos) As Long
    If IsMissing(lCharPos) Then
        lCharPos = Len(sContents)
    End If
    Dim lPos As Long
    CountLines = 1
    lPos = -1
    Do
        lPos = InStr(lPos + 2, sContents, vbCrLf)
        If lPos = 0 Or lPos >= lCharPos Then Exit Do
        CountLines = CountLines + 1
    Loop
End Function

Private Function MatchModifyCase(a As String, b As String) As Boolean
    If UCase(a) = UCase(b) Then
        a = b
        MatchModifyCase = True
    End If
End Function

Private Function ReservedWord(sToken As String) As Boolean
    Dim sUpperToken As String
    ReservedWord = True
    If MatchModifyCase(sToken, "Public") Then Exit Function
    If MatchModifyCase(sToken, "Private") Then Exit Function
    If MatchModifyCase(sToken, "Friend") Then Exit Function
    If MatchModifyCase(sToken, "Static") Then Exit Function
    If MatchModifyCase(sToken, "Sub") Then Exit Function
    If MatchModifyCase(sToken, "Function") Then Exit Function
    If MatchModifyCase(sToken, "Property") Then Exit Function
    If MatchModifyCase(sToken, "Event") Then Exit Function
    If MatchModifyCase(sToken, "And") Then Exit Function
    If MatchModifyCase(sToken, "Or") Then Exit Function
    If MatchModifyCase(sToken, "Not") Then Exit Function
    If MatchModifyCase(sToken, "On") Then Exit Function
    If MatchModifyCase(sToken, "Error") Then Exit Function
    If MatchModifyCase(sToken, "Is") Then Exit Function
    If MatchModifyCase(sToken, "Dim") Then Exit Function
    If MatchModifyCase(sToken, "ReDim") Then Exit Function
    If MatchModifyCase(sToken, "True") Then Exit Function
    If MatchModifyCase(sToken, "False") Then Exit Function
    If MatchModifyCase(sToken, "Exit") Then Exit Function
    If MatchModifyCase(sToken, "End") Then Exit Function
    If MatchModifyCase(sToken, "As") Then Exit Function
    If MatchModifyCase(sToken, "String") Then Exit Function
    If MatchModifyCase(sToken, "Byte") Then Exit Function
    If MatchModifyCase(sToken, "Int") Then Exit Function
    If MatchModifyCase(sToken, "Long") Then Exit Function
    If MatchModifyCase(sToken, "Double") Then Exit Function
    If MatchModifyCase(sToken, "Variant") Then Exit Function
    If MatchModifyCase(sToken, "If") Then Exit Function
    If MatchModifyCase(sToken, "Then") Then Exit Function
    If MatchModifyCase(sToken, "Else") Then Exit Function
    If MatchModifyCase(sToken, "ElseIf") Then Exit Function
    If MatchModifyCase(sToken, "Call") Then Exit Function
    If MatchModifyCase(sToken, "Do") Then Exit Function
    If MatchModifyCase(sToken, "Loop") Then Exit Function
    If MatchModifyCase(sToken, "For") Then Exit Function
    If MatchModifyCase(sToken, "Next") Then Exit Function
    If MatchModifyCase(sToken, "While") Then Exit Function
    If MatchModifyCase(sToken, "Wend") Then Exit Function
    If MatchModifyCase(sToken, "GoTo") Then Exit Function
    If MatchModifyCase(sToken, "GoSub") Then Exit Function
    If MatchModifyCase(sToken, "Return") Then Exit Function
    If MatchModifyCase(sToken, "Get") Then Exit Function
    If MatchModifyCase(sToken, "Let") Then Exit Function
    If MatchModifyCase(sToken, "Set") Then Exit Function
    If MatchModifyCase(sToken, "ByVal") Then Exit Function
    If MatchModifyCase(sToken, "ByRef") Then Exit Function
    If MatchModifyCase(sToken, "Optional") Then Exit Function
    If MatchModifyCase(sToken, "ParamArray") Then Exit Function
    ReservedWord = False
End Function

Private Sub ParseCode()
    Dim lCodeLen As Long, nTokenType As Integer, sToken As String
    Dim lStartAt As Long, lEndAt As Long, lLine As Long
    Dim bInsideDefinition As Boolean
    lCodeLen = Len(sFileContents)
    lStartAt = 1
    lLine = 1
    
    Set oTokens = New Collection
    Set oTokenTypes = New Collection
    Set oMemberNames = New Collection
    Set oMemberTypes = New Collection
    Set oMemberVisibility = New Collection
    Set oMemberReturnTypes = New Collection
    Set oMemberParamLists = New Collection
    Set oMemberReadable = New Collection
    Set oMemberWritable = New Collection
    Set oMemberComments = New Collection
    
    Do
        sToken = NextToken(sFileContents, lCodeLen, lStartAt, lEndAt, lLine, nTokenType)
        If sToken = "" Then Exit Do
        ctlDetail.SelStart = lStartAt - 1
        ctlDetail.SelLength = lEndAt - lStartAt
        If nTokenType = 1 Then 'Comment
            ctlDetail.SelColor = RGB(0, 100, 0)
        ElseIf nTokenType >= 2 And nTokenType <= 3 Then 'Data
            ctlDetail.SelColor = RGB(100, 0, 100)
        ElseIf nTokenType = 4 Then 'Identifier
            If ReservedWord(sToken) Then
                ctlDetail.SelColor = RGB(0, 0, 150)
            End If
        Else 'Symbol
            'ctlDetail.SelColor = RGB(255, 0, 0)
        End If
        
        If nTokenNumber = 1 Then
            If oTokens.Count > 0 Then
                If oTokenTypes.Item(1) = 1 Then
                    If sCommentsAccum <> "" Then
                        sCommentsAccum = sCommentsAccum & vbCrLf
                    End If
                    sCommentsAccum = sCommentsAccum & Mid(oTokens.Item(1), 2) 'Trim off leading single quote
                Else
                    Call ProcessLine
                    sCommentsAccum = ""
                End If
            Else
                sCommentsAccum = ""
            End If
            Set oTokens = New Collection
            Set oTokenTypes = New Collection
        End If
        Call oTokens.Add(sToken)
        Call oTokenTypes.Add(nTokenType)
        
        Call SetProgress("Parsing code", 100 * lLine / lFileLines)
        lStartAt = lEndAt
    Loop
End Sub

Private Function NextToken(sCode As String, lCodeLen As Long, lStartAt As Long, lEndAt As Long, lLine As Long, nTokenType As Integer)
    Dim sChar As String
    'Skip whitespace
    
    If lStartAt = 1 Then
        nTokenNumber = 0
    End If
    
    Do
        If lStartAt > lCodeLen Then Exit Function
        sChar = Mid(sCode, lStartAt, 1)
        If sChar <> " " And sChar <> vbTab And sChar <> vbCr Then
            Exit Do
        End If
        If sChar = vbCr Then
            nTokenNumber = 0
            lLine = lLine + 1
            lStartAt = lStartAt + 2
        Else
            lStartAt = lStartAt + 1
        End If
    Loop
    lEndAt = lStartAt
    
    'Determine token type
    If sChar = "'" Then 'Comment
        nTokenType = 1
        NextToken = NextToken & sChar
        lEndAt = lEndAt + 1
        
        'Scan until end of string
        Do
            If lEndAt > lCodeLen Then Exit Function
            sChar = Mid(sCode, lEndAt, 1)
            If sChar = vbCr Then
                Exit Do
            End If
            NextToken = NextToken & sChar
            lEndAt = lEndAt + 1
        Loop
        
    ElseIf sChar = """" Then 'String
        nTokenType = 2
        NextToken = NextToken & sChar
        lEndAt = lEndAt + 1
        
        'Scan until end of string
        Do
            If lEndAt > lCodeLen Then Exit Function
            sChar = Mid(sCode, lEndAt, 1)
            If sChar = """" Then
                If Mid(sCode, lEndAt + 1, 1) = """" Then  'Embeded quote character
                    NextToken = NextToken & """"
                    lEndAt = lEndAt + 1
                Else
                    NextToken = NextToken & sChar
                    lEndAt = lEndAt + 1
                    Exit Do
                End If
            End If
            NextToken = NextToken & sChar
            lEndAt = lEndAt + 1
        Loop
        
    ElseIf Asc(sChar) >= Asc("0") And Asc(sChar) <= Asc("9") Then
        nTokenType = 3 'Number
        Do
            If lEndAt > lCodeLen Then Exit Function
            sChar = Mid(sCode, lEndAt, 1)
            If (Asc(sChar) < Asc("0") Or Asc(sChar) > Asc("9")) And sChar <> "." Then
                Exit Do
            End If
            NextToken = NextToken & sChar
            lEndAt = lEndAt + 1
        Loop
    
    ElseIf (Asc(sChar) >= Asc("A") And Asc(sChar) <= Asc("Z")) Or (Asc(sChar) >= Asc("a") And Asc(sChar) <= Asc("z")) Then  'Identifier
        nTokenType = 4
        Do
        
            If lEndAt > lCodeLen Then Exit Function
            sChar = Mid(sCode, lEndAt, 1)
            'Data type suffix
            If sChar = "!" Or sChar = "@" Or sChar = "#" Or sChar = "$" Or sChar = "%" Or sChar = "&" Then
                lEndAt = lEndAt + 1
                Exit Do
            'whitespace or not(alphanumeric or '.' or '_')
            ElseIf sChar = " " Or sChar = vbTab Or sChar = vbCr _
              Or Not ((Asc(sChar) >= Asc("A") And Asc(sChar) <= Asc("Z")) Or (Asc(sChar) >= Asc("a") And Asc(sChar) <= Asc("z")) _
              Or (Asc(sChar) >= Asc("0") And Asc(sChar) <= Asc("9")) _
              Or sChar = "." Or sChar = "_") Then
                Exit Do
            End If
            NextToken = NextToken & sChar
            lEndAt = lEndAt + 1
        Loop

    ElseIf sChar = "_" Then 'Line continuation
        lStartAt = lEndAt + 1
        'Skip whitespace
        Do
            If lStartAt > lCodeLen Then Exit Function
            sChar = Mid(sCode, lStartAt, 1)
            If sChar <> " " And sChar <> vbTab And sChar <> vbCr Then
                Exit Do
            End If
            If sChar = vbCr Then
                lLine = lLine + 1
                lStartAt = lStartAt + 2
            Else
                lStartAt = lStartAt + 1
            End If
        Loop
        'Move on to next token
        NextToken = NextToken(sCode, lCodeLen, lStartAt, lEndAt, lLine, nTokenType)

    Else 'Symbol
        nTokenType = 5
        NextToken = NextToken & sChar
        lEndAt = lEndAt + 1
    End If
    
    nTokenNumber = nTokenNumber + 1

End Function

Private Sub ProcessLine()
    Dim sType As String, sName As String, sPropDirection As String
    Dim sReturnType As String, sParamList As String, sComments As String
    Dim sVisibility As String

    On Error GoTo ProcessLineError
    'Trim trailing comments
    If oTokenTypes.Item(oTokens.Count) = 1 Then
        Call oTokenTypes.Remove(oTokens.Count)
        Call oTokens.Remove(oTokens.Count)
    End If
    
    'Implicitly "Public"
    If oTokens.Item(1) = "Sub" Or oTokens.Item(1) = "Function" Or oTokens.Item(1) = "Property" Then
        Call oTokenTypes.Add(4, , 1)
        Call oTokens.Add("Public", , 1) 'Make it explicit
    End If
    
    'Public declaration
    If oTokens.Item(1) = "Public" Or oTokens.Item(1) = "Private" Then
        sVisibility = oTokens.Item(1)
        Call oTokenTypes.Remove(1)
        Call oTokens.Remove(1)
        'Parse declaration
        If oTokens.Item(1) = "Sub" Then
            sType = "Method"
            Call oTokenTypes.Remove(1)
            Call oTokens.Remove(1)
            sName = oTokens.Item(1)
            Call oTokenTypes.Remove(1)
            Call oTokens.Remove(1)
            Call oTokenTypes.Remove(1) 'Open parenthesis
            Call oTokens.Remove(1)
            Call oTokenTypes.Remove(oTokens.Count) 'Close parenthesis
            Call oTokens.Remove(oTokens.Count)
            sReturnType = "<nothing>"
        ElseIf oTokens.Item(1) = "Function" Then
            sType = "Method"
            Call oTokenTypes.Remove(1)
            Call oTokens.Remove(1)
            sName = oTokens.Item(1)
            Call oTokenTypes.Remove(1)
            Call oTokens.Remove(1)
            Call oTokenTypes.Remove(1) 'Open parenthesis
            Call oTokens.Remove(1)
            If oTokens.Item(oTokens.Count) = ")" Then
                sReturnType = "Variant"
            Else
                sReturnType = oTokens.Item(oTokens.Count)
                Call oTokenTypes.Remove(oTokens.Count)
                Call oTokens.Remove(oTokens.Count)
                Call oTokenTypes.Remove(oTokens.Count) 'As (actually, close parenthesis; next Removes remove As)
                Call oTokens.Remove(oTokens.Count)
            End If
            Call oTokenTypes.Remove(oTokens.Count) 'Close parenthesis
            Call oTokens.Remove(oTokens.Count)
        ElseIf oTokens.Item(1) = "Event" Then
            sType = "Event"
            Call oTokenTypes.Remove(1)
            Call oTokens.Remove(1)
            sName = oTokens.Item(1)
            Call oTokenTypes.Remove(1)
            Call oTokens.Remove(1)
            Call oTokenTypes.Remove(1) 'Open parenthesis
            Call oTokens.Remove(1)
            Call oTokenTypes.Remove(oTokens.Count) 'Close parenthesis
            Call oTokens.Remove(oTokens.Count)
            sReturnType = "<nothing>"
        ElseIf oTokens.Item(1) = "Property" Then
            sType = "Property"
            Call oTokenTypes.Remove(1)
            Call oTokens.Remove(1)
            sPropDirection = oTokens.Item(1) 'Get/Let/Set
            Call oTokenTypes.Remove(1)
            Call oTokens.Remove(1)
            sName = oTokens.Item(1)
            Call oTokenTypes.Remove(1) 'Open parenthesis
            Call oTokens.Remove(1)
            If sPropDirection = "Get" Then
                If oTokens.Item(oTokens.Count) = ")" Then
                    sReturnType = "Variant"
                Else
                    sReturnType = oTokens.Item(oTokens.Count)
                    Call oTokenTypes.Remove(oTokens.Count)
                    Call oTokens.Remove(oTokens.Count)
                    Call oTokenTypes.Remove(oTokens.Count) 'As (actually, close parenthesis; next Removes remove As)
                    Call oTokens.Remove(oTokens.Count)
                End If
                Call oTokenTypes.Remove(oTokens.Count) 'Close parenthesis
                Call oTokens.Remove(oTokens.Count)
            Else
                If oTokens.Item(1) = "ByVal" Or oTokens.Item(1) = "ByRef" Then
                    Call oTokenTypes.Remove(1)
                    Call oTokens.Remove(1)
                End If
                Call oTokenTypes.Remove(1) 'Var name
                Call oTokens.Remove(1)
                If oTokens.Item(1) = "As" Then
                    Call oTokenTypes.Remove(1)
                    Call oTokens.Remove(1)
                    sReturnType = oTokens.Item(oTokens.Count)
                Else
                    sReturnType = "Variant"
                End If
            End If
        Else
            sType = "Property"
            sName = oTokens.Item(1)
            Call oTokenTypes.Remove(1)
            Call oTokens.Remove(1)
            If oTokens.Count = 1 Then
                sReturnType = "Variant"
            Else
                sReturnType = oTokens.Item(oTokens.Count)
                Call oTokenTypes.Remove(oTokens.Count)
                Call oTokens.Remove(oTokens.Count)
            End If
        End If
        
        'Extract parameters, if any
        sParamList = ""
        If sType = "Method" Or sType = "Event" Then
            While oTokens.Count > 0
                If oTokens.Item(1) = "," Then
                    sParamList = RTrim(sParamList)
                End If
                sParamList = sParamList & oTokens.Item(1) & " "
                Call oTokenTypes.Remove(1)
                Call oTokens.Remove(1)
            Wend
            sParamList = RTrim(sParamList)
        End If
        
        'Add member to collection
        Dim oNode As Node, sKey As String
        sKey = sType & "/" & sName
        On Error Resume Next
        Call oMemberNames.Item(sKey)
        If Err Then 'Not already added
            On Error GoTo 0
            Call oMemberNames.Add(sName, sKey)
            Call oMemberVisibility.Add(sVisibility, sKey)
            Call oMemberTypes.Add(sType, sKey)
            Call oMemberReturnTypes.Add(sReturnType, sKey)
            Call oMemberParamLists.Add(sParamList, sKey)
            Call oMemberComments.Add(sCommentsAccum, sKey)
            If sPropDirection = "" Then 'Variable
                Call oMemberReadable.Add(True, sKey)
                Call oMemberWritable.Add(True, sKey)
            ElseIf sPropDirection = "Get" Then
                Call oMemberReadable.Add(True, sKey)
                Call oMemberWritable.Add(False, sKey)
            Else 'Let or Set
                Call oMemberReadable.Add(False, sKey)
                Call oMemberWritable.Add(True, sKey)
            End If
            Set oNode = ctlAspects.Nodes.Add("Root/" & sVisibility & "/" & sType, tvwChild, "Root/" & sVisibility & "/" & sType & "/" & sName, sName)
        Else
            On Error Resume Next
            If sPropDirection = "Get" Then
                Call oMemberReadable.Remove(sKey)
                Call oMemberReadable.Add(True, sKey)
            Else
                Call oMemberWritable.Remove(sKey)
                Call oMemberWritable.Add(True, sKey)
            End If
            On Error GoTo 0
        End If
    End If
    
ProcessLineError: 'Skip the other tests
End Sub

Private Sub DetailProperties(sVisibility As String)
    Dim i As Integer, nCount As Integer
    For i = 1 To oMemberNames.Count
        If oMemberVisibility.Item(i) = sVisibility And oMemberTypes.Item(i) = "Property" Then
            nCount = nCount + 1
            If oMemberReadable.Item(i) And Not oMemberWritable.Item(i) Then
                ctlDetail.SelText = "R   "
            ElseIf Not oMemberReadable.Item(i) And oMemberWritable.Item(i) Then
                ctlDetail.SelText = "W   "
            Else 'Readable and writable
                ctlDetail.SelText = "R/W "
            End If
            ctlDetail.SelColor = RGB(0, 0, 100)
            ctlDetail.SelBold = True
            ctlDetail.SelText = oMemberNames.Item(i)
            ctlDetail.SelBold = False
            ctlDetail.SelColor = RGB(0, 0, 0)
            ctlDetail.SelText = " As "
            ctlDetail.SelColor = RGB(100, 0, 100)
            ctlDetail.SelText = oMemberReturnTypes.Item(i)
            ctlDetail.SelColor = RGB(0, 0, 0)
            ctlDetail.SelText = vbCrLf
        End If
    Next i
    If nCount = 0 Then
        ctlDetail.SelText = "<none>" & vbCrLf
    End If
End Sub

Private Sub DetailMethods(sVisibility As String)
    Dim i As Integer, nCount As Integer
    For i = 1 To oMemberNames.Count
        If oMemberVisibility.Item(i) = sVisibility And oMemberTypes.Item(i) = "Method" Then
            nCount = nCount + 1
            ctlDetail.SelColor = RGB(0, 0, 100)
            ctlDetail.SelBold = True
            ctlDetail.SelText = oMemberNames.Item(i)
            ctlDetail.SelBold = False
            ctlDetail.SelColor = RGB(0, 0, 0)
            ctlDetail.SelText = "(" & oMemberParamLists.Item(i) & ")"
            ctlDetail.SelText = " As "
            ctlDetail.SelColor = RGB(100, 0, 100)
            ctlDetail.SelText = oMemberReturnTypes.Item(i)
            ctlDetail.SelColor = RGB(0, 0, 0)
            ctlDetail.SelText = vbCrLf
        End If
    Next i
    If nCount = 0 Then
        ctlDetail.SelText = "<none>" & vbCrLf
    End If
End Sub

Private Sub DetailEvents(sVisibility As String)
    Dim i As Integer, nCount As Integer
    For i = 1 To oMemberNames.Count
        If oMemberVisibility.Item(i) = sVisibility And oMemberTypes.Item(i) = "Event" Then
            nCount = nCount + 1
            ctlDetail.SelColor = RGB(0, 0, 100)
            ctlDetail.SelBold = True
            ctlDetail.SelText = oMemberNames.Item(i)
            ctlDetail.SelBold = False
            ctlDetail.SelColor = RGB(0, 0, 0)
            ctlDetail.SelText = "(" & oMemberParamLists.Item(i) & ")"
            ctlDetail.SelColor = RGB(0, 0, 0)
            ctlDetail.SelText = vbCrLf
        End If
    Next i
    If nCount = 0 Then
        ctlDetail.SelText = "<none>" & vbCrLf
    End If
End Sub

Private Sub WordWrap()
    If ctlAspects.SelectedItem.Key = "Root" And sWrappedFileContentsRtf <> "" Then
        ctlDetail.TextRTF = sWrappedFileContentsRtf
    Else
        Dim lStart As Long, lEnd As Long, lPos As Long
        Dim sLine As String, sLeadingWhitespace As String, nLWLen As Integer
        Dim sChar As String
        lStart = 1
        Do
            lEnd = InStr(lStart + 1, ctlDetail.Text, vbCrLf)
            If lEnd = 0 Then Exit Do
            If lEnd - lStart > LineWidth Then
                sLine = Mid(ctlDetail.Text, lStart, lEnd - lStart)
                nLWLen = Len(sLine) - Len(LTrim(sLine))
                sLeadingWhitespace = Left(sLine, nLWLen)
                lEnd = lStart + LineWidth - 2
                
                'Find a previous word break, if any
                lPos = lEnd - 2
                Do
                    If lPos < lStart + nLWLen + 4 Then Exit Do 'Too far back
                    sChar = Mid(ctlDetail.Text, lPos, 1)
                    If sChar = " " Or sChar = "." Then
                        lEnd = lPos
                        Exit Do
                    End If
                    lPos = lPos - 1
                Loop
                'Insert indicators of line continuation
                ctlDetail.SelStart = lEnd - 1
                ctlDetail.SelText = " "
                ctlDetail.SelFontName = "Symbol"
                ctlDetail.SelColor = RGB(255, 0, 0)
                ctlDetail.SelBold = True
                ctlDetail.SelText = "®"
                ctlDetail.SelFontName = "Courier New"
                ctlDetail.SelText = vbCrLf
                lEnd = ctlDetail.SelStart + 1
                ctlDetail.SelText = sLeadingWhitespace
                ctlDetail.SelFontName = "Symbol"
                ctlDetail.SelText = "®"
                ctlDetail.SelFontName = "Courier New"
                ctlDetail.SelText = " "
            Else
                lStart = lEnd + 2
            End If
            Call SetProgress("Word-wrapping contents", 100 * ctlDetail.SelStart / Len(ctlDetail.Text))
        Loop
        If ctlAspects.SelectedItem.Key = "Root" Then
            sWrappedFileContentsRtf = ctlDetail.TextRTF
        End If
        ctlDetail.SelStart = 0
    End If
End Sub


'---- Private Event Handlers --------------------------------

Private Sub UserControl_Initialize()
    dProportionMiddle = ctlAspects.Width / UserControl.Width
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    If UserControl.Width < 200 Then
        UserControl.Width = 200
        Exit Sub
    ElseIf UserControl.Height < 200 Then
        UserControl.Height = 200
        Exit Sub
    End If
    ctlAspects.Height = UserControl.ScaleHeight - ctlAspects.Top
    ctlDetail.Height = UserControl.ScaleHeight - ctlDetail.Top + 1 * Screen.TwipsPerPixelY
    ctlAspects.Width = UserControl.ScaleWidth * dProportionMiddle
    ctlDetail.Left = ctlAspects.Width + 3 * Screen.TwipsPerPixelX
    ctlDetail.Width = UserControl.ScaleWidth - ctlDetail.Left - 2 * ctlAspects.Left
    ctlCopy.Left = ctlDetail.Left + ctlDetail.Width - ctlCopy.Width - 1 * Screen.TwipsPerPixelX
    ctlWrap.Left = ctlCopy.Left - ctlWrap.Width - 2 * ctlAspects.Left
    ctlDetailStatus.Left = ctlDetail.Left + Screen.TwipsPerPixelX
    ctlDetailStatus.Width = ctlWrap.Left - ctlDetail.Left - 4 * Screen.TwipsPerPixelX
    ctlProgressStatus.Top = ctlDetail.Top + ctlDetail.Height * 0.5 - ctlProgressStatus.Height - ctlAspects.Left
    ctlProgressStatus.Left = ctlDetail.Left + 2 * ctlAspects.Left
    ctlProgressStatus.Width = ctlDetail.Width - 2 * ctlAspects.Left
    ctlProgress.Top = ctlProgressStatus.Top + ctlProgressStatus.Height + 2 * ctlAspects.Left
    ctlProgress.Left = ctlProgressStatus.Left
    ctlProgress.Width = ctlProgressStatus.Width
    ctlSizerMiddle.Left = ctlAspects.Width + ctlAspects.Left
    ctlSizerMiddle.Height = UserControl.ScaleHeight
End Sub

Private Sub ctlSizerMiddle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    bSizingMiddle = True
End Sub

Private Sub ctlSizerMiddle_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bSizingMiddle = False
End Sub

Private Sub ctlSizerMiddle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bSizingMiddle Then
        dProportionMiddle = dProportionMiddle + (x / UserControl.Width)
        If dProportionMiddle < 0.1 Then dProportionMiddle = 0.1
        If dProportionMiddle > 0.9 Then dProportionMiddle = 0.9
        Call UserControl_Resize
    End If
End Sub

Private Sub ctlDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = Asc(vbTab) Then
        KeyCode = 0
        ctlAspects.SetFocus
    End If
End Sub

Private Sub ctlDetail_KeyUp(KeyCode As Integer, Shift As Integer)
    Call UpdateLineNumber
End Sub

Private Sub ctlDetail_Click()
    Call UpdateLineNumber
End Sub

Private Sub ctlAspects_NodeClick(ByVal Node As ComctlLib.Node)
    Dim oNode As Node, i As Integer, sViewing As String
    Dim sVisibility As String, sType As String, sKey As String
    'Set oNode = ctlAspects.SelectedItem
    Set oNode = Node
    If oNode.Key <> sCurrentlyViewing Then
        Call ResetDetail
        sCurrentlyViewing = oNode.Key
        sViewing = sCurrentlyViewing
        ctlDetail.Visible = False
        ctlDetailStatus.Caption = ""
        If sViewing = "Root" Then 'Display file contents
            Call ResetDetail
            ctlDetail.TextRTF = sFileContentsRtf
        Else 'Below root
            sViewing = Mid(sViewing, 6) 'Trim off "Root/"
            i = InStr(1, sViewing, "/")
            If i <> 0 Then
                sVisibility = Left(sViewing, i - 1)
                sViewing = Mid(sViewing, i + 1)
                i = InStr(1, sViewing, "/")
                If i <> 0 Then
                    sType = Left(sViewing, i - 1)
                    sViewing = Mid(sViewing, i + 1)
                    sKey = sType & "/" & sViewing
                    
                    If sType = "Property" Then
                        ctlDetail.SelColor = RGB(0, 0, 0)
                        If oMemberReadable.Item(sKey) And Not oMemberWritable.Item(sKey) Then
                            ctlDetail.SelText = "Read-Only "
                        ElseIf Not oMemberReadable.Item(sKey) And oMemberWritable.Item(sKey) Then
                            ctlDetail.SelText = "Write-Only "
                        Else 'Readable and writable
                            ctlDetail.SelText = "Read/Write "
                        End If
                    End If
                    
                    ctlDetail.SelColor = RGB(0, 0, 100)
                    ctlDetail.SelBold = True
                    ctlDetail.SelText = oMemberNames.Item(sKey)
                    ctlDetail.SelBold = False
                    ctlDetail.SelColor = RGB(0, 0, 0)
                    If Not sType = "Property" Then
                        ctlDetail.SelText = "(" & oMemberParamLists.Item(sKey) & ")"
                    End If
                    ctlDetail.SelText = " As "
                    ctlDetail.SelColor = RGB(100, 0, 100)
                    ctlDetail.SelText = oMemberReturnTypes.Item(sKey) & vbCrLf & vbCrLf
                    ctlDetail.SelColor = RGB(0, 100, 0)
                    ctlDetail.SelText = oMemberComments.Item(sKey)
                    
                Else
                    If sViewing = "Property" Then 'Properties
                        Call DetailProperties(sVisibility)
                    ElseIf sViewing = "Method" Then 'Methods
                        Call DetailMethods(sVisibility)
                    ElseIf sViewing = "Event" Then 'Events
                        Call DetailEvents(sVisibility)
                    End If
                End If
            
            Else
                sVisibility = sViewing
                ctlDetail.SelFontSize = 12
                ctlDetail.SelBold = True
                ctlDetail.SelColor = RGB(100, 0, 0)
                ctlDetail.SelText = sVisibility & " Properties" & vbCrLf
                ctlDetail.SelColor = RGB(0, 0, 0)
                ctlDetail.SelBold = False
                ctlDetail.SelFontSize = 10
                ctlDetail.SelText = vbCrLf
                Call DetailProperties(sVisibility)
                ctlDetail.SelText = vbCrLf
                ctlDetail.SelFontSize = 12
                ctlDetail.SelBold = True
                ctlDetail.SelColor = RGB(100, 0, 0)
                ctlDetail.SelText = "............................................................" & vbCrLf
                ctlDetail.SelText = sVisibility & " Methods" & vbCrLf
                ctlDetail.SelColor = RGB(0, 0, 0)
                ctlDetail.SelBold = False
                ctlDetail.SelFontSize = 10
                ctlDetail.SelText = vbCrLf
                Call DetailMethods(sVisibility)
                If sVisibility = "Public" Then
                    ctlDetail.SelText = vbCrLf
                    ctlDetail.SelFontSize = 12
                    ctlDetail.SelBold = True
                    ctlDetail.SelColor = RGB(100, 0, 0)
                    ctlDetail.SelText = "............................................................" & vbCrLf
                    ctlDetail.SelText = sVisibility & " Events" & vbCrLf
                    ctlDetail.SelColor = RGB(0, 0, 0)
                    ctlDetail.SelBold = False
                    ctlDetail.SelFontSize = 10
                    ctlDetail.SelText = vbCrLf
                    Call DetailEvents(sVisibility)
                End If
            End If
        End If
        ctlDetail.SelStart = 0
        ctlDetail.Visible = True
        Call UpdateLineNumber
    End If
    If ctlWrap = 1 Then
        Call ShowProgress("Word-wrapping contents")
        Call WordWrap
        Call HideProgress
    End If
End Sub

Private Sub ctlCopy_Click()
    On Error Resume Next
    Call Clipboard.Clear
    Call Clipboard.SetText(ctlDetail, vbCFRTF)
    If Err Then
        MsgBox Err.Description, vbCritical, "Error"
    End If
    ctlAspects.SetFocus
End Sub

Private Sub ctlWrap_Click()
    ctlDetail.Visible = False
    If ctlWrap = 1 Then
        Call ShowProgress("Word-wrapping contents")
        Call WordWrap
        Call HideProgress
    Else
        sCurrentlyViewing = ""
        Call ctlAspects_NodeClick(ctlAspects.SelectedItem)
    End If
    ctlDetail.Visible = True
    ctlDetail.SetFocus
End Sub


