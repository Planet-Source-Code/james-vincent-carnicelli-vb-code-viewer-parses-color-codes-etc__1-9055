VERSION 5.00
Begin VB.Form frmCodeView 
   Caption         =   "Code Module"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VbCodeBrowser.ctlCodeView ctlCode 
      Height          =   2085
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   4845
      _ExtentX        =   5583
      _ExtentY        =   4075
   End
End
Attribute VB_Name = "frmCodeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AllowClose As Boolean

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
    ctlCode.Width = Me.ScaleWidth - 2 * ctlCode.Left
    ctlCode.Height = Me.ScaleHeight - 2 * ctlCode.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If AllowClose Then
        'To do
    Else
        Cancel = 1
        Me.Hide
    End If
End Sub
