VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "SIRI IS QUESTIONABLE - GOOGLE"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCloseProg 
      Caption         =   "BE SIRI AND CLOSE THE GODDAMN THING"
      Height          =   735
      Left            =   3240
      TabIndex        =   8
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CommandButton btnProg 
      Caption         =   "BE SIRI AND LAUNCH THE GODDAMN THING"
      Height          =   735
      Left            =   600
      TabIndex        =   5
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox txtProgQuery 
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   7455
   End
   Begin VB.TextBox txtQuery 
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   6495
   End
   Begin VB.CommandButton btnGoogle 
      Caption         =   "BE SIRI AND GOOGLE THE GODDAMN THING"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Couldn't be bothered to open a new project eh. "
      Height          =   615
      Index           =   1
      Left            =   6360
      TabIndex        =   7
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label lblQuery 
      BackStyle       =   0  'Transparent
      Caption         =   "Query"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Probably still buggy and so forth and ugly as well knock yourself out"
      Height          =   615
      Index           =   0
      Left            =   3360
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label lblQuery 
      BackStyle       =   0  'Transparent
      Caption         =   "Query"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim program As Integer

Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam _
    As Long) As Long


Private Sub btnCloseProg_Click()
PostMessage program, WM_QUIT, 0&, 0&
End Sub

Private Sub btnGoogle_Click()
query = txtQuery.Text
With CreateObject("Shell.Application")
    .ShellExecute "http://www.google.com.au/search?q=" & query
End With
End Sub

Private Sub btnProg_Click()
If Trim$(txtProgQuery.Text) <> "" Then
    progPath = Trim$(txtProgQuery.Text)
    If My.Computer.FileSystem.FileExists(progPath) Then
        program = Shell(progPath, vbNormalFocus)
    Else
        MsgBox "Nope. GTFO I has no program"
    End If
End If
End Sub
