VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Haizim Password Changer"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CHANGE"
      Height          =   735
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "PASSWORD"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "USERNAME"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Private Sub Command1_Click()
If Text1.Text = "" And Text2.Text = "" Then
MsgBox "Complete all first", , "Haizim Password Changer"
Else
Shell "net user " & Text1.Text & " " & Text2.Text
MsgBox "Password has been changed", , "Haizim Password Changer"
End If
End Sub

Private Sub Form_Load()
ShellExecute 0, "runas", App.Path & "\" & App.EXEName & ".exe", Command & "/admin", vbNullString, SW_SHOWNORMAL
Unload Me
End Sub
