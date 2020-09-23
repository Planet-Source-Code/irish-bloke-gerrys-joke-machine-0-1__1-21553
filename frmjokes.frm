VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Gerry mc donnells Random Joke Machine"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "frmjokes.frx":0000
      Top             =   120
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gimme a joke"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   6360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VERY SIMPLE VB JOKE FILE MACHINE BY GERRY MC DONNELL 2001
'HAS NO ERROR HADLING , JSUT BASIC FILE INPUT FEATURES
'FROM THE FILE JOKES.JOK PUT THIS IN THE SAME
'DIR AS UR APP
'http://go.to/wdwtbam
Dim temp As String
Private Sub Command1_Click()
Text1.Text = ""
    temp = Input(1, #1)
Do While temp <> ")"
If temp = Chr(13) Or temp = Chr(10) Then
    Text1.Text = Text1.Text & Chr(13)
End If
    Text1.Text = Text1.Text & temp
    temp = Input(1, #1)
Loop
End Sub

Private Sub Form_Load()
''LAOD JOKE FILE
On Error GoTo err
curpath = App.Path
Open curpath & "\jokes.jok" For Input As #1





err:
If err.Number <> 0 Then
    MsgBox "Joke file not found.", vbCritical
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
MsgBox "thank you for using gerrys jokes visit http://go.to/wdwtbam for more cool stuff.", vbInformation, "Gerrys Jokes"
End
End Sub
