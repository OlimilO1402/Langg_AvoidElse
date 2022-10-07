VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1: Test FuncOfString & ActionOfString"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton BtnForm2Show 
      Caption         =   ">> Form2: TestTBConsole"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   0
      Width           =   2295
   End
   Begin VB.TextBox TxtOutput 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   1080
      Width           =   8535
   End
   Begin VB.TextBox TxtInput 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Input:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'just for testing class Func and Action:
Private mTestActi  As ActionOfString
Private mTestFunc  As FuncOfString
Private mTBConsole As TBConsole

Private Sub Command1_Click()
    Dim aNewAction As Action: Set aNewAction = MNew.Action(Me, "Doo")
    aNewAction
End Sub

Private Sub Form_Activate()
    'TxtInput.SetFocus
End Sub

Private Sub Form_Load()
    
    Set mTestActi = MNew.ActionOfString(Me, "DoIt")
    Set mTestFunc = MNew.FuncOfString(Me, "Gimme")
    
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = TxtInput.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = TxtInput.Height
    
    If W > 0 And H > 0 Then TxtInput.Move L, T, W, H
    T = TxtOutput.Top
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then TxtOutput.Move L, T, W, H
End Sub

Public Sub Doo()
    MsgBox "Dedododo dedadada is all i want to say to you"
End Sub

Public Function Gimme() As String
    Gimme = TxtInput.Text
    TxtInput.Text = ""
End Function

Public Sub DoIt(s As String)
    TxtOutput.Text = TxtOutput.Text & s & vbCrLf
    TxtOutput.SelStart = Len(TxtOutput.Text)
End Sub

Private Sub TxtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyReturn Then
        Dim s As String
        s = mTestFunc.Invoke
        mTestActi.Invoke s
    End If
End Sub
Private Sub TxtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = KeyCodeConstants.vbKeyReturn Then
        KeyAscii = 0 'JIPPIAIYEAH!
    End If
End Sub

Private Sub BtnForm2Show_Click()
    Form2.Show
End Sub

