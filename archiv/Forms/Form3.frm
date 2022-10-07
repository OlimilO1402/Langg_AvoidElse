VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670
   LinkTopic       =   "Form3"
   ScaleHeight     =   5295
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnForm3Show 
      Caption         =   ">> Form4: Test Avoid ELSE"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton BtnStartBartender 
      Caption         =   "Start Bartender"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
   Begin VB.TextBox TxtOutput 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
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
      TabIndex        =   3
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      TabIndex        =   2
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mTBConsole As TBConsole
Private mBarTender As Bartender1

Private Sub BtnForm3Show_Click()
    Form4.Show
End Sub

Private Sub Form_Load()
    
    Set mTBConsole = MNew.TBConsole(Me.TxtInput, Me.TxtOutput)
    
    Set mBarTender = MNew.Bartender1(MNew.FuncOfString(mTBConsole, "ReadLine"), MNew.ActionOfString(mTBConsole, "WriteLine"))
    
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

Private Sub BtnStartBartender_Click()
    While mBarTender.AskForDrink
        'mBarTender.AskForDrink
    Wend
End Sub

