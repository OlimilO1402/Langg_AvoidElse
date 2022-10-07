VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2: Test TBConsole"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8670
   LinkTopic       =   "Form2"
   ScaleHeight     =   5295
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton BtnForm3Show 
      Caption         =   ">> Form3: Test Bartender"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   0
      Width           =   2295
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
      TabIndex        =   1
      Top             =   360
      Width           =   8535
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
      TabIndex        =   0
      Top             =   1080
      Width           =   8535
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
      TabIndex        =   3
      Top             =   120
      Width           =   510
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
      TabIndex        =   2
      Top             =   840
      Width           =   645
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mTestActi  As ActionOfString
Private mTestFunc  As FuncOfString
Private mTBConsole As TBConsole

Private Sub Form_Load()
        
    Set mTBConsole = MNew.TBConsole(Me.TxtInput, Me.TxtOutput)
    
    Set mTestActi = MNew.ActionOfString(mTBConsole, "WriteLine")
    Set mTestFunc = MNew.FuncOfString(mTBConsole, "ReadLine")
    
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

Private Sub Command1_Click()
    
    mTestActi.Invoke "What do you want to drink? (beer, juice)"
    
    Dim drink As String: drink = mTestFunc.Invoke
    
    If Len(drink) Then
        
        If drink = "beer" Then
            
            mTestActi.Invoke "Nicht so schnell, wie alt bist du?"
            
            Dim age As String: age = mTestFunc.Invoke
            
            If CLng(age) >= 18 Then
                
                mTestActi.Invoke "OK here you go! Cold fresh " & drink
                
            Else
                
                mTestActi.Invoke "Sorry you are not old enough to drink beer!"
                
                Command1_Click
                
            End If
            
        ElseIf drink = "juice" Then
            
            mTestActi.Invoke "OK there you are geniesse deinen " & drink
            
            Command1_Click
        Else
            
            mTestActi.Invoke "Sorry we don't serve " & drink
            
        End If
        
    End If
    
End Sub

Private Sub BtnForm3Show_Click()
    Form3.Show
End Sub

