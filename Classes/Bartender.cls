VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bartender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_inputProvider  As FuncOfString
Private m_outputProvider As ActionOfString

Friend Sub New_(aInputProvider As FuncOfString, aOutputProvider As ActionOfString)
    Set m_inputProvider = aInputProvider: Set m_outputProvider = aOutputProvider
End Sub

Public Function AskForDrink() As Boolean
    AskForDrink = True
    m_outputProvider.Invoke "What drink do you want? (beer, juice), or exit"
    
    Dim drink As String: drink = m_inputProvider.Invoke()
    
    If drink = "beer" Then
        
        m_outputProvider.Invoke "Not so fast cowboy. How old are you?"
        Dim sAge As String: sAge = m_inputProvider.Invoke()
        Dim age As Long
        If Not Lng_TryParse(sAge, age) Then
            
            m_outputProvider.Invoke "could not parse the age provided"
            
        Else
            
            If age >= 18 Then
                
                m_outputProvider.Invoke "Here you go! Cold beer."
                
            Else
                
                m_outputProvider.Invoke "Sorry but you're not old enough to drink beer (in the UK)"
                
            End If
            
        End If
        
    ElseIf drink = "juice" Then
        
        m_outputProvider.Invoke "Here you go! Fresh and nice juice."
    
    ElseIf drink = "exit" Then
        
        AskForDrink = False
        
    Else
        
        m_outputProvider.Invoke "Sorry mate but we don't do " & drink
        
    End If
    '
    'this class is violating open close principle
    '
End Function
