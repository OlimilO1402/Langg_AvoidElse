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
Private m_recipeBook     As RecipeBook

Private Sub Class_Initialize()
    Set m_recipeBook = New RecipeBook
    m_recipeBook.Add "beer", MNew.Action(Me, "HandleBeerServing")
    m_recipeBook.Add "juice", MNew.Action(Me, "HandleJuiceServing")
End Sub

Friend Sub New_(aInputProvider As FuncOfString, aOutputProvider As ActionOfString)
    Set m_inputProvider = aInputProvider: Set m_outputProvider = aOutputProvider
End Sub

Public Function AskForDrink() As Boolean
    AskForDrink = True
    
    m_outputProvider.Invoke "What drink do you want? (beer, juice)"
    
    Dim drink As String: drink = m_inputProvider.Invoke()
    
    If m_recipeBook.Contains(drink) Then
        Dim ServeDrink As Action: Set ServeDrink = m_recipeBook.Item(drink)
        ServeDrink.Invoke
        Exit Function
    End If
    
    HandleUnavailableDrink drink
    
End Function

Public Sub HandleJuiceServing() 'HandleJuiceServing
    
    m_outputProvider.Invoke "Here you go! Fresh and nice juice."
    
End Sub

Public Sub HandleBeerServing() 'HandleBeerServing
    
    m_outputProvider.Invoke "Not so fast cowboy. How old are you?"
    Dim sAge As String: sAge = m_inputProvider.Invoke()
    Dim age As Long
    If Not Lng_TryParse(sAge, age) Then
        
        HandleInvalidAge
        Exit Sub
        
    End If
    
    HandleBeerAgeCheck age
    
End Sub

Public Sub HandleInvalidAge()
    
    m_outputProvider.Invoke "could not parse the age provided"
    
End Sub

Public Sub HandleBeerAgeCheck(ByVal age As Long)
    
    If age >= 18 Then
        
        m_outputProvider.Invoke "Here you go! Cold beer."
        Exit Sub
        
    End If
    
    m_outputProvider.Invoke "Sorry but you're not old enough to drink beer (in the UK)"
    
End Sub

Public Sub HandleUnavailableDrink(ByVal drink As String) 'HandleUnavailableDrink
    
    m_outputProvider.Invoke "Sorry mate but we don't do " & drink
    
End Sub
