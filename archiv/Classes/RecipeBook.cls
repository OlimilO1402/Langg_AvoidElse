VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecipeBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_recipes As Collection '(Of String, Action)

Private Sub Class_Initialize()
    Set m_recipes = New Collection
End Sub

Public Sub Add(key As String, Act As Action)
    m_recipes.Add Act, key
End Sub

Public Function Contains(key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_recipes(key)) Then: 'DoNothing
    Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Property Get Item(key As String) As Action
    Set Item = m_recipes.Item(key)
End Property
