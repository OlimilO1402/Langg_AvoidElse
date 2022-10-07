VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncOfString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Obj As Object
Private m_FuncName As String

Friend Sub New_(aObj As Object, ByVal aFuncName As String)
    Set m_Obj = aObj: m_FuncName = aFuncName
End Sub

Public Function Invoke() As String
    Invoke = CallByName(m_Obj, m_FuncName, VbCallType.VbMethod)
End Function
