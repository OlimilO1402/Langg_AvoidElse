VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Action"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Obj As Object
Private m_ActionName As String

Friend Sub New_(aObj As Object, aActionName As String)
    Set m_Obj = aObj: m_ActionName = aActionName
End Sub

Public Sub Invoke()
Attribute Invoke.VB_UserMemId = 0
Attribute Invoke.VB_MemberFlags = "40"
    CallByName m_Obj, m_ActionName, VbCallType.VbMethod
End Sub

