Attribute VB_Name = "MNew"
Option Explicit

Public Function Action(aObj As Object, ByVal aActionName As String) As Action
    Set Action = New Action: Action.New_ aObj, aActionName
End Function

Public Function ActionOfString(aObj As Object, ByVal aActionName As String) As ActionOfString
    Set ActionOfString = New ActionOfString: ActionOfString.New_ aObj, aActionName
End Function

Public Function FuncOfString(aObj As Object, ByVal aFuncName As String) As FuncOfString
    Set FuncOfString = New FuncOfString: FuncOfString.New_ aObj, aFuncName
End Function

Public Function TBConsole(aTBInput As TextBox, ATBOutput As TextBox) As TBConsole
    Set TBConsole = New TBConsole: TBConsole.New_ aTBInput, ATBOutput
End Function

Public Function Bartender(aInputProvider As FuncOfString, aOutputProvider As ActionOfString) As Bartender
    Set Bartender = New Bartender: Bartender.New_ aInputProvider, aOutputProvider
End Function
