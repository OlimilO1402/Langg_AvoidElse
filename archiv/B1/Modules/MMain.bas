Attribute VB_Name = "MMain"
Option Explicit
'Nick Chapsas
'Why I don't use the "else" keyword in my code anymore
'https://www.youtube.com/watch?v=_ougvb8mT7k

'Object Calisthenics
'By Jeff Bay in The ThoughtWorks Anthology

'1. Only One Level Of Indentation Per Method

'2. Don't Use the ELSE Keyword

'3. Wrap All Primitives And Strings
'4. First Class Collections
'5. One Dot Per Line
'6. Don't Abbreviate
'7. Keep All Entities Small
'8. No Classes With More Than Two Instance Variables
'9. No Getters/Setters/Properties

'zu 2
'It's not that you should remove it
'It's more that you don't have to use it in the first place

Public Console As New Console

Private mBarTender As Bartender

Sub Main()
    Console.IsInIDE = False
    Set mBarTender = MNew.Bartender(MNew.FuncOfString(Console, "ReadLine"), MNew.ActionOfString(Console, "WriteLine"))
    
    While mBarTender.AskForDrink
        'DoEvents
    Wend
    
End Sub

Public Function Lng_TryParse(s As String, Lng_out As Long) As Boolean
Try: On Error GoTo Catch
    If IsNumeric(s) Then
        Lng_out = CLng(s)
        Lng_TryParse = True
    End If
    Exit Function
Catch:
End Function

