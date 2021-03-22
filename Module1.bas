Attribute VB_Name = "Module1"
Sub Chat()
Dim userString As String
Dim nameString As String
Dim gradeString As String
userString = InputBox("請輸入學號")
MsgBox "你的學號為" & userString
nameString = InputBox("請輸入你名字")
MsgBox "你的名字為" & nameString
gradeString = InputBox("請輸入VBA成績")
MsgBox "你的分數為" & gradeString
End Sub

