Attribute VB_Name = "ƒаты"
Option Explicit
Function Ёто¬исокосный√од(год As Integer) As Boolean
    Ёто¬исокосный√од = (год Mod 400) = 0 Or ((год Mod 4) = 0 And (год Mod 100) <> 0)
End Function
Function Ќазваниећес€ца¬„исло(мес€ц As String) As Integer
    Dim n As Integer
    Select Case мес€ц
    Case "сент€брь"
     n = 9
    Case "окт€брь"
     n = 10
    Case "но€брь"
     n = 11
    Case "декабрь"
     n = 12
    Case "€нварь"
     n = 1
    Case "февраль"
     n = 2
    Case "март"
     n = 3
    Case "апрель"
     n = 4
    Case "май"
     n = 5
    Case "июнь"
     n = 6
    Case "июль"
     n = 7
    Case "август"
     n = 8
    End Select
    Ќазваниећес€ца¬„исло = n
End Function

Function —колькоƒней¬ћес€це(мес€ц As String, год As Integer) As Integer
    Dim p As Integer
    Dim мес€ц„ислом As Integer
    мес€ц„ислом = Ќазваниећес€ца¬„исло(мес€ц)
    Select Case мес€ц„ислом
    Case 4, 6, 9, 11
     p = 30
    Case 1, 3, 5, 7, 8, 10, 12
     p = 31
    Case 2
     If Ёто¬исокосный√од(год) Then
     p = 29
     Else
     p = 28
     End If
    End Select
    —колькоƒней¬ћес€це = p
End Function


