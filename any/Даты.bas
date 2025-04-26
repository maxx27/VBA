Attribute VB_Name = "Даты"
Option Explicit
Function ЭтоВисокосныйГод(год As Integer) As Boolean
    ЭтоВисокосныйГод = (год Mod 400) = 0 Or ((год Mod 4) = 0 And (год Mod 100) <> 0)
End Function
Function НазваниеМесяцаВЧисло(месяц As String) As Integer
    Dim n As Integer
    Select Case месяц
    Case "сентябрь"
     n = 9
    Case "октябрь"
     n = 10
    Case "ноябрь"
     n = 11
    Case "декабрь"
     n = 12
    Case "январь"
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
    НазваниеМесяцаВЧисло = n
End Function

Function СколькоДнейВМесяце(месяц As String, год As Integer) As Integer
    Dim p As Integer
    Dim месяцЧислом As Integer
    месяцЧислом = НазваниеМесяцаВЧисло(месяц)
    Select Case месяцЧислом
    Case 4, 6, 9, 11
     p = 30
    Case 1, 3, 5, 7, 8, 10, 12
     p = 31
    Case 2
     If ЭтоВисокосныйГод(год) Then
     p = 29
     Else
     p = 28
     End If
    End Select
    СколькоДнейВМесяце = p
End Function


