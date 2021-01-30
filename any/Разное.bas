Attribute VB_Name = "Разное"
Option Explicit

Function ЭтоЧисло(Строка As String) As Boolean
    Dim i, Длина As Integer
    
    If Строка = "" Then
        ЭтоЧисло = False
        Exit Function
    End If
    
    Длина = Len(Строка)
    For i = 1 To Длина
        If Mid(Строка, i, 1) < "0" Or Mid(Строка, i, 1) > "9" Then
            ЭтоЧисло = False
            Exit Function
        End If
    Next
    
    ЭтоЧисло = True
End Function

Function МаксимумИзЦелых(Число1, Число2 As Integer) As Integer
    If Число1 > Число2 Then
        МаксимумИзЦелых = Число1
    Else
        МаксимумИзЦелых = Число2
    End If
End Function

