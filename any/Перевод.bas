Attribute VB_Name = "Перевод"
Option Explicit

Public Function ЧислоПрописью(ByVal s As String, Optional ByVal IsFemale As Boolean = False) As String
    Dim i, Длина As Integer
    Dim ИзНулей As Boolean
    Dim p1, p2, p3, ПолноеЧисло As String
    
    If Not ЭтоЧисло(s) Then
        ЧислоПрописью = ""
        Exit Function
    End If
    
    Длина = Len(s)
    ИзНулей = True
    For i = 1 To Длина
        If Mid(s, i, 1) <> "0" Then ИзНулей = False
    Next
    If ИзНулей Then
        ЧислоПрописью = "ноль"
        Exit Function
    End If
    
    Select Case Длина
        Case 1, 2, 3
            p1 = s
            p2 = ""
            p3 = ""
            ПолноеЧисло = ГруппаПрописью(s, IsFemale)
                 
        Case 4, 5, 6
            p1 = Right(s, 3)
            p2 = Left(s, Len(s) - 3)
            p3 = ""
            ПолноеЧисло = ГруппаПрописью(p2, True) & " " & Тысяча(p2) & " " & ГруппаПрописью(p1, IsFemale)
        
        Case 7, 8, 9
            p1 = Right(s, 3)
            p2 = Mid(s, Len(s) - 5, 3)
            p3 = Left(s, Len(s) - 6)
            ПолноеЧисло = ГруппаПрописью(p3, False) & " " & Миллион(p3) & " " & ГруппаПрописью(p2, True) & " " & Тысяча(p2) & " " & ГруппаПрописью(p1, IsFemale)
    End Select
    
    ЧислоПрописью = Trim(ПолноеЧисло)
End Function

Private Function ГруппаПрописью(ByVal s As String, ByVal IsFemale As Boolean) As String
    Dim l, m, p, Количество As String
    
    If Len(s) = 1 Then s = "00" & s
    If Len(s) = 2 Then s = "0" & s
    
    p = Right(s, 1)
    m = Mid(s, 2, 1)
    l = Left(s, 1)
    
    Select Case l
        Case "0"
            Количество = ""
        Case "1"
            Количество = "сто"
        Case "2"
            Количество = "двести"
         Case "3"
            Количество = "триста"
        Case "4"
            Количество = "четыреста"
        Case "5"
            Количество = "пятьсот"
         Case "6"
            Количество = "шестьсот"
        Case "7"
            Количество = "семьсот"
        Case "8"
            Количество = "восемьсот"
        Case "9"
            Количество = "девятьсот"
    End Select
    
    Select Case m
        Case "1"
                    
            Select Case p
                Case "0"
                    Количество = Количество & " " & "десять"
                Case "1"
                    Количество = Количество & " " & "одиннадцать"
                Case "2"
                    Количество = Количество & " " & "двенадцать"
                Case "3"
                    Количество = Количество & " " & "тринадцать"
                Case "4"
                    Количество = Количество & " " & "четырнадцать"
                Case "5"
                    Количество = Количество & " " & "пятнадцать"
                Case "6"
                    Количество = Количество & " " & "шестнадцать"
                Case "7"
                    Количество = Количество & " " & "семнадцать"
                Case "8"
                    Количество = Количество & " " & "восемнадцать"
                Case "9"
                    Количество = Количество & " " & "девятнадцать"
            End Select
                    
        Case "2"
            Количество = Количество & " " & "двадцать"
        Case "3"
            Количество = Количество & " " & "тридцать"
        Case "4"
            Количество = Количество & " " & "сорок"
        Case "5"
            Количество = Количество & " " & "пятьдесят"
        Case "6"
            Количество = Количество & " " & "шестьдесят"
        Case "7"
            Количество = Количество & " " & "семьдесят"
        Case "8"
            Количество = Количество & " " & "восемьдесят"
        Case "9"
            Количество = Количество & " " & "девяносто"
    End Select
    
    If m <> 1 Then
        Select Case p
            Case "1"
                If IsFemale = True Then
                    Количество = Количество & " " & "одна"
                Else
                    Количество = Количество & " " & "один"
                End If
            Case "2"
                If IsFemale = True Then
                    Количество = Количество & " " & "две"
                Else
                    Количество = Количество & " " & "два"
                End If
             Case "3"
                Количество = Количество & " " & "три"
            Case "4"
                Количество = Количество & " " & "четыре"
            Case "5"
                Количество = Количество & " " & "пять"
             Case "6"
                Количество = Количество & " " & "шесть"
            Case "7"
                Количество = Количество & " " & "семь"
            Case "8"
                Количество = Количество & " " & "восемь"
            Case "9"
                Количество = Количество & " " & "девять"
        End Select
    End If
               
    ГруппаПрописью = Trim(Количество)
End Function

Private Function Тысяча(ByVal ГруппаЦифр As String) As String
    Dim Прописью As String
    
    If ГруппаЦифр = "000" Then
        Тысяча = ""
        Exit Function
    End If
        
    Select Case Len(ГруппаЦифр)
        Case 1
            If Right(ГруппаЦифр, 1) = "1" Then
                Прописью = "тысяча"
            ElseIf Right(ГруппаЦифр, 1) = "2" Or Right(ГруппаЦифр, 1) = "3" Or Right(ГруппаЦифр, 1) = "4" Then
                Прописью = "тысячи"
            Else
                Прописью = "тысяч"
            End If
        
        Case 2
            If Left(ГруппаЦифр, 1) = "1" Then
                Прописью = "тысяч"
            ElseIf Right(ГруппаЦифр, 1) = "2" Or Right(ГруппаЦифр, 1) = "3" Or Right(ГруппаЦифр, 1) = "4" Then
                Прописью = "тысячи"
            Else
                Прописью = "тысяч"
            End If
        
        Case 3
            If Mid(ГруппаЦифр, 2, 1) = "1" Then
                Прописью = "тысяч"
            ElseIf Right(ГруппаЦифр, 1) = "2" Or Right(ГруппаЦифр, 1) = "3" Or Right(ГруппаЦифр, 1) = "4" Then
                Прописью = "тысячи"
            Else
                Прописью = "тысяч"
            End If
    End Select

    Тысяча = Прописью
End Function

Private Function Миллион(ByVal ГруппаЦифр As String) As String
    Dim Прописью As String
    
    If ГруппаЦифр = "000" Then
        Миллион = ""
        Exit Function
    End If
    
    Select Case Len(ГруппаЦифр)
        Case 1
            If Right(ГруппаЦифр, 1) = "1" Then
              Прописью = "миллион"
            ElseIf Right(ГруппаЦифр, 1) = "2" Or Right(ГруппаЦифр, 1) = "3" Or Right(ГруппаЦифр, 1) = "4" Then
                Прописью = "миллиона"
            Else
                Прописью = "миллионов"
            End If
            
        Case 2
            If Left(ГруппаЦифр, 1) = "1" Then
              Прописью = "миллионов"
            ElseIf Right(ГруппаЦифр, 1) = "2" Or Right(ГруппаЦифр, 1) = "3" Or Right(ГруппаЦифр, 1) = "4" Then
                Прописью = "миллиона"
            Else
                Прописью = "миллионов"
            End If
            
        Case 3
            If Mid(ГруппаЦифр, 2, 1) = "1" Then
              Прописью = "миллионов"
            ElseIf Right(ГруппаЦифр, 1) = "2" Or Right(ГруппаЦифр, 1) = "3" Or Right(ГруппаЦифр, 1) = "4" Then
                Прописью = "миллиона"
            Else
                Прописью = "миллионов"
            End If
    End Select
    
    Миллион = Прописью
End Function
