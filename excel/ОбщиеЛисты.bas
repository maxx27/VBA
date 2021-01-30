Attribute VB_Name = "ОбщиеЛисты"
Option Explicit

Function ЯчейкаДоступнаДляЗаписи(Ячейка As Range) As Boolean
    МожноИспользоватьЯчейку = Ячейка.Value <> "" And _
        Ячейка.Locked = False And _
        CStr(Ячейка.Formula) = CStr(Ячейка.Value)
End Function

Function ЯчейкаДоступнаДляЧтения(Ячейка As Range) As Boolean
    МожноИспользоватьЯчейку = Ячейка.Value <> ""
End Function

Function НайтиЛист(НазваниеЛиста) As Excel.Worksheet
    Dim Лист As Excel.Worksheet

    If НазваниеЛиста = "" Then
        Set НайтиЛист = ActiveSheet
    Else
        On Error GoTo ЛистНеНайден
        Set НайтиЛист = ThisWorkbook.Sheets(НазваниеЛиста)
        On Error GoTo 0
    End If
    Exit Function

ЛистНеНайден:
    Set НайтиЛист = Nothing
    Resume Next
End Function

Function НайтиОбластьНаЛисте(НазваниеЛиста As String, Optional Диапазон As String = "") As Range
    Dim Лист As Excel.Worksheet
    
    ' Возвращаемое значение по умолчанию
    Set НайтиОбластьНаЛисте = Nothing
    
    ' Найти лист
    Set Лист = НайтиЛист(НазваниеЛиста)
    If Лист Is Nothing Then Exit Function
    
    ' Определить область поиска
    If Диапазон = "" Then
        Set НайтиОбластьНаЛисте = Лист.Cells
    ElseIf RegExpTest(Диапазон, "^\d+(\:\d)?$") Then
        Set НайтиОбластьНаЛисте = Лист.Rows(Диапазон)
    ElseIf RegExpTest(Диапазон, "^[A-Za-z]+(\:[A-Za-z]+)?$") Then
        Set НайтиОбластьНаЛисте = Лист.Columns(Диапазон)
    Else
        Set НайтиОбластьНаЛисте = Лист.Range(Диапазон)
    End If
End Function

Function НайтиЯчейкиНаЛисте(НазваниеЛиста, Содержимое As String, _
    Optional Диапазон As String = "", _
    Optional УчитыватьРегистр As Boolean = False) As Range
    '
    ' Назначение:
    ' Позволяет найти на листе [НазваниеЛиста] ячейки, которые содержат
    ' значение [Содержимое].
    ' Если [НазваниеЛиста] не указано, то поиск осуществляется на текущем листе.
    ' Если [Диапазон] не указан, то поиск осуществляется по всем ячейкам листа.
    ' Возвращает диапазон ячеек с заданным содержимым.
    
    Dim Лист As Excel.Worksheet
    Dim ОбластьПоиска, РезультатПоиска As Range
    Dim i As Integer
    Dim ПервыйАдрес, АдресаЯчеек As String
    
    ' Возвращаемое значение по умолчанию
    Set НайтиЯчейкиНаЛисте = Nothing
    
    ' Найти лист
    Set Лист = НайтиЛист(НазваниеЛиста)
    If Лист Is Nothing Then Exit Function
    
    ' Определить область поиска
    If Диапазон = "" Then
        Set ОбластьПоиска = Лист.Cells
    ElseIf RegExpTest(Диапазон, "^\d+(\:\d)?$") Then
        Set ОбластьПоиска = Лист.Rows(Диапазон)
    ElseIf RegExpTest(Диапазон, "^[A-Za-z]+(\:[A-Za-z]+)?$") Then
        Set ОбластьПоиска = Лист.Columns(Диапазон)
    Else
        Set ОбластьПоиска = Лист.Range(Диапазон)
    End If
    
    ' Выполнить поиск назад (иначе нельзя найти значение в A1)
    Set РезультатПоиска = ОбластьПоиска.Find(What:=Содержимое, _
        LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=УчитыватьРегистр, SearchFormat:=False)
    
    ' Найти адреса всех ячеек
    ПервыйАдрес = ""
    Do While Not РезультатПоиска Is Nothing
        If ПервыйАдрес = "" Then
            ПервыйАдрес = РезультатПоиска.Address
        ElseIf РезультатПоиска.Address = ПервыйАдрес Then
            Exit Do
        End If
            
        АдресаЯчеек = АдресаЯчеек & "," & РезультатПоиска.Address
        Set РезультатПоиска = ОбластьПоиска.FindNext(РезультатПоиска)
    Loop
    
    ' Получить все ячейки согласно условию поиска (игнорируем лидирующую запятую)
    If АдресаЯчеек <> "" Then
        Set НайтиЯчейкиНаЛисте = Лист.Range(Mid(АдресаЯчеек, 2))
    End If
End Function

Function НайтиСтолбецНаЛисте(НазваниеЛиста, НазваниеСтолбца As String, _
    Optional Диапазон As String = "", _
    Optional УчитыватьРегистр As Boolean = False) As Collection
    '
    ' Назначение:
    ' Позволяет найти на листе [НазваниеЛиста] ячейки, которые содержат
    ' значение [Содержимое].
    ' Если [НазваниеЛиста] не указано, то поиск осуществляется на текущем листе.
    ' Если [Диапазон] не указан, то поиск осуществляется по всем ячейкам листа.
    ' Возвращает коллекцию номеров столбцов в буквенной нотации.
    
    Dim Ячейка, РезультатПоиска As Range
    Dim Результат As New Collection
    
    Set РезультатПоиска = НайтиЯчейкиНаЛисте(НазваниеЛиста, НазваниеСтолбца, Диапазон, УчитыватьРегистр)
    If Not РезультатПоиска Is Nothing Then
        For Each Ячейка In РезультатПоиска
           Результат.Add (Mid(Ячейка.Address, 2, 1))
        Next
    End If
    
    Set НайтиСтолбецНаЛисте = СортироватьКоллекцию(УникальныеИзКоллекции(Результат))
End Function

Function НайтиСтрокуПоЗначениям(НазваниеЛиста As String, _
    ИменаСтолбцов As Collection, ЗначенияСтолбцов As Collection, _
    Optional ТолькоПервое As Boolean = False) As Collection
    
    Dim Лист As Excel.Worksheet
    Dim i, n, nMax, СтрокаПреподавателя, СтрокаСтавки As Integer
    Dim Совпадает As Boolean
    Dim Найденное As New Collection
    
    Set Лист = НайтиЛист(НазваниеЛиста)
    
    If Лист Is Nothing Then
        НайтиСтрокуПоЗначениям = "! Лист не найден !"
        Exit Function
    End If

    If ИменаСтолбцов.Count <> ЗначенияСтолбцов.Count Then
        НайтиСтрокуПоЗначениям = "! Аргументы не согласованы !"
        Exit Function
    End If

    ' Найти номер последней строки
    nMax = 0
    For i = 1 To ИменаСтолбцов.Count
        n = Лист.Columns(ИменаСтолбцов.Item(i)).SpecialCells(xlLastCell).Row
        If nMax < n Then nMax = n
    Next
    
    ' Найти строку
    For n = 1 To nMax
        
        ' Сравнить по всем значениям
        Совпадает = True
        For i = 1 To ИменаСтолбцов.Count
            If Лист.Range(ИменаСтолбцов.Item(i) & n).Value <> ЗначенияСтолбцов.Item(i) Then
                Совпадает = False
                Exit For
            End If
        Next
        
        ' Обработать найденное
        If Совпадает Then
            Найденное.Add (n)
            If ТолькоПервое Then Exit For
        End If
    Next
    
    Set НайтиПервуюСтрокуПоЗначениям = Найденное
End Function

Function ПолучитьСтолбецЯчейки(Ячейка As Range) As String
    ' Может возникнуть ошибка, если подсунуть строку "$1:$1"
    ПолучитьСтолбецЯчейки = RegExpFind(Ячейка.Address, "[a-zA-Z]+").Item(1)
End Function

Function ПолучитьСтрокуЯчейки(Ячейка As Range) As String
    ' Может возникнуть ошибка, если подсунуть столбец "$A:$A"
    ПолучитьСтрокуЯчейки = RegExpFind(Ячейка.Address, "[0-9]+").Item(1)
End Function

