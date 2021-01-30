Attribute VB_Name = "РегулярныеВыражения"
Option Explicit

' Модуль содержит функции общего назначения, которые можно
' использовать для упрощения кода в листах Excel и документах Word.
' Требуется подключить Microsoft VBScript Regular Expression 1.0
' (есть версия 5.5, но в чём отличие я не знаю)
' P.S. Версия 5.5 менее распространена чем 1.0

' Функции для работы с регулярными выражениями взяты и доработаны с
' http://www.tmehta.com/regexp/add_code.htm

' Проверяет соответствие строки WhatMatch
Function RegExpTest(TestIn As String, TestWhat As String, Optional IgnoreCase As Boolean = True) As Boolean
    Dim RE As RegExp

    Set RE = New RegExp
    RE.Pattern = TestWhat
    RE.Global = False
    RE.IgnoreCase = IgnoreCase
    RegExpTest = RE.test(TestIn)
End Function

' Производит замену ReplaceWhat на ReplaceWith в ReplaceIn.
Function RegExpSubstitute(ReplaceIn As String, _
    ReplaceWhat As String, _
    ReplaceWith As String, _
    Optional IgnoreCase As Boolean = True, _
    Optional Globally As Boolean = True) As String
    Dim RE As Regexp

    Set RE = New RegExp
    RE.Pattern = ReplaceWhat
    RE.Global = Globally
    RE.IgnoreCase = IgnoreCase
    RegExpSubstitute = RE.Replace(ReplaceIn, ReplaceWith)
End Function

' Производит глобальный поиск по тексту.
' Возвращает коллекцию строк совпадения

Function RegExpFind(FindIn As String, FindWhat As String, Optional IgnoreCase As Boolean = False) As Collection
    Dim RE As RegExp, allMatches As MatchCollection, aMatch As Match
    Dim aResults As Collection

    Set aResults = New Collection
    Set RE = New RegExp
    RE.Pattern = FindWhat
    RE.IgnoreCase = IgnoreCase
    RE.Global = True

    Set allMatches = RE.Execute(FindIn)
    For Each aMatch In allMatches
        aResults.Add (aMatch.Value)
    Next
    Set RegExpFind = aResults
End Function

' Производит поиск и возвращает найденное согласно заданным подвыражениям
' (элементы заключенные в скобки).
' Пример: RegExpMatch("123 asd 4567 asd 89", "(\d)(\d+)") = ['1', '23', '4', '567', '8', '9']

Function RegExpMatch(FindIn As String, FindWhat As String, Optional IgnoreCase As Boolean = False) As Collection
    Dim RE As RegExp, aMatch As Match
    Dim aResults As New Collection
    Dim i As Integer
    Dim sName, sFragment As String

    Set RE = New RegExp
    RE.Pattern = FindWhat
    RE.IgnoreCase = IgnoreCase
    RE.Global = True

    For Each aMatch In RE.Execute(FindIn)
        If RE.test(aMatch.Value) Then
            i = 1
            GoTo MatchNext
            Do
                aResults.Add (sFragment)
                i = i + 1
MatchNext:
                sName = "$" & i
                sFragment = RE.Replace(aMatch.Value, sName)
            Loop While sFragment <> sName
        End If
    Next
    Set RegExpMatch = aResults
End Function
