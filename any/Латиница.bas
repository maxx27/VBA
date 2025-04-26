Sub ЗаменитьБуквы()
    ЗаменитьФрагмент "а", "a"
    ЗаменитьФрагмент "А", "A"
'    ЗаменитьФрагмент "к", "k"
    ЗаменитьФрагмент "К", "K"
    ЗаменитьФрагмент "М", "M"
    ЗаменитьФрагмент "е", "e"
    ЗаменитьФрагмент "Е", "E"
    ЗаменитьФрагмент "о", "o"
    ЗаменитьФрагмент "О", "O"
    ЗаменитьФрагмент "х", "x"
    ЗаменитьФрагмент "Х", "X"
    ЗаменитьФрагмент "с", "c"
    ЗаменитьФрагмент "С", "C"
    ЗаменитьФрагмент "Т", "T"
    ЗаменитьФрагмент "Н", "H"
    ЗаменитьФрагмент "В", "B"
End Sub

Sub ЗаменитьФрагмент(ByVal Старый As String, ByVal Новый As String)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = Старый
        .Replacement.Text = Новый
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
