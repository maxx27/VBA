Sub �������������()
    ���������������� "�", "a"
    ���������������� "�", "A"
'    ���������������� "�", "k"
    ���������������� "�", "K"
    ���������������� "�", "M"
    ���������������� "�", "e"
    ���������������� "�", "E"
    ���������������� "�", "o"
    ���������������� "�", "O"
    ���������������� "�", "x"
    ���������������� "�", "X"
    ���������������� "�", "c"
    ���������������� "�", "C"
    ���������������� "�", "T"
    ���������������� "�", "H"
    ���������������� "�", "B"
End Sub

Sub ����������������(ByVal ������ As String, ByVal ����� As String)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ������
        .Replacement.Text = �����
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
