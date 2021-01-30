Attribute VB_Name = "Ref"

'
' Добавить ссылки
'

Sub AddCrossrefToHeading()
    ' http://windowssecrets.com/forums/showthread.php/119370-Word-2003-VBA-tool-to-find-the-targets-of-cross-references-and-insert-them
    ' http://my.safaribooksonline.com/book/office-and-productivity-applications/0596004931/editing-power-tools/wordhks-chp-4-sect-18
        
    ' Пропустить разделители в начале и в конце выделения
    Selection.MoveStartWhile " " & vbCr, Selection.Characters.Count
    Selection.MoveEndWhile " " & vbCr, -Selection.Characters.Count
    
    ' Выход, если selection включает несколько абзацев
    If Selection.Range.Paragraphs.Count <> 1 Then Exit Sub

    ' Искомый текст
    Dim text As String
    text = LCase(Selection.Range.text)

    Dim i As Integer
    i = 1
    Dim v As Variant
    For Each v In ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)
        If LCase(Trim(v)) = text Then
            ' Вставить маркер
            Selection.InsertAfter " - "
            Selection.Collapse wdCollapseEnd
            
            ' Вставить ссылку (разные варианты берутся из типа WdReferenceKind)
            ' Вставить текст заголовока
            Selection.InsertCrossReference wdRefTypeHeading, wdContentText, i, True
            ' Вставить страницу
            Selection.InsertCrossReference wdRefTypeHeading, wdPageNumber, i, True
            ' Вставить 'выше/ниже'
            Selection.InsertCrossReference wdRefTypeHeading, wdPosition, i, True
            
            Exit Sub
        End If
        i = i + 1
    Next
    
    MsgBox "Can't find a heading with text '" & Selection.Range.text & "'."
End Sub

Sub AddHyperlinkToSelection()
    AddLink True
End Sub

Sub AddCrossrefToSelection()
    AddLink False
End Sub

Sub AddLink(asHyper As Boolean)
    ' http://my.safaribooksonline.com/book/office-and-productivity-applications/0596004931/editing-power-tools/wordhks-chp-4-sect-18
   
    ' Пропустить разделители в начале и в конце выделения
    Selection.MoveStartWhile " " & vbCr, Selection.Characters.Count
    Selection.MoveEndWhile " " & vbCr, -Selection.Characters.Count
    
    ' Выход, если selection включает несколько абзацев
    If Selection.Range.Paragraphs.Count <> 1 Then Exit Sub
    
    Dim paraIndex As Long
    paraIndex = GetParagraphIndex(Selection.Range.Paragraphs.First)
    
    ' Искомый текст
    Dim text As String
    text = Selection.Range.text
    Dim textLcase As String
    textLcase = LCase(text)

    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        ' Пропустить разделители в начале и в конце выделения абзаца
        Dim r As Range
        Set r = para.Range
        r.MoveStartWhile " " & vbCr, r.Characters.Count
        r.MoveEndWhile " " & vbCr, -r.Characters.Count
        If LCase(r.text) = textLcase Then
            If GetParagraphIndex(para) = paraIndex Then
                MsgBox "Can't self reference!"
                Exit Sub
            End If

            ' Вставить маркер
            Selection.InsertAfter " "
            Selection.Collapse wdCollapseEnd
            
            ' Вставить ссылку
            If asHyper Then
                ' 1) как гиперссылку
                ActiveDocument.Hyperlinks.Add Selection.Range, "", GetOrSetCrossrefBookmark(para), "", text
            Else
                ' 2) как перекрёстную ссылку
                ' HERE: Имена параметров обязательны!
                Selection.InsertCrossReference _
                    ReferenceKind:=wdContentText, _
                    ReferenceItem:=GetOrSetCrossrefBookmark(para), _
                    ReferenceType:=wdRefTypeBookmark, _
                    InsertAsHyperlink:=True
            End If
            
            Exit Sub
        End If
    Next
    
    MsgBox "Can't find a paragraph with text '" & text & "'."
End Sub

Function GetParagraphIndex(para As Paragraph) As Long
    GetParagraphIndex = para.Range.Document.Range(0, para.Range.End).Paragraphs.Count
End Function

Function GetOrSetCrossrefBookmark(para As Paragraph) As Bookmark
    Dim BookmarkPrefix As String
    BookmarkPrefix = "XREF_"

    If para.Range.bookmarks.Count <> 0 Then
        Dim i As Long
        For i = 1 To para.Range.bookmarks.Count
            ' проверяем по префиксу, чтобы среди всех закладок на этот абзац была именно нами созданная
            If InStr(1, para.Range.bookmarks(i).Name, BookmarkPrefix) Then
                Set GetOrSetCrossrefBookmark = para.Range.bookmarks(i)
                Exit Function
            End If
        Next
    End If
    
    ' пропустить с конца символ перевода строки
    Dim rng As Range
    Set rng = para.Range
    rng.MoveEnd wdCharacter, -1

    ' закладки не было, создать имя для новой
    Dim sBookmarkName As String
    sBookmarkName = BookmarkPrefix & ConvertStringRefBookmarkName(rng.text)

    ' проверить на повтор имени закладки (могла быть создана ранее)
    Dim iSuffix As Integer, sSuffix As String
    iSuffix = 0
    sSuffix = ""
    Do While para.Range.Document.bookmarks.Exists(sBookmarkName + sSuffix)
        iSuffix = iSuffix + 1
        sSuffix = "_" & CStr(iSuffix)
    Loop
    sBookmarkName = sBookmarkName + sSuffix
    
    Set GetOrSetCrossrefBookmark = para.Range.Document.bookmarks.Add(sBookmarkName, rng)
End Function

Function ConvertStringRefBookmarkName(ByVal str As String) As String
    str = RemoveInvalidBookmarkCharsFromString(str)
    str = Replace(str, " ", "_")
    ConvertStringRefBookmarkName = str
End Function

Function RemoveInvalidBookmarkCharsFromString(ByVal str As String) As String
    Dim i As Integer
    For i = 0 To 255
        Select Case i
            ' оставляем символы "192 To 255" (они соответствуют кириллице)
            Case 0 To 31, 33 To 47, 58 To 64, 91 To 96, 123 To 191
                str = Replace(str, Chr(i), vbNullString)
        End Select
    Next
    RemoveInvalidBookmarkCharsFromString = str
End Function

'
' Переименовать закладки
'

Sub RenameBookmarks()
    ' http://superuser.com/questions/359066/batch-rename-multiple-bookmarks-in-word-docx-file
    ' http://social.technet.microsoft.com/Forums/office/en-US/7a24d7c0-5960-409b-ae35-be9f99ebfeea/word-2007-vba-updating-only-specific-cross-references
    ' https://groups.google.com/forum/#!topic/microsoft.public.word.vba.general/1nW_TNO3gnw
    ' http://support.microsoft.com/kb/247507
    
    ' На самом деле переименовать закладку нельзя!
    ' Можно создать другую с тем же диапазоном и удалить старую.
    ' Перед удалением обновить все ссылки со старой закладки на новую.
    ' Удаление и добавление меняет элементы в конце коллекции
    ' Поэтому перебор элементов коллекции следует делать в обратном порядке.
    Dim i As Long
    For i = ActiveDocument.bookmarks.Count To 1 Step -1
        ' Переименовать каждую закладку
        Dim BookmarkOld As Bookmark
        Set BookmarkOld = ActiveDocument.bookmarks(i)
        Dim BookmarkNew As Bookmark
        Set BookmarkNew = BookmarkOld.Range.bookmarks.Add("NEW_" & BookmarkOld.Name, BookmarkOld.Range)
        UpdateHyperlinks BookmarkOld, BookmarkNew
        UpdateCrossrefs BookmarkOld, BookmarkNew
        BookmarkOld.Delete
    Next
End Sub

Sub UpdateHyperlinks(BookmarkOld As Bookmark, BookmarkNew As Bookmark)
    Dim i As Long
    For i = ActiveDocument.Hyperlinks.Count To 1 Step -1
    'For Each hLink In ActiveDocument.Hyperlinks
        Dim hLink As Hyperlink
        Set hLink = ActiveDocument.Hyperlinks(i)
        If hLink.SubAddress = BookmarkOld Then
            ' нельзя изменить HyperLink: hLink.SubAddress = BookmarkNew
            ' но можно удалить и создать новую
            Dim r As Range
            Set r = hLink.Range.Duplicate
            hLink.Delete
            ActiveDocument.Hyperlinks.Add r, "", BookmarkNew
        End If
    Next
End Sub

Sub UpdateCrossrefs(BookmarkOld As Bookmark, BookmarkNew As Bookmark)
    Dim f As Field
    For Each f In ActiveDocument.Fields
        If f.Type = wdFieldRef Or f.Type = wdFieldPageRef Then
            ' возможное содержимое ссылки {REF _REF12344 \h }
            Dim s As String
            s = LTrim(Replace(f.Code.text, "REF", ""))
            ' выделить первое слово до пробела - это ссылка
            Dim n As Long
            n = InStr(s, " ") - 1
            If n = -1 Then n = Len(s)
            Dim ref As String
            ref = Left(s, n)
            s = Replace(s, ref, "")
            ' Можно использовать дополнительное условие ActiveDocument.bookmarks.Exists(s)
            ' s хранит оставшиеся параметры ссылки (например, \h для гиперссылки и т.д.)
            f.Code.text = "REF " & BookmarkNew & s
            f.Update
        End If
    Next
End Sub

Sub RemoveAllBookmarksInSelection()
    ' Удаляет все закладки в выделенном фрагменте, но не ссылки на них

    Dim b As Bookmark
    For Each b In Selection.Range.bookmarks
        b.Delete
    Next
End Sub
