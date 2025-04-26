Attribute VB_Name = "DOS"
Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" _
   Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
   ByVal lpszShortPath As String, ByVal cchBuffer As Long) _
   As Long

Public Function GetShortFileName(ByVal FullPath As String) As String

    ' Взято с http://www.freevbcode.com/ShowCode.Asp?ID=506
    ' Назначение: получить имя DOS (формат 8.3) для заданного длинного имени [FullPath]
    ' Возвращает имя в формате 8.3 или "" в случае ошибки (такой файл не существует или другая причина)
    ' Пример: Debug.Print GetShortFileName("C:\My Documents\My Very Long File Name.doc") возвращает
    ' если файл существует, то в окне отладчика отобразиться "C:\MYDOCU~1\MYVERY~1.DOC"

    Dim lAns As Long
    Dim sAns As String
    Dim iLen As Integer
   
    ' Функция не работает, если файл не существует
    If Dir(FullPath) = "" Then Exit Function

    sAns = Space(255)
    lAns = GetShortPathName(FullPath, sAns, 255)
    GetShortFileName = Left(sAns, lAns)
End Function


