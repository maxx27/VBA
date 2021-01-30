Attribute VB_Name = "DOS"
Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" _
   Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
   ByVal lpszShortPath As String, ByVal cchBuffer As Long) _
   As Long

Public Function GetShortFileName(ByVal FullPath As String) As String

    ' ����� � http://www.freevbcode.com/ShowCode.Asp?ID=506
    ' ����������: �������� ��� DOS (������ 8.3) ��� ��������� �������� ����� [FullPath]
    ' ���������� ��� � ������� 8.3 ��� "" � ������ ������ (����� ���� �� ���������� ��� ������ �������)
    ' ������: Debug.Print GetShortFileName("C:\My Documents\My Very Long File Name.doc") ����������
    ' ���� ���� ����������, �� � ���� ��������� ������������ "C:\MYDOCU~1\MYVERY~1.DOC"

    Dim lAns As Long
    Dim sAns As String
    Dim iLen As Integer
   
    ' ������� �� ��������, ���� ���� �� ����������
    If Dir(FullPath) = "" Then Exit Function

    sAns = Space(255)
    lAns = GetShortPathName(FullPath, sAns, 255)
    GetShortFileName = Left(sAns, lAns)
End Function


