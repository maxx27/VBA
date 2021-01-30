Attribute VB_Name = "����������"
Option Explicit

Function �����������������������(������ As Range) As Boolean
    ����������������������� = ������.Value <> "" And _
        ������.Locked = False And _
        CStr(������.Formula) = CStr(������.Value)
End Function

Function �����������������������(������ As Range) As Boolean
    ����������������������� = ������.Value <> ""
End Function

Function ���������(�������������) As Excel.Worksheet
    Dim ���� As Excel.Worksheet

    If ������������� = "" Then
        Set ��������� = ActiveSheet
    Else
        On Error GoTo ������������
        Set ��������� = ThisWorkbook.Sheets(�������������)
        On Error GoTo 0
    End If
    Exit Function

������������:
    Set ��������� = Nothing
    Resume Next
End Function

Function �������������������(������������� As String, Optional �������� As String = "") As Range
    Dim ���� As Excel.Worksheet
    
    ' ������������ �������� �� ���������
    Set ������������������� = Nothing
    
    ' ����� ����
    Set ���� = ���������(�������������)
    If ���� Is Nothing Then Exit Function
    
    ' ���������� ������� ������
    If �������� = "" Then
        Set ������������������� = ����.Cells
    ElseIf RegExpTest(��������, "^\d+(\:\d)?$") Then
        Set ������������������� = ����.Rows(��������)
    ElseIf RegExpTest(��������, "^[A-Za-z]+(\:[A-Za-z]+)?$") Then
        Set ������������������� = ����.Columns(��������)
    Else
        Set ������������������� = ����.Range(��������)
    End If
End Function

Function ������������������(�������������, ���������� As String, _
    Optional �������� As String = "", _
    Optional ���������������� As Boolean = False) As Range
    '
    ' ����������:
    ' ��������� ����� �� ����� [�������������] ������, ������� ��������
    ' �������� [����������].
    ' ���� [�������������] �� �������, �� ����� �������������� �� ������� �����.
    ' ���� [��������] �� ������, �� ����� �������������� �� ���� ������� �����.
    ' ���������� �������� ����� � �������� ����������.
    
    Dim ���� As Excel.Worksheet
    Dim �������������, ��������������� As Range
    Dim i As Integer
    Dim �����������, ����������� As String
    
    ' ������������ �������� �� ���������
    Set ������������������ = Nothing
    
    ' ����� ����
    Set ���� = ���������(�������������)
    If ���� Is Nothing Then Exit Function
    
    ' ���������� ������� ������
    If �������� = "" Then
        Set ������������� = ����.Cells
    ElseIf RegExpTest(��������, "^\d+(\:\d)?$") Then
        Set ������������� = ����.Rows(��������)
    ElseIf RegExpTest(��������, "^[A-Za-z]+(\:[A-Za-z]+)?$") Then
        Set ������������� = ����.Columns(��������)
    Else
        Set ������������� = ����.Range(��������)
    End If
    
    ' ��������� ����� ����� (����� ������ ����� �������� � A1)
    Set ��������������� = �������������.Find(What:=����������, _
        LookIn:=xlValues, LookAt:=xlWhole, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=����������������, SearchFormat:=False)
    
    ' ����� ������ ���� �����
    ����������� = ""
    Do While Not ��������������� Is Nothing
        If ����������� = "" Then
            ����������� = ���������������.Address
        ElseIf ���������������.Address = ����������� Then
            Exit Do
        End If
            
        ����������� = ����������� & "," & ���������������.Address
        Set ��������������� = �������������.FindNext(���������������)
    Loop
    
    ' �������� ��� ������ �������� ������� ������ (���������� ���������� �������)
    If ����������� <> "" Then
        Set ������������������ = ����.Range(Mid(�����������, 2))
    End If
End Function

Function �������������������(�������������, ��������������� As String, _
    Optional �������� As String = "", _
    Optional ���������������� As Boolean = False) As Collection
    '
    ' ����������:
    ' ��������� ����� �� ����� [�������������] ������, ������� ��������
    ' �������� [����������].
    ' ���� [�������������] �� �������, �� ����� �������������� �� ������� �����.
    ' ���� [��������] �� ������, �� ����� �������������� �� ���� ������� �����.
    ' ���������� ��������� ������� �������� � ��������� �������.
    
    Dim ������, ��������������� As Range
    Dim ��������� As New Collection
    
    Set ��������������� = ������������������(�������������, ���������������, ��������, ����������������)
    If Not ��������������� Is Nothing Then
        For Each ������ In ���������������
           ���������.Add (Mid(������.Address, 2, 1))
        Next
    End If
    
    Set ������������������� = ��������������������(���������������������(���������))
End Function

Function ����������������������(������������� As String, _
    ������������� As Collection, ���������������� As Collection, _
    Optional ������������ As Boolean = False) As Collection
    
    Dim ���� As Excel.Worksheet
    Dim i, n, nMax, �������������������, ������������ As Integer
    Dim ��������� As Boolean
    Dim ��������� As New Collection
    
    Set ���� = ���������(�������������)
    
    If ���� Is Nothing Then
        ���������������������� = "! ���� �� ������ !"
        Exit Function
    End If

    If �������������.Count <> ����������������.Count Then
        ���������������������� = "! ��������� �� ����������� !"
        Exit Function
    End If

    ' ����� ����� ��������� ������
    nMax = 0
    For i = 1 To �������������.Count
        n = ����.Columns(�������������.Item(i)).SpecialCells(xlLastCell).Row
        If nMax < n Then nMax = n
    Next
    
    ' ����� ������
    For n = 1 To nMax
        
        ' �������� �� ���� ���������
        ��������� = True
        For i = 1 To �������������.Count
            If ����.Range(�������������.Item(i) & n).Value <> ����������������.Item(i) Then
                ��������� = False
                Exit For
            End If
        Next
        
        ' ���������� ���������
        If ��������� Then
            ���������.Add (n)
            If ������������ Then Exit For
        End If
    Next
    
    Set ���������������������������� = ���������
End Function

Function ���������������������(������ As Range) As String
    ' ����� ���������� ������, ���� ��������� ������ "$1:$1"
    ��������������������� = RegExpFind(������.Address, "[a-zA-Z]+").Item(1)
End Function

Function ��������������������(������ As Range) As String
    ' ����� ���������� ������, ���� ��������� ������� "$A:$A"
    �������������������� = RegExpFind(������.Address, "[0-9]+").Item(1)
End Function

