Attribute VB_Name = "���������"
Option Explicit


' ��������� ��������� �������� � ���������.

Function ����������������(������� As Variant, ��������� As Collection) As Boolean
    Dim ���������������� As Variant

    For Each ���������������� In ���������
        If ���������������� = ������� Then
            ���������������� = True
            Exit Function
        End If
    Next
    
    ���������������� = False
End Function

' �������� ������ ���������� �������� ���������.

Function ���������������������(��������� As Collection) As Collection
    Dim ������������������� As New Collection
    Dim ������� As Variant
    
    If ��������� Is Nothing Then
        Set ��������������������� = Nothing
        Exit Function
    End If
    
    For Each ������� In ���������
        If Not ����������������(�������, �������������������) Then �������������������.Add (�������)
    Next
    
    Set ��������������������� = �������������������
End Function

' ��������� �������� ��������� ������� ��������
' TODO: ����� �������� ��������� ����� ��������� (����� �������������� ����������� �� ����, � ����� ������ ����)

Function ��������������������(��������� As Collection) As Collection
    Dim ���������������� As Boolean
    Dim ��������� As Variant
    Dim i As Integer
    
    If ��������� Is Nothing Then
        Set �������������������� = Nothing
        Exit Function
    End If
    
    Do
        ���������������� = False
        For i = 2 To ���������.Count
            If ���������.Item(i - 1) > ���������.Item(i) Then
                ���������.Add ���������.Item(i), Before:=i - 1
                ���������.Remove i + 1
                ���������������� = True
            End If
        Next
    Loop While ����������������
    
    Set �������������������� = ���������
End Function


