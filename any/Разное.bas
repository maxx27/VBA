Attribute VB_Name = "������"
Option Explicit

Function ��������(������ As String) As Boolean
    Dim i, ����� As Integer
    
    If ������ = "" Then
        �������� = False
        Exit Function
    End If
    
    ����� = Len(������)
    For i = 1 To �����
        If Mid(������, i, 1) < "0" Or Mid(������, i, 1) > "9" Then
            �������� = False
            Exit Function
        End If
    Next
    
    �������� = True
End Function

Function ���������������(�����1, �����2 As Integer) As Integer
    If �����1 > �����2 Then
        ��������������� = �����1
    Else
        ��������������� = �����2
    End If
End Function

