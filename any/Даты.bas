Attribute VB_Name = "����"
Option Explicit
Function ����������������(��� As Integer) As Boolean
    ���������������� = (��� Mod 400) = 0 Or ((��� Mod 4) = 0 And (��� Mod 100) <> 0)
End Function
Function ��������������������(����� As String) As Integer
    Dim n As Integer
    Select Case �����
    Case "��������"
     n = 9
    Case "�������"
     n = 10
    Case "������"
     n = 11
    Case "�������"
     n = 12
    Case "������"
     n = 1
    Case "�������"
     n = 2
    Case "����"
     n = 3
    Case "������"
     n = 4
    Case "���"
     n = 5
    Case "����"
     n = 6
    Case "����"
     n = 7
    Case "������"
     n = 8
    End Select
    �������������������� = n
End Function

Function ������������������(����� As String, ��� As Integer) As Integer
    Dim p As Integer
    Dim ����������� As Integer
    ����������� = ��������������������(�����)
    Select Case �����������
    Case 4, 6, 9, 11
     p = 30
    Case 1, 3, 5, 7, 8, 10, 12
     p = 31
    Case 2
     If ����������������(���) Then
     p = 29
     Else
     p = 28
     End If
    End Select
    ������������������ = p
End Function


