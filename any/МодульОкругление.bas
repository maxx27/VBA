Attribute VB_Name = "����������������"
Option Explicit

Function ���������(ByVal ����� As Double, Optional ByVal ��������������� As Integer = 0) As Double
    ��������� = Format(Abs(�����), "0." & Replace(Space(���������������), " ", "0")) * Sgn(�����)
End Function

Function ��������������(ByVal ����� As Double, Optional ByVal ��������������� As Integer = 0) As Double
    Dim s As Double
    s = CDbl("0," & Replace(Space(���������������), " ", "0") & "5")
    If ����� = ���������(�����, ���������������) Then
        �������������� = �����
    Else
        �������������� = ���������(����� + s, ���������������)
    End If
End Function

Function �������������(ByVal ����� As Double, Optional ByVal ��������������� As Integer = 0) As Double
    Dim s As Double
    s = CDbl("0," & Replace(Space(���������������), " ", "0") & "5")
    ������������� = ���������(Abs(�����) - s, ���������������) * Sgn(�����)
End Function

