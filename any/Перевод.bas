Attribute VB_Name = "�������"
Option Explicit

Public Function �������������(ByVal s As String, Optional ByVal IsFemale As Boolean = False) As String
    Dim i, ����� As Integer
    Dim ������� As Boolean
    Dim p1, p2, p3, ����������� As String
    
    If Not ��������(s) Then
        ������������� = ""
        Exit Function
    End If
    
    ����� = Len(s)
    ������� = True
    For i = 1 To �����
        If Mid(s, i, 1) <> "0" Then ������� = False
    Next
    If ������� Then
        ������������� = "����"
        Exit Function
    End If
    
    Select Case �����
        Case 1, 2, 3
            p1 = s
            p2 = ""
            p3 = ""
            ����������� = ��������������(s, IsFemale)
                 
        Case 4, 5, 6
            p1 = Right(s, 3)
            p2 = Left(s, Len(s) - 3)
            p3 = ""
            ����������� = ��������������(p2, True) & " " & ������(p2) & " " & ��������������(p1, IsFemale)
        
        Case 7, 8, 9
            p1 = Right(s, 3)
            p2 = Mid(s, Len(s) - 5, 3)
            p3 = Left(s, Len(s) - 6)
            ����������� = ��������������(p3, False) & " " & �������(p3) & " " & ��������������(p2, True) & " " & ������(p2) & " " & ��������������(p1, IsFemale)
    End Select
    
    ������������� = Trim(�����������)
End Function

Private Function ��������������(ByVal s As String, ByVal IsFemale As Boolean) As String
    Dim l, m, p, ���������� As String
    
    If Len(s) = 1 Then s = "00" & s
    If Len(s) = 2 Then s = "0" & s
    
    p = Right(s, 1)
    m = Mid(s, 2, 1)
    l = Left(s, 1)
    
    Select Case l
        Case "0"
            ���������� = ""
        Case "1"
            ���������� = "���"
        Case "2"
            ���������� = "������"
         Case "3"
            ���������� = "������"
        Case "4"
            ���������� = "���������"
        Case "5"
            ���������� = "�������"
         Case "6"
            ���������� = "��������"
        Case "7"
            ���������� = "�������"
        Case "8"
            ���������� = "���������"
        Case "9"
            ���������� = "���������"
    End Select
    
    Select Case m
        Case "1"
                    
            Select Case p
                Case "0"
                    ���������� = ���������� & " " & "������"
                Case "1"
                    ���������� = ���������� & " " & "�����������"
                Case "2"
                    ���������� = ���������� & " " & "����������"
                Case "3"
                    ���������� = ���������� & " " & "����������"
                Case "4"
                    ���������� = ���������� & " " & "������������"
                Case "5"
                    ���������� = ���������� & " " & "����������"
                Case "6"
                    ���������� = ���������� & " " & "�����������"
                Case "7"
                    ���������� = ���������� & " " & "����������"
                Case "8"
                    ���������� = ���������� & " " & "������������"
                Case "9"
                    ���������� = ���������� & " " & "������������"
            End Select
                    
        Case "2"
            ���������� = ���������� & " " & "��������"
        Case "3"
            ���������� = ���������� & " " & "��������"
        Case "4"
            ���������� = ���������� & " " & "�����"
        Case "5"
            ���������� = ���������� & " " & "���������"
        Case "6"
            ���������� = ���������� & " " & "����������"
        Case "7"
            ���������� = ���������� & " " & "���������"
        Case "8"
            ���������� = ���������� & " " & "�����������"
        Case "9"
            ���������� = ���������� & " " & "���������"
    End Select
    
    If m <> 1 Then
        Select Case p
            Case "1"
                If IsFemale = True Then
                    ���������� = ���������� & " " & "����"
                Else
                    ���������� = ���������� & " " & "����"
                End If
            Case "2"
                If IsFemale = True Then
                    ���������� = ���������� & " " & "���"
                Else
                    ���������� = ���������� & " " & "���"
                End If
             Case "3"
                ���������� = ���������� & " " & "���"
            Case "4"
                ���������� = ���������� & " " & "������"
            Case "5"
                ���������� = ���������� & " " & "����"
             Case "6"
                ���������� = ���������� & " " & "�����"
            Case "7"
                ���������� = ���������� & " " & "����"
            Case "8"
                ���������� = ���������� & " " & "������"
            Case "9"
                ���������� = ���������� & " " & "������"
        End Select
    End If
               
    �������������� = Trim(����������)
End Function

Private Function ������(ByVal ���������� As String) As String
    Dim �������� As String
    
    If ���������� = "000" Then
        ������ = ""
        Exit Function
    End If
        
    Select Case Len(����������)
        Case 1
            If Right(����������, 1) = "1" Then
                �������� = "������"
            ElseIf Right(����������, 1) = "2" Or Right(����������, 1) = "3" Or Right(����������, 1) = "4" Then
                �������� = "������"
            Else
                �������� = "�����"
            End If
        
        Case 2
            If Left(����������, 1) = "1" Then
                �������� = "�����"
            ElseIf Right(����������, 1) = "2" Or Right(����������, 1) = "3" Or Right(����������, 1) = "4" Then
                �������� = "������"
            Else
                �������� = "�����"
            End If
        
        Case 3
            If Mid(����������, 2, 1) = "1" Then
                �������� = "�����"
            ElseIf Right(����������, 1) = "2" Or Right(����������, 1) = "3" Or Right(����������, 1) = "4" Then
                �������� = "������"
            Else
                �������� = "�����"
            End If
    End Select

    ������ = ��������
End Function

Private Function �������(ByVal ���������� As String) As String
    Dim �������� As String
    
    If ���������� = "000" Then
        ������� = ""
        Exit Function
    End If
    
    Select Case Len(����������)
        Case 1
            If Right(����������, 1) = "1" Then
              �������� = "�������"
            ElseIf Right(����������, 1) = "2" Or Right(����������, 1) = "3" Or Right(����������, 1) = "4" Then
                �������� = "��������"
            Else
                �������� = "���������"
            End If
            
        Case 2
            If Left(����������, 1) = "1" Then
              �������� = "���������"
            ElseIf Right(����������, 1) = "2" Or Right(����������, 1) = "3" Or Right(����������, 1) = "4" Then
                �������� = "��������"
            Else
                �������� = "���������"
            End If
            
        Case 3
            If Mid(����������, 2, 1) = "1" Then
              �������� = "���������"
            ElseIf Right(����������, 1) = "2" Or Right(����������, 1) = "3" Or Right(����������, 1) = "4" Then
                �������� = "��������"
            Else
                �������� = "���������"
            End If
    End Select
    
    ������� = ��������
End Function
