Attribute VB_Name = "mdlCommon"
Option Explicit

Public Const C_MSG_001 = "��������ȷ����"
Public Const C_MSG_002 = "��ֹ���ڲ��ܴ�����ʼ����"
Public Const C_MSG_TITLE_ERR = "����"


'**************************************************************
'*�������������Ƿ�Ϊ�����������
'*����ֵ:
'*  TRUE: ����
'*  FALSE: ������
'**************************************************************
Public Function Ymd_chek_proc(a As String) As Boolean
Dim b As Integer
    Ymd_chek_proc = True
    If IsNumeric(a) Then
        If Len(Trim(a)) = 8 Then
            If Left(a, 4) >= "1900" Then
                If Mid(a, 5, 2) >= "01" And Mid(a, 5, 2) <= "12" Then
                    If Right(a, 2) >= "01" And Right(a, 2) <= "31" Then
                        If Not IsDate(Format(a, "####/##/##")) Then
                            b = 1
                        Else
                            b = 0
                        End If
                    Else
                        b = 1
                    End If
                Else
                    b = 1
                End If
            Else
                b = 1
            End If
        Else
            b = 1
        End If
    Else
        b = 1
    End If
    If b = 1 Then
        Ymd_chek_proc = False
    End If
End Function
