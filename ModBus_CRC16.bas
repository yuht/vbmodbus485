Attribute VB_Name = "ModBus_CRC16"
Option Explicit

Public Sub CRC16(data() As Byte, CrcStartPos As Integer, CrcEndPos As Integer, CrcRet_1 As Byte, CrcRet_2 As Byte)   'CRC���㺯��
    Dim CRC16Lo As Byte, CRC16Hi As Byte   'CRC�Ĵ���
    Dim CL As Byte, CH As Byte            '����ʽ��&HA001
    Dim SaveHi As Byte, SaveLo As Byte
    Dim i As Integer
    Dim Flag As Integer
    CRC16Lo = &HFF
    CRC16Hi = &HFF
    CL = &H1
    CH = &HA0
    For i = CrcStartPos To CrcEndPos
        CRC16Lo = CRC16Lo Xor data(i) 'ÿһ��������CRC�Ĵ����������
        For Flag = 0 To 7
            SaveHi = CRC16Hi
            SaveLo = CRC16Lo
            CRC16Hi = CRC16Hi \ 2            '��λ����һλ
            CRC16Lo = CRC16Lo \ 2            '��λ����һλ
            If ((SaveHi And &H1) = &H1) Then '�����λ�ֽ����һλΪ1
                CRC16Lo = CRC16Lo Or &H80      '���λ�ֽ����ƺ�ǰ�油1
            End If                           '�����Զ���0
            If ((SaveLo And &H1) = &H1) Then '���LSBΪ1���������ʽ��������
                CRC16Hi = CRC16Hi Xor CH
                CRC16Lo = CRC16Lo Xor CL
            End If
        Next
    Next
    CrcRet_1 = CRC16Lo
    CrcRet_2 = CRC16Hi
End Sub
