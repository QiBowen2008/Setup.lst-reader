Attribute VB_Name = "Module1"
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long '������ȡϵͳ����API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" _
           (ByVal lpApplicationName As Long, _
            ByVal lpKeyName As Long, _
            ByVal lpDefault As Long, _
            ByVal lpReturnedString As Long, _
            ByVal nSize As Long, _
            ByVal lpFileName As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Dim lstname As String
'��INI�ļ�
Public Function GetINI(ByVal SectionName As String, _
    ByVal KeyName As String, _
    ByVal IniFileName As String) As String

    Dim strBuf As String
    '128���ַ�����ʼ��ʱ�� 0 ���
    strBuf = String(128, 0)

    GetPrivateProfileString StrPtr(SectionName), _
        StrPtr(KeyName), _
        StrPtr(""), _
        StrPtr(strBuf), _
        128, _
        StrPtr(IniFileName)
    'ȥ������� 0
    strBuf = Replace(strBuf, Chr(0), "")
    GetINI = strBuf
End Function
'ע��д�����ȡ��[������]��[������]һ��Ҫ��ͬ��
Private Function WriteINI(txtym1 As String, txtym2 As String, txtym3 As String) As String
    Dim ULR As String
    ULR = App.Path & "\config.ini" 'INI�ļ�·��
    Dim txtBuff As String '��ȡ
        '�����ȡ�ַ����ĳ��ȣ���Space" ȡ ʵ �� �� �� ȥ �� �� �� �� �� �� �� �� �� �� ��
        txtBuff = Space(1000)
        't x t B u f f = S p a c e "ȡʵ���ַ�ȥ���ַ��������Ŀո� txtBuff = Space"ȡʵ���ַ�ȥ���ַ��������Ŀո�
        '��ȡINI�ļ�(������,������,��,��ȡ������ֵ,��ȡ�ַ�������,·��)
        Call GetPrivateProfileString(txtym1, txtym2, "", txtBuff, Len(txtBuff), ULR)
        '��ʾʵ���ַ�����ȡ"txtBuff"��ߵ��ַ���(ȡ�õ��ַ���,�ַ����ܳ���(ȥ���ַ����ұ߶���Ŀո��ַ�(ȡ�õ��ַ���))�ó��ַ���ʵ�ʳ��ȶ�һ������˼�1)
        txtBuff = Left(txtBuff, Len(RTrim(txtBuff)) - 1)
        '�Ѷ�ȡ�����ַ������ݵ�"STRYMINI"����
        WriteINI = txtBuff
    End If
End Function
