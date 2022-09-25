Attribute VB_Name = "Module1"
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long '声明读取系统语言API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" _
           (ByVal lpApplicationName As Long, _
            ByVal lpKeyName As Long, _
            ByVal lpDefault As Long, _
            ByVal lpReturnedString As Long, _
            ByVal nSize As Long, _
            ByVal lpFileName As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Dim lstname As String
'读INI文件
Public Function GetINI(ByVal SectionName As String, _
    ByVal KeyName As String, _
    ByVal IniFileName As String) As String

    Dim strBuf As String
    '128个字符，初始化时用 0 填充
    strBuf = String(128, 0)

    GetPrivateProfileString StrPtr(SectionName), _
        StrPtr(KeyName), _
        StrPtr(""), _
        StrPtr(strBuf), _
        128, _
        StrPtr(IniFileName)
    '去除多余的 0
    strBuf = Replace(strBuf, Chr(0), "")
    GetINI = strBuf
End Function
'注：写入与读取的[主项名]和[子项名]一定要相同！
Private Function WriteINI(txtym1 As String, txtym2 As String, txtym3 As String) As String
    Dim ULR As String
    ULR = App.Path & "\config.ini" 'INI文件路径
    Dim txtBuff As String '读取
        '定义读取字符串的长度，“Space" 取 实 际 字 符 去 掉 字 符 后 面 多 余 的 空 格 。
        txtBuff = Space(1000)
        't x t B u f f = S p a c e "取实际字符去掉字符后面多余的空格。 txtBuff = Space"取实际字符去掉字符后面多余的空格。
        '读取INI文件(主项名,子项名,空,读取子项名值,读取字符串长度,路径)
        Call GetPrivateProfileString(txtym1, txtym2, "", txtBuff, Len(txtBuff), ULR)
        '显示实际字符串。取"txtBuff"左边的字符串(取得的字符串,字符串总长度(去掉字符串右边多余的空格字符(取得的字符串))得出字符串实际长度多一个，因此减1)
        txtBuff = Left(txtBuff, Len(RTrim(txtBuff)) - 1)
        '把读取到的字符串传递到"STRYMINI"函数
        WriteINI = txtBuff
    End If
End Function
