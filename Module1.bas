Attribute VB_Name = "Module1"
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long '声明读取系统语言API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" _
           (ByVal lpApplicationName As Long, _
            ByVal lpKeyName As Long, _
            ByVal lpDefault As Long, _
            ByVal lpReturnedString As Long, _
            ByVal nSize As Long, _
            ByVal lpFileName As Long) As Long
'声明写INI文件API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public lstname As String, appname As String
'读INI文件函数
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
'写ini文件函数
Private Function WriteINI(txtym1 As String, txtym2 As String, txtym3 As String) As String
    lstname = App.Path & "\config.ini" 'INI文件路径
    Dim txtBuff As String '读取
    Call WritePrivateProfileString(txtym1, txtym2, txtym3, ULR)
End Function
