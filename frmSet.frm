VERSION 5.00
Begin VB.Form frmSet 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'注：写入与读取的[主项名]和[子项名]一定要相同！
Private Function STRYMINI(txtym1 As String, txtym2 As String, txtym3 As String, ONOFF As Boolean) As String
    Dim ULR As String
    ULR = App.Path & "\config.ini" 'INI文件路径
    Dim txtBuff As String
    If ONOFF = True Then '读取
        '定义读取字符串的长度，“Space" 取 实 际 字 符 去 掉 字 符 后 面 多 余 的 空 格 。
        txtBuff = Space(1000)
        't x t B u f f = S p a c e "取实际字符去掉字符后面多余的空格。 txtBuff = Space"取实际字符去掉字符后面多余的空格。
        '读取INI文件(主项名,子项名,空,读取子项名值,读取字符串长度,路径)
        Call GetPrivateProfileString(txtym1, txtym2, "", txtBuff, Len(txtBuff), ULR)
        '显示实际字符串。取"txtBuff"左边的字符串(取得的字符串,字符串总长度(去掉字符串右边多余的空格字符(取得的字符串))得出字符串实际长度多一个，因此减1)
        txtBuff = Left(txtBuff, Len(RTrim(txtBuff)) - 1)
        '把读取到的字符串传递到"STRYMINI"函数
        STRYMINI = txtBuff
    Else
        '把字符串写入INI文件(主项名，子项名，值，保存INI文件的路径)
        Call WritePrivateProfileString(txtym1, txtym2, txtym3, ULR)
    End If
End Function
