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
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'ע��д�����ȡ��[������]��[������]һ��Ҫ��ͬ��
Private Function STRYMINI(txtym1 As String, txtym2 As String, txtym3 As String, ONOFF As Boolean) As String
    Dim ULR As String
    ULR = App.Path & "\config.ini" 'INI�ļ�·��
    Dim txtBuff As String
    If ONOFF = True Then '��ȡ
        '�����ȡ�ַ����ĳ��ȣ���Space" ȡ ʵ �� �� �� ȥ �� �� �� �� �� �� �� �� �� �� ��
        txtBuff = Space(1000)
        't x t B u f f = S p a c e "ȡʵ���ַ�ȥ���ַ��������Ŀո� txtBuff = Space"ȡʵ���ַ�ȥ���ַ��������Ŀո�
        '��ȡINI�ļ�(������,������,��,��ȡ������ֵ,��ȡ�ַ�������,·��)
        Call GetPrivateProfileString(txtym1, txtym2, "", txtBuff, Len(txtBuff), ULR)
        '��ʾʵ���ַ�����ȡ"txtBuff"��ߵ��ַ���(ȡ�õ��ַ���,�ַ����ܳ���(ȥ���ַ����ұ߶���Ŀո��ַ�(ȡ�õ��ַ���))�ó��ַ���ʵ�ʳ��ȶ�һ������˼�1)
        txtBuff = Left(txtBuff, Len(RTrim(txtBuff)) - 1)
        '�Ѷ�ȡ�����ַ������ݵ�"STRYMINI"����
        STRYMINI = txtBuff
    Else
        '���ַ���д��INI�ļ�(����������������ֵ������INI�ļ���·��)
        Call WritePrivateProfileString(txtym1, txtym2, txtym3, ULR)
    End If
End Function
