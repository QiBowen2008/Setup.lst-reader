VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Setup.Lst�޸���"
   ClientHeight    =   5910
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1560
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   360
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "setup.lst"
      Filter          =   ".lst��"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LST�ļ�Ŀ¼"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��װ�������"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CAB�ļ���"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   840
   End
   Begin VB.Menu File 
      Caption         =   "�ļ�"
      Begin VB.Menu Open 
         Caption         =   "��"
      End
      Begin VB.Menu Save 
         Caption         =   "����"
      End
      Begin VB.Menu Saveas 
         Caption         =   "���Ϊ"
      End
      Begin VB.Menu New 
         Caption         =   "�½�"
      End
   End
   Begin VB.Menu OP 
      Caption         =   "ѡ��"
      Begin VB.Menu Set 
         Caption         =   "����"
      End
      Begin VB.Menu About 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long '������ȡϵͳ����API
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" _
           (ByVal lpApplicationName As Long, _
            ByVal lpKeyName As Long, _
            ByVal lpDefault As Long, _
            ByVal lpReturnedString As Long, _
            ByVal nSize As Long, _
            ByVal lpFileName As Long) As Long
Dim lstname As String
Private Sub Command1_Click()
    CommonDialog1.ShowOpen
    lstname = CommonDialog1.FileName
End Sub
'��INI�ļ�
Public Function GetValueFromINIFile(ByVal SectionName As String, _
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
    GetValueFromINIFile = strBuf
End Function
