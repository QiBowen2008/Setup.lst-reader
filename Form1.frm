VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Setup.Lst�޸���"
   ClientHeight    =   4170
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
   ScaleHeight     =   4170
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   1560
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
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
      TabIndex        =   3
      Text            =   "Combo1"
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
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "setup.lst"
      Filter          =   ".lst��"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ĭ�ϰ�װĿ¼"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   1080
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
Private Sub Command1_Click()
    CommonDialog1.ShowOpen
    lstname = CommonDialog1.FileName
    Combo1.Text = lstname
    Combo1.AddItem lstname
End Sub


