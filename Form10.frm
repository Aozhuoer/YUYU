VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�����������Ĺ���"
   ClientHeight    =   8625
   ClientLeft      =   7725
   ClientTop       =   1260
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   10200
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   1455
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "Form10.frx":74CA
      Top             =   6120
      Width           =   5775
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "ʹ�ý̳�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   1695
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Form10.frx":74D0
      Top             =   3240
      Width           =   5775
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "һ������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   1335
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form10.frx":74D6
      Top             =   600
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "һ������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "����adbʵ�ֵ�Ӧ���Ƽ������ṩ�������ӣ��������κ��ƹ㣬���κ���������ѯ���ǵ�ԭ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   8280
      Width           =   7545
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ظó���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   15
      Top             =   7680
      Width           =   1260
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ֻ�������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   14
      Top             =   5640
      Width           =   1200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   10080
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ֻ�Ͷ��������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   11
      Top             =   2760
      Width           =   1680
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ظó���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   10
      Top             =   5040
      Width           =   1260
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "github"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   5520
      TabIndex        =   9
      Top             =   5040
      Width           =   675
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   300
      Left            =   6480
      TabIndex        =   8
      Top             =   5040
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   300
      Left            =   5520
      TabIndex        =   5
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ظ�Ӧ�ã�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ֻ�����״̬��ͼ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   10080
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Image Image3 
      Height          =   2415
      Left            =   120
      Picture         =   "Form10.frx":74DC
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   3780
   End
   Begin VB.Image Image2 
      Height          =   2415
      Left            =   120
      Picture         =   "Form10.frx":C373
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   3780
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   120
      Picture         =   "Form10.frx":14EDF
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3780
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ȥ�ᰲ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006DC725&
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   7680
      Width           =   915
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Caption = "����ִ��"
Command1.Enabled = 0
Shell "cmd /c adb shell pm grant com.ksxkq.floating android.permission.WRITE_SECURE_SETTINGS"
Savetime = Timer
While Timer < Savetime + 2
DoEvents
Wend
Shell "cmd /c adb shell settings put global enable_freeform_support 1"
MsgBox "������ɣ��������������ֻ�"
Command1.Enabled = 1
Command1.Caption = "����"
End Sub

Private Sub Command2_Click()
Command2.Caption = "����ִ��"
Command2.Enabled = 0
Shell "cmd /c adb shell pm grant com.zacharee1.systemuituner android.permission.WRITE_SECURE_SETTINGS"
Shell "cmd /c adb shell pm grant com.zacharee1.systemuituner android.permission.PACKAGE_USAGE_STATS"
Shell "cmd /c adb shell pm grant com.zacharee1.systemuituner android.permission.DUMP"
MsgBox "������ɣ�"
Command2.Enabled = 1
Command2.Caption = "����"
End Sub

Private Sub Command3_Click()
Shell "explorer https://www.iplaysoft.com/scrcpy.html"
End Sub

Private Sub Form_Load()
Text1 = "������Ϊ״̬����������ͼ������գ�YUYU���ֿ���һ������SystemUI Tuner��ʵ�����ֻ��˷�������ػ�ָ��ֻ�״̬��ͼ�꣬�����ܷ�������ز���Ҫ��ͼ�ꡣҲ����������Google play�̵�����SystemUI Tuner���ء�"
Text2 = "Scrcpy��һ�Դ����������������adb����׿�ֻ��Ļ���Ͷ��������������ʾ�ϲ����вٿأ�ʵ�������ڻ�Ϊ����Эͬ��Ч�����ֻ������谲װ�κ�Ӧ�á���������YUYU������ʹ������/���߷�ʽ�������ֻ���������Scrcpy.exe�����������κ�ָ���ʵ�����ߺ�����Ͷ����"
Text3 = "����һ����ʵ������EMUI10.1�ǻ۷����������������Ҳ���Կ�ݻظ�΢�š���Ȼ�Ҹ�����Ϊ��̫���ã���������λ����Ȥ�������ԡ�YUYU����ͬ������һ��������������̳�����Ӧ���ڲ鿴��"
End Sub

Private Sub Label2_Click()
Shell "explorer https://www.coolapk.com/apk/com.wintheshow.quickreply"
End Sub


Private Sub Label5_Click()
Shell "explorer https://azetrue.lanzous.com/iTbWQff3moj"
End Sub

Private Sub Label6_Click()
Shell "explorer https://azetrue.lanzous.com/iFrlRff3hej"
End Sub

Private Sub Label7_Click()
Shell "explorer https://github.com/Genymobile/scrcpy"
End Sub
