VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FFFFFF&
   Caption         =   "配合其他软件的功能"
   ClientHeight    =   8625
   ClientLeft      =   7725
   ClientTop       =   1260
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "微软雅黑"
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
      Caption         =   "使用教程"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "一键激活"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "一键激活"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "基于adb实现的应用推荐，仅提供下载链接，不属于任何推广，有任何问题请咨询它们的原作者"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "下载该程序："
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "手机悬浮窗"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "手机投屏到电脑"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "下载该程序："
      BeginProperty Font 
         Name            =   "微软雅黑"
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
         Name            =   "微软雅黑"
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
      Caption         =   "蓝奏云"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "蓝奏云"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "下载该应用："
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "手机隐藏状态栏图标"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
      Caption         =   "去酷安下载"
      BeginProperty Font 
         Name            =   "微软雅黑"
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
Command1.Caption = "正在执行"
Command1.Enabled = 0
Shell "cmd /c adb shell pm grant com.ksxkq.floating android.permission.WRITE_SECURE_SETTINGS"
Savetime = Timer
While Timer < Savetime + 2
DoEvents
Wend
Shell "cmd /c adb shell settings put global enable_freeform_support 1"
MsgBox "激活完成！现在请您重启手机"
Command1.Enabled = 1
Command1.Caption = "重试"
End Sub

Private Sub Command2_Click()
Command2.Caption = "正在执行"
Command2.Enabled = 0
Shell "cmd /c adb shell pm grant com.zacharee1.systemuituner android.permission.WRITE_SECURE_SETTINGS"
Shell "cmd /c adb shell pm grant com.zacharee1.systemuituner android.permission.PACKAGE_USAGE_STATS"
Shell "cmd /c adb shell pm grant com.zacharee1.systemuituner android.permission.DUMP"
MsgBox "激活完成！"
Command2.Enabled = 1
Command2.Caption = "重试"
End Sub

Private Sub Command3_Click()
Shell "explorer https://www.iplaysoft.com/scrcpy.html"
End Sub

Private Sub Form_Load()
Text1 = "若您因为状态栏经常挤满图标而烦恼，YUYU助手可以一键激活SystemUI Tuner以实现在手机端方便的隐藏或恢复手机状态栏图标，让您能方便的隐藏不需要的图标。也可以自行在Google play商店搜索SystemUI Tuner下载。"
Text2 = "Scrcpy是一款开源免费软件，可以利用adb将安卓手机的画面投屏到电脑桌面显示上并进行操控，实现类似于华为多屏协同的效果。手机端无需安装任何应用。您可以在YUYU助手中使用有线/无线方式先连接手机，再运行Scrcpy.exe，无需输入任何指令即可实现有线和无线投屏。"
Text3 = "简窗是一款能实现类似EMUI10.1智慧分屏悬浮窗的软件，也可以快捷回复微信。虽然我个人认为不太好用，但是若各位有兴趣可以试试。YUYU助手同样可以一键激活它，具体教程请在应用内查看。"
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
