VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "YUYU助手v1.9"
   ClientHeight    =   8325
   ClientLeft      =   2280
   ClientTop       =   1260
   ClientWidth     =   5280
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   5280
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6480
      TabIndex        =   39
      Text            =   "&HC0C000"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "调整动画速度"
      Enabled         =   0   'False
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
      Left            =   1440
      MaskColor       =   &H00FFFF80&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "安装应用到手机"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   6720
      TabIndex        =   27
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command22 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F1E7&
      Caption         =   "查看所有应用"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "设定手机IP"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   5640
      TabIndex        =   23
      Top             =   3480
      Width           =   3735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2655
      Left            =   5640
      TabIndex        =   22
      Top             =   240
      Width           =   4815
      ExtentX         =   8493
      ExtentY         =   4683
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton Command21 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "重试"
      Enabled         =   0   'False
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
      Left            =   3120
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5640
      TabIndex        =   18
      Text            =   "al"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Text            =   "co"
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Text            =   "cannot"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "快捷操作"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   4815
      Begin VB.CommandButton Command19 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "启用自定清单应用"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command18 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "启用所有已停用应用"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command17 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "停用自定清单应用"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton Command16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "自定义应用清单"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "其他功能"
      Enabled         =   0   'False
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "建立无线连接"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "重试有线连接"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "强制开启"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC671&
      Caption         =   "启动"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFC0C0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   5400
      Top             =   7800
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00DAEBB8&
      Caption         =   "退出助手"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7680
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "退出但不结束adb"
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
      Height          =   300
      Left            =   2160
      MaskColor       =   &H00808080&
      TabIndex        =   33
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "收集错误"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      TabIndex        =   41
      Top             =   6480
      Width           =   840
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   120
      Top             =   6840
      Width           =   5055
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "打开cmd窗口"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   3480
      TabIndex        =   40
      Top             =   6480
      Width           =   1275
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "更多设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   2400
      TabIndex        =   38
      Top             =   6480
      Width           =   840
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "检查更新"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   1320
      TabIndex        =   37
      Top             =   6480
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "找到我："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      TabIndex        =   35
      Top             =   6960
      Width           =   840
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "酷安"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006DC725&
      Height          =   300
      Left            =   3360
      TabIndex        =   32
      Top             =   6960
      Width           =   420
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bilibili"
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
      Left            =   2400
      TabIndex        =   31
      Top             =   6960
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "花粉俱乐部"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1080
      TabIndex        =   30
      Top             =   6960
      Width           =   1050
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "点击更改adb端口"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005809FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "总是连接失败？"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005809FF&
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   2280
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "查看详细教程"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005809FF&
      Height          =   330
      Left            =   3720
      TabIndex        =   24
      Top             =   2640
      Width           =   1440
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "请勿更改下列值！！！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   21
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "断开无线连接"
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
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "启动中. . ."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006DC725&
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "更多功能"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   960
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   120
      Top             =   5640
      Width           =   5055
   End
   Begin VB.Image Image2 
      Height          =   50
      Left            =   -6400
      Picture         =   "Form1.frx":74CA
      Stretch         =   -1  'True
      Top             =   1650
      Visible         =   0   'False
      Width           =   6480
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "操作应用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   120
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   1650
      Left            =   0
      Picture         =   "Form1.frx":B24F
      Top             =   0
      Width           =   5280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L As String
Dim sr As String

Private Sub Command11_Click()
MsgBox "打开无线连接将使局域网内的其他设备均可能控制您的设备，请确认局域网设备及其安装的软件均安全后，再开启", 1 + 48, "警告"
Dim strFile  As String
Dim intFile  As Integer
Label4 = "正在建立无线连接. . ."
Shell "cmd /c adb devices"
Shell "cmd /c adb tcpip 5555"
If Text1 = "" Then  'F1
Form9.Show
Label4 = "没有连接记录，请先设定手机IP地址"
Else          'else
Label4 = Label4 & vbCrLf & "正在连接至：" & Text1
Shell "cmd /c adb connect " & Text1 & ">" & Text5 & "\IPC.txt"
Savetime = Timer
While Timer < Savetime + 2
DoEvents
Wend
 Dim strDa  As String
    strFile = Text5 & "\IPC.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strDa = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strDa
    Close intFile
If Text2.Text = Left(strDa, 6) Or strDa = "" Then   'F2
Label4 = "建立无线连接失败!"
Call Command7_Click
ElseIf Text3.Text = Left(strDa, 2) Or Text4.Text = Left(strDa, 2) Then
Label4 = "建立无线连接成功！" & vbCrLf & "现在请您拔掉数据线后点击“重试”按钮"
Shell "cmd /c adb disconnect "
Command11.Enabled = False
Command11.Caption = "请按 重试 按钮"
Call Command7_Click
End If
End If
End Sub
Private Sub Command12_Click()
Form1.Show
d = 0
Image2.Left = -6240
Timer1.Interval = 28
Timer1.Enabled = True
Image2.Visible = True
Dim Savetime As Single
Label4 = "等待adb服务启动. . . " & vbCrLf & "亮屏更容易无线连接成功哦"
Savetime = Timer
While Timer < Savetime + 2
DoEvents
Wend

Label4 = Label4 & vbCrLf & "检查无线连接. . . "
Shell "cmd /c adb connect " & Text1 & ">" & Text5 & "\IPC.txt"
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend

Dim strFile  As String
Dim intFile  As Integer
 Dim strDa  As String
    strFile = Text5 & "\IPC.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strDa = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strDa
    Close intFile
Shell "cmd /c adb shell getprop ro.product.model" & ">" & Text5 & "\YUYU.txt"
Label4 = Label4 & vbCrLf & "检查有线连接. . . "
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend

    Dim strData  As String
    strFile = Text5 & "\YUYU.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile

If Text2.Text = Left(strDa, 6) Or strDa = "" Then   'F1
Label4 = "无线连接失败!"
If InStr(strData, "-") = 0 Then '有线连接也失败    'F2
Label4 = Label4 & vbCrLf & "有线连接也失败!" & vbCrLf & strDa & "没有已连接的设备，建议您检查操作后重试"
MsgBox "没有检查到已连接的设备！", 0 + 48, "设备未连接"
If Dir(Text5 & "\adb.txt") <> "" Then d = 1
ElseIf InStr(strData, "-") > 0 Then '有线连接成功
Label4 = Label4 & vbCrLf & "有线连接成功！" & vbCrLf & "欢迎你：" & strData
Command11.Enabled = 1
If InStr(Mid(sr, 3, 1), "1") > 0 Then Form11.Show
Call Command2_Click
End If    'end1
ElseIf Text3.Text = Left(strDa, 2) Or Text4.Text = Left(strDa, 2) Then
If InStr(strData, "-") = 0 Then  'f3
Label4 = "无线连接成功！但您同时连接了数据线！" & vbCrLf & "需要拨出数据线 或 点击“改为有线连接”，" & vbCrLf & "才能正常执行操作！(若识别错误请点击“重试”)"
ElseIf InStr(strData, "-") > 0 Then
Label4 = "无线连接成功！" & vbCrLf & strDa & "欢迎你：" & strData
Call Command2_Click
Label7.Visible = True
If InStr(Mid(sr, 3, 1), "1") > 0 Then Form11.Show
End If  'end2
Command6.Caption = "改为有线连接"
End If       'end3
Command21.Enabled = True
Command6.Enabled = True

Timer1.Enabled = False
Timer1.Interval = 15
Image2.Visible = False

If d = 1 Then
Label10.Visible = True: Label11.Visible = True
ElseIf d = 0 Then
Label10.Visible = False: Label11.Visible = False
End If

End Sub

Private Sub Command13_Click()
Form5.Show
End Sub

Private Sub Command14_Click()
Label4 = ""
Form10.Show
End Sub

Private Sub Command16_Click()
Label4 = ""
Form7.Show
End Sub

Private Sub Command17_Click()
Dim S, c, n As String
Dim iLine As Integer
If Dir(Text5 & "\List1.txt") <> "" Then
Open Text5 & "\List1.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
n = n & vbCrLf & S
Loop
m = MsgBox("确认停用以下应用吗？" & n, 1, "停用应用")
Close #1
If m = 1 Then
Open Text5 & "\List1.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
If InStr(S, ":") > 0 Then
c = Mid(S, InStr(S, ":") + 1)
ElseIf InStr(S, ":") = 0 Then
c = S
End If
Shell "cmd /c adb shell pm disable-user " & c
Loop
Close #1
Label4 = "已停用所有自定义清单的应用"
MsgBox "已停用" & n
Else
Label4 = "已取消"
End If
Else
MsgBox "您还未自定义应用清单！"
End If
End Sub

Private Sub Command18_Click()
Label4 = "请稍后..."
Shell "cmd /c adb shell pm list packages -d" & ">" & Text5 & "\disable.txt"
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Dim S, V, n As String
Open Text5 & "\disable.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
V = Mid(S, 9)
n = n & vbCrLf & V
Loop
Close #1
If S = "" Then
MsgBox "当前没有已停用的应用！"
Label4 = "当前没有已停用的应用"
Else
m = MsgBox("确认启用以下应用吗？" & n, 1, "启用所有已停用应用")
If m = 1 Then
Open Text5 & "\disable.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
V = Mid(S, 9)
Shell "cmd /c adb shell pm enable " & V
Loop
Close #1
Label4 = "已启用所有停用的应用"
MsgBox "已启用" & n
Else
Label4 = "已取消"
End If
End If
End Sub

Private Sub Command19_Click()
Dim S, c, n As String
Dim iLine As Integer
If Dir(Text5 & "\List1.txt") <> "" Then
Open Text5 & "\List1.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
n = n & vbCrLf & S
Loop
m = MsgBox("确认启用以下应用吗？" & n, 1, "启用应用")
Close #1
If m = 1 Then
Open Text5 & "\List1.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
If InStr(S, ":") > 0 Then
c = Mid(S, InStr(S, ":") + 1)
ElseIf InStr(S, ":") = 0 Then
c = S
End If
Shell "cmd /c adb shell pm enable " & c
Loop
Close #1
Label4 = "已启用所有自定义清单的应用"
MsgBox "已启用" & n
Else
Label4 = "已取消"
End If
Else
MsgBox "您还未自定义应用清单！"
End If
End Sub

Private Sub Command2_Click()
Command3.Enabled = True
Command14.Enabled = True
Command17.Enabled = True
Command18.Enabled = True
Command19.Enabled = True
Command16.Enabled = True
Command22.Enabled = True
Command8.Enabled = True
If Dir(Text5 & "\List1.txt") = "" Then
Command19.Enabled = 0: Command17.Enabled = 0
End If
End Sub

Private Sub Command20_Click()
Form9.Show
End Sub

Private Sub Command21_Click() '重试
d = 0
Command6.Enabled = False: Command11.Enabled = False: Command21.Enabled = False:
Image2.Left = -6240
Timer1.Interval = 28
Timer1.Enabled = True
Image2.Visible = True
Label4 = "检查无线连接. . . "
Shell "cmd /c adb connect " & Text1 & ">" & Text5 & "\IPC.txt"
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Shell "cmd /c adb shell getprop ro.product.model" & ">" & Text5 & "\YUYU.txt"
Label4 = Label4 & vbCrLf & "检查有线连接. . . "
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Dim strFile  As String
Dim intFile  As Integer
 Dim strDa  As String
    strFile = Text5 & "\IPC.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strDa = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strDa
    Close intFile
    Dim strData  As String
    strFile = Text5 & "\YUYU.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile

If Text2.Text = Left(strDa, 6) Or strDa = "" Then   'F1
Label4 = "无线连接失败!"
If InStr(strData, "-") = 0 Then '有线连接也失败    'F2
Label4 = Label4 & vbCrLf & "有线连接也失败!" & vbCrLf & strDa & "没有已连接的设备，建议您查看详细教程"
Call Command7_Click
MsgBox "没有检查到已连接的设备，建议您查看详细教程", 0 + 48, "设备未连接"
d = 1
ElseIf InStr(strData, "-") > 0 Then '有线连接成功
Label4 = Label4 & vbCrLf & "有线连接成功！" & vbCrLf & "欢迎你：" & strData
Command11.Enabled = 1
If InStr(Mid(sr, 3, 1), "1") > 0 Then Form11.Show
Call Command2_Click
End If    'end1
ElseIf Text3.Text = Left(strDa, 2) Or Text4.Text = Left(strDa, 2) Then
If InStr(strData, "-") = 0 Then  'f3
Label4 = "无线连接成功！但您同时连接了数据线！" & vbCrLf & "需要拨出数据线或点击按钮改为有线连接，" & vbCrLf & "才能正常执行操作！"
Call Command7_Click
ElseIf InStr(strData, "-") > 0 Then
Label4 = "无线连接成功！" & vbCrLf & strDa & "欢迎你：" & strData
Call Command2_Click
If InStr(Mid(sr, 3, 1), "1") > 0 Then Form11.Show
Label7.Visible = True
End If  'end2
Command6.Caption = "改为有线连接"
End If       'end3
Command21.Enabled = True
Command6.Enabled = True

Timer1.Enabled = False
Image2.Visible = False
If d = 1 Then
Label10.Visible = True: Label11.Visible = True
ElseIf d = 0 Then
Label10.Visible = False: Label11.Visible = False
End If

End Sub

Private Sub Command22_Click()
If Dir(Text5 & "\all.txt") = "" Then
Open Text5 & "\all.txt" For Append As #5
Close #5
End If
Form11.Show
End Sub

Private Sub Command3_Click()
Label4 = ""
Form13.Show
End Sub

Private Sub Command4_Click()
If Check1.Value = 1 Then
End
Else
Shell "cmd /c adb kill-server"
Label4 = "已结束adb进程"
Dim Savetime As Single
Savetime = Timer
While Timer < Savetime + 0.5
DoEvents
Wend
End
End If
End Sub

Private Sub Command6_Click() '改为有线
d = 0
Command6.Caption = "重试有线连接"
Command6.Enabled = False: Command11.Enabled = False: Command21.Enabled = False:
Image2.Left = -6240
Timer1.Enabled = True
Image2.Visible = True
Dim Savetime As Single
Label4 = "改为有线连接. . . "
Shell "cmd /c adb disconnect " & Text1
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Shell "cmd /c adb shell getprop ro.product.model" & ">" & Text5 & "\YUYU.txt"
Label4 = Label4 & vbCrLf & "检查有线连接. . . "
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Dim strFile  As String
Dim intFile  As Integer
    Dim strData  As String
    strFile = Text5 & "\YUYU.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile

If InStr(strData, "-") = 0 Then '有线连接失败
Label4 = "有线连接失败!" & vbCrLf & strDa & "没有检查到有线连接的设备"
Call Command7_Click
d = 1
MsgBox "没有检查到已连接的设备，建议您查看详细教程", 0 + 48, "设备未连接"
ElseIf InStr(strData, "-") > 0 Then '有线连接成功
Label4 = "有线连接成功！" & vbCrLf & "欢迎你：" & strData
Command11.Enabled = 1
If InStr(Mid(sr, 3, 1), "1") > 0 Then Form11.Show
Call Command2_Click
End If    'end1
Command21.Enabled = True
Command6.Enabled = True

Timer1.Enabled = False
Image2.Visible = False
If d = 1 Then
Label10.Visible = True: Label11.Visible = True
ElseIf d = 0 Then
Label10.Visible = False: Label11.Visible = False
End If
Kill Text5 & "\YUYU.txt"
End Sub

Private Sub Command7_Click()
Command3.Enabled = 0
Command14.Enabled = 0
Command17.Enabled = 0
Command18.Enabled = 0
Command19.Enabled = 0
Command16.Enabled = 0
Command22.Enabled = 0
Command8.Enabled = 0
End Sub


Private Sub Command8_Click()
Form14.Show
End Sub

Private Sub Command9_Click()
Form9.Show
End Sub

Private Sub Form_Load()
Dim p
Dim FSO As New FileSystemObject
If FSO.FolderExists("c:\YUYU") Then
p = "c:\YUYU"
ElseIf FSO.FolderExists("d:\YUYU") Then
p = "d:\YUYU"
ElseIf FSO.FolderExists("e:\YUYU") Then
p = "e:\YUYU"
ElseIf FSO.FolderExists("f:\YUYU") Then
p = "f:\YUYU"
ElseIf FSO.FolderExists("g:\YUYU") Then
p = "g:\YUYU"
Else
p = "c:\YUYU"
FSO.CreateFolder (p)
MsgBox "数据默认存储在" & p & "，可自行更改位置。", 0, "欢迎使用YUYU助手！"
Shell "explorer https://c.xiumi.us/board/v5/3JD5p/215117062"
End If
Text5 = p

Shell "cmd /c adb devices" & ">" & Text5 & "\dev.txt" '调起adb
WebBrowser1.navigate "http://www.wwnote.xyz/ao/check.html"
'WebBrowser1.navigate "https://v.xiumi.us/board/v5/3JD5p/245146051" '暂时停用
If Dir(Text5 & "\set.txt") <> "" Then
Open Text5 & "\set.txt" For Input As #1
Input #1, sr
Close #1
Else
Open Text5 & "\set.txt" For Append As #1 '没有ip新建
Close #1
Open Text5 & "\set.txt" For Output As #1 '写入

sr = "010"
Print #1, sr
Close #1
End If

If Dir(Text5 & "\IP1.txt") <> "" Then
Text1 = ""
Else
Open Text5 & "\IP1.txt" For Append As #1 '没有ip新建
Close #1
End If
i = 0
Open Text5 & "\IP1.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
i = i + 1
Loop
Close #1
Dim strTemp As String
If i = 1 Then
Open Text5 & "\IP1.txt" For Input As #1
Line Input #1, strTemp    '读入一行数据
Text1 = strTemp
Close #1
End If
Call Command12_Click  '启动

If InStr(Mid(sr, 1, 1), "1") > 0 Then
Call Command2_Click
Label4 = "默认连接成功模式，已解锁按钮"
End If
On Error Resume Next
'Text6.Text = WebBrowser1.Document.body.innerText
'St = Mid(Text6, InStr(Text6, "http"), InStr(Text6, "html") - InStr(Text6, "http") + 4)
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Text6.Text = WebBrowser1.Document.body.innerText
L = Mid(Text6, 5, InStr(Text6, "RG") - 5)
u = Mid(Text6, 1, 3)

If u > 1.9 And u < 3 Then                   '数值！！！
Label14.ForeColor = vbRed
Label14 = "有新版！"
If InStr(Mid(sr, 2, 1), "1") > 0 Then
m = MsgBox("新版本v" & u & "发布啦，是否去升级", vbYesNo, "有新版本可用")
If m = vbYes Then Shell "explorer https://v.xiumi.us/board/v5/3JD5p/228148533"
End If
End If
Dim v_Path As String, K As Long, e, v_Range
v_Path = Text5 & "\ad1"
WebBrowser1.Silent = True '关闭交互，禁止脚本错误
For Each e In WebBrowser1.Document.All
    If e.tagName = "IMG" Then
        Set v_Range = WebBrowser1.Document.body.createControlRange()
        v_Range.Add e
        v_Range.execCommand "Copy" '复制到剪贴板
        K = K + 1
        SavePicture Clipboard.GetData, v_Path & ".jpg" '保存到硬盘
    End If
Next
If Dir(Text5 & "\YUYU.txt") <> "" Then
Kill Text5 & "\YUYU.txt"
End If
If Dir(Text5 & "\ad1.jpg") <> "" Then
Image1.Picture = LoadPicture(Text5 & "\ad1.jpg")
Kill Text5 & "\ad1.jpg"
End If
Text7 = Mid(Text6, InStr(Text6, "RG") + 3)
Label4.ForeColor = Text7
End Sub


Private Sub Form_Unload(Cancel As Integer)
m = MsgBox("您真的要退出？", vbExclamation + vbYesNo + vbDefaultButton2, "退出")
If m = vbNo Then
Cancel = True
ElseIf m = vbYes Then
Call Command4_Click
End If
End Sub

Private Sub Image1_Click()
Shell "explorer " & L
End Sub

Private Sub Label11_Click()
If Dir(Text5 & "\adb.txt") <> "" Then
w = MsgBox("请您确定你已经以管理员身份运行本程序", 1 + 48)
If w = 1 Then
Randomize
Dim m As Integer
m = Int(Rnd * 200 + 7800)
Dim n As String
n = """" & m & """"
Shell "cmd /c setx /M ANDROID_ADB_SERVER_PORT " & n
MsgBox "已将adb端口更改为" & n & "，请重启电脑使环境变量生效"
Label4 = "请重启电脑后再尝试连接！"
Else
Label4 = "请右键以管理员身份运行本程序"
End If
ElseIf Dir(Text5 & "\adb.txt") = "" Then
Open Text5 & "\adb.txt" For Append As #1
Close #1
Shell "explorer https://v.xiumi.us/board/v5/3JD5p/203664485"  '!!!
Call Command4_Click
End If
End Sub

Private Sub Label12_Click()
Shell "explorer http://www.coolapk.com/u/3801286"
End Sub



Private Sub Label13_Click()
Shell "C:\WINDOWS\system32\cmd.exe"
Label13 = "注意在任务栏"
Savetime = Timer
While Timer < Savetime + 4
DoEvents
Wend
Label13 = "打开cmd窗口"
End Sub

Private Sub Label14_Click()
Label4 = "正在检查更新"

u = Mid(Text6, 1, 3)
If u > 1.9 Then           '数值！！！
m = MsgBox("新版本v" & u & "发布啦，是否去升级", vbYesNo, "有新版本可用")
Label4 = "有新版本可升级！"
Label14.ForeColor = vbRed
Label14 = "有版本啦！"
If m = vbYes Then Shell "explorer https://v.xiumi.us/board/v5/3JD5p/228148533"
Else
MsgBox "暂无新版本"
Label4 = "暂无新版本"
End If

End Sub

Private Sub Label15_Click()
Label4 = ""
Form4.Show
End Sub

Private Sub Label16_Click()
Form6.Show
End Sub

Private Sub Label3_Click()
Shell "explorer https://club.huawei.com/space-uid-7120962.html"
End Sub

Private Sub Label2_Click()
Shell "explorer https://v.xiumi.us/board/v5/3JD5p/203664485"
End Sub

Private Sub Label6_Click()
Shell "explorer https://space.bilibili.com/41840415"
End Sub

Private Sub Label7_Click()
Label7 = "正在断开"
Label7.Enabled = False
Shell "cmd /c adb disconnect " & Text1 & ">" & Text5 & "\DISC.txt"
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Dim strFile  As String
    Dim intFile  As Integer
    Dim strData  As String
    strFile = Text5 & "\DISC.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile
Label4 = "已断开无线连接 " & strData
Label7.Enabled = True
Label7 = "断开无线连接"
End Sub



Private Sub Timer1_Timer()
Image2.Left = Image2.Left + 44
End Sub

