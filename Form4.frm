VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "更多设置"
   ClientHeight    =   7575
   ClientLeft      =   7605
   ClientTop       =   2085
   ClientWidth     =   6060
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   6060
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   6960
      TabIndex        =   19
      Text            =   "Text4"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   2295
      Left            =   6840
      TabIndex        =   18
      Text            =   "Text3"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "连接成功后自动打开应用列表"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   7
      Top             =   3480
      Width           =   3135
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "有更新时弹窗提示"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "默认连接成功"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "清除所有数据"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "更改数据存储位置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "打赏"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "作者邮箱：1483544237@qq.com"
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "启动设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   5775
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "可能出现无中文的bug"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "启动时弹窗提示更新"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "非华为手机若无法连接，请勾选"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   2400
         TabIndex        =   9
         Top             =   480
         Width           =   2940
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "存储设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   5775
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "YUYU助手的数据文件存放位置"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   2520
         TabIndex        =   14
         Top             =   550
         Width           =   2850
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "这会清除所有用户数据！"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   300
         Left            =   2520
         TabIndex        =   13
         Top             =   1140
         Width           =   2310
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "欢迎大神帮我优化代码"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   2520
      TabIndex        =   17
      Top             =   7080
      Width           =   2100
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "如果你觉得还不错的话"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   1680
      TabIndex        =   16
      Top             =   6520
      Width           =   2100
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "去Github查看本项目"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   15
      Top             =   7080
      Width           =   1950
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1 = 1 Then
Text3 = "1" & Mid(Text3, 2, 2)
ElseIf Check1 = 0 Then
Text3 = "0" & Mid(Text3, 2, 2)
End If
Open Text4 & "\set.txt" For Output As #1
Print #1, Text3
Close #1
End Sub

Private Sub Check2_Click()
Dim ma, mc As String
ma = Mid(Text3, 1, 1)
mc = Mid(Text3, 3, 1)
If Check2.Value = 1 Then
Text3 = ma & "1" & mc
ElseIf Check2 = 0 Then
Text3 = ma & "0" & mc
End If
Open Text4 & "\set.txt" For Output As #1
Print #1, Text3
Close #1
End Sub

Private Sub Check3_Click()
Dim mb As String
mb = Mid(Text3, 1, 2)
If Check3 = 1 Then
Text3 = mb & "1"
ElseIf Check3 = 0 Then
Text3 = mb & "0"
End If
Open Text4 & "\set.txt" For Output As #1
Print #1, Text3
Close #1
End Sub

Private Sub Command1_Click()
Form8.Show
End Sub

Private Sub Command3_Click()
Form12.Show
End Sub

Private Sub Command4_Click()
m = MsgBox("清除所有数据会使程序恢复到初始状态！所有数据都将丢失！", 1 + 48, "警告！")
If m = 2 Then
Cancel = True
ElseIf m = 1 Then
Dim FSO As New FileSystemObject
FSO.DeleteFolder Form1.Text5
MsgBox "已清除所有数据，即将退出程序"
End
End If
End Sub

Private Sub Form_Load()
Text2 = "当前版本：v1.9" & vbCrLf & "作者：花粉俱乐部 @敖左耳" & vbCrLf & "声明：本程序永久免费，不得二次打包销售。若发现有人售卖请告知我!"
Text4 = Form1.Text5
Dim S As String
Open Text4 & "\set.txt" For Input As #1
Input #1, S
Text3 = S
Close #1
If InStr(Mid(Text3, 1, 1), "0") Then
Check1 = 0
ElseIf InStr(Mid(Text3, 1, 1), "1") Then
Check1 = 1
End If
If InStr(Mid(Text3, 2, 1), "0") Then
Check2 = 0
ElseIf InStr(Mid(Text3, 2, 1), "1") Then
Check2 = 1
End If
If InStr(Mid(Text3, 3, 1), "0") Then
Check3 = 0
ElseIf InStr(Mid(Text3, 3, 1), "1") Then
Check3 = 1
End If
End Sub

Private Sub Label4_Click()
Shell "explorer https://github.com/Aozhuoer/YUYU.exe"
End Sub

