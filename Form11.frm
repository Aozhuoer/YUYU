VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form11 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "所有应用"
   ClientHeight    =   8490
   ClientLeft      =   7650
   ClientTop       =   1275
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   9960
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "列表中显示的应用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   5640
      TabIndex        =   25
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "停用的应用"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "启用的应用"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "全部应用"
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "全选"
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
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   160
      TabIndex        =   24
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C9A8FF&
      Caption         =   "中断操作"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9480
      Top             =   1680
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   13680
      TabIndex        =   22
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   1455
      Left            =   13560
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "提取"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   975
   End
   Begin VB.ListBox List3 
      Height          =   1335
      Left            =   11520
      TabIndex        =   18
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   13800
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "应用分类"
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
      Left            =   5640
      TabIndex        =   13
      Top             =   5880
      Width           =   4095
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "其他系统应用"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "华为全家桶"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "谷歌全家桶"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C9A8FF&
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MaskColor       =   &H00C9A8FF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7560
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6630
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.ListBox List2 
      Height          =   3120
      Left            =   10680
      TabIndex        =   9
      Top             =   3360
      Width           =   3735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3015
      Left            =   10680
      TabIndex        =   8
      Top             =   480
      Width           =   3735
      ExtentX         =   6588
      ExtentY         =   5318
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
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   11160
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "清除"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "启用"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "停用"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F1E7&
      Caption         =   "搜索应用"
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
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "在这里输入应用中文名或英文包名"
      Top             =   120
      Width           =   4335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7290
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   960
      Width           =   5295
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "操作选中的应用"
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
      Left            =   5640
      TabIndex        =   19
      Top             =   3840
      Width           =   4095
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "强烈建议您卸载前先提取备份应用"
         ForeColor       =   &H005809FF&
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   1160
         Width           =   2415
      End
   End
   Begin VB.Image Image1 
      Height          =   75
      Left            =   -5760
      Picture         =   "Form11.frx":74CA
      Stretch         =   -1  'True
      Top             =   520
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005809FF&
      Height          =   1545
      Left            =   5520
      TabIndex        =   12
      Top             =   840
      Width           =   4320
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r, g As Integer

Private Sub Check1_Click()

End Sub

Private Sub Check2_Click()
Dim i As Integer
If Check2 = 1 Then
If g = 0 Then
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
ElseIf g = 1 Then
For i = 0 To List4.ListCount - 1
List4.Selected(i) = True
Next
End If

ElseIf Check2 = 0 Then
If g = 0 Then
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
ElseIf g = 1 Then
For i = 0 To List4.ListCount - 1
List4.Selected(i) = False
Next
End If
End If
End Sub

Private Sub Command11_Click()
m = MsgBox("您确定要中断操作并退出YUYU助手吗？可能造成未知的后果！", 1 + 48, "警告")
If m = 1 Then
Shell "cmd /c adb kill-server"
End
End If
End Sub

Private Sub Option1_Click()
List1.Clear

Open Text2 & "\all.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
j = 0
V = Mid(S, 9)      '手机应用
For i = 0 To List2.ListCount - 1  '网上列表
y = List2.List(i)
n = InStr(y, ":")
If n > 0 Then
c = Mid(y, n + 1)
Else
c = ""
End If
If c = V Then
List1.AddItem y
j = 1
Exit For
End If
Next i
If j = 0 Then List1.AddItem V
Loop
Close #1

r = 1
g = 0
Label2 = "加载完毕！"

End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Label2 = "正在加载..."
If Dir(Text2 & "\enable.txt") <> "" Then
Kill Text2 & "\enable.txt"
End If

Shell "cmd /c adb shell pm list packages -e" & ">" & Text2 & "\enable.txt"
List1.Clear
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend

Dim S, V, n As String
Open Text2 & "\enable.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
j = 0
V = Mid(S, 9)      '手机应用
For i = 0 To List2.ListCount - 1  '网上列表
y = List2.List(i)
n = InStr(y, ":")
If n > 0 Then
c = Mid(y, n + 1)
Else
c = ""
End If
If c = V Then
List1.AddItem y
j = 1
Exit For
End If
Next i
If j = 0 Then List1.AddItem V
Loop
Close #1
r = 1
g = 0
Label2 = "加载完毕！"
If c = "" Then MsgBox "当前没有已停用的应用！"
If Dir(Text2 & "\enable.txt") <> "" Then
Kill Text2 & "\enable.txt"
End If
ElseIf Option3.Value Then

List1.Clear

Open Text2 & "\all.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
j = 0
V = Mid(S, 9)      '手机应用
For i = 0 To List2.ListCount - 1  '网上列表
y = List2.List(i)
n = InStr(y, ":")
If n > 0 Then
c = Mid(y, n + 1)
Else
c = ""
End If
If c = V Then
List1.AddItem y
j = 1
Exit For
End If
Next i
If j = 0 Then List1.AddItem V
Loop
Close #1

r = 1
g = 0
Label2 = "加载完毕！"
End If
End Sub


Private Sub Option3_Click()
If Option3.Value = True Then
Label2 = "正在加载..."
If Dir(Text2 & "\disable.txt") <> "" Then
Kill Text2 & "\disable.txt"
End If

Shell "cmd /c adb shell pm list packages -d" & ">" & Text2 & "\disable.txt"
List1.Clear
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend

Dim S, V, n As String
Open Text2 & "\disable.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
j = 0
V = Mid(S, 9)      '手机应用
For i = 0 To List2.ListCount - 1  '网上列表
y = List2.List(i)
n = InStr(y, ":")
If n > 0 Then
c = Mid(y, n + 1)
Else
c = ""
End If
If c = V Then
List1.AddItem y
j = 1
Exit For
End If
Next i
If j = 0 Then List1.AddItem V
Loop
Close #1
r = 1
g = 0
Label2 = "加载完毕！"
If List1.ListCount = 0 Then MsgBox "当前没有已停用的应用！"
If Dir(Text2 & "\disable.txt") <> "" Then
Kill Text2 & "\disable.txt"
End If
ElseIf Option3.Value Then

List1.Clear

Open Text2 & "\all.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
j = 0
V = Mid(S, 9)      '手机应用
For i = 0 To List2.ListCount - 1  '网上列表
y = List2.List(i)
n = InStr(y, ":")
If n > 0 Then
c = Mid(y, n + 1)
Else
c = ""
End If
If c = V Then
List1.AddItem y
j = 1
Exit For
End If
Next i
If j = 0 Then List1.AddItem V
Loop
Close #1

r = 1
g = 0
Label2 = "加载完毕！"
End If
End Sub

Private Sub Timer1_Timer()
If Image1.Left > 9600 Then Image1.Left = -5780
Image1.Left = Image1.Left + 50
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, Url As Variant)

If (pDisp Is WebBrowser1.Object) Then
Label2 = "正在获取中文名称，请耐心等待..."
End If

End Sub


Private Sub Command1_Click()
List4.Clear
j = 0
For i = 0 To List1.ListCount - 1
y = List1.List(i)
n = InStr(y, Text1)
If n > 0 Then
List4.Visible = True
List1.Visible = 0
List4.AddItem y
j = 1
g = 1
Command7.Visible = True
End If
Next i
Frame3.Enabled = 0
Option1.ForeColor = &H808080: Option2.ForeColor = &H808080: Option3.ForeColor = &H808080
If j = 0 Then
MsgBox "没有找到任何结果"
Frame3.Enabled = 1
Option1.ForeColor = &H8000000D: Option2.ForeColor = &H8000000D: Option3.ForeColor = &H8000000D
End If
Command6.Visible = True

End Sub


Private Sub Command10_Click()
Command2.Enabled = False: Command3.Enabled = False: Command4.Enabled = False: Command10.Enabled = False
If List1.SelCount = 0 And List4.SelCount = 0 Then  'IF1
MsgBox "没有选中的项目！双击应用名称或勾选名称前的方框即可选中项目"
'----------------------------------------------------------------------------------
Else
Shell "cmd /c adb decices"
Image1.Visible = True
Timer1.Enabled = True
Dim Savetime As Single
f = MsgBox("推荐使用有线连接传输。若使用无线连接请保持手机亮屏！！否则失败率极高！" & vbCrLf & "现在，您是否要开始提取应用？", 4 + 48, "提示")
If f = vbYes Then
Command11.Visible = True
Dim i, t, h, e As Integer
Dim y, z As String
h = 0
Label2 = "正在获取应用路径..."
If g = 0 Then             'IF2
For i = 0 To List1.ListCount - 1
If List1.Selected(i) Then           'IF3
y = List1.List(i)

n = InStr(y, ":")
If n > 0 Then     'IF4
z = Mid(y, n + 1)
ElseIf n = 0 Then
z = y
End If          'END4

h = h + 1
If Dir(Text2 & "\appname" & h & ".txt") <> "" Then Kill Text2 & "\appname" & h & ".txt"
Open Text2 & "\appname" & h & ".txt" For Append As #1 '没有新建
Print #1, y
Close #1
Shell "cmd /c adb shell pm path " & z & ">" & Text2 & "\path" & h & ".txt"
If h = 1 Then
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
End If
End If     'END3

Next i
ElseIf g = 1 Then
For t = 0 To List4.ListCount - 1
If List4.Selected(t) Then  'IF3.2
y = List4.List(t)
n = InStr(y, ":")
If n > 0 Then    'IF4.2
z = Mid(y, n + 1)
ElseIf n = 0 Then
z = y
End If
'E4.2
h = h + 1
If Dir(Text2 & "\appname" & h & ".txt") <> "" Then Kill Text2 & "\appname" & h & ".txt"
Open Text2 & "\appname" & h & ".txt" For Append As #1 '没有新建
Print #1, y
Close #1
Shell "cmd /c adb shell pm path " & z & ">" & Text2 & "\path" & h & ".txt"
If h = 1 Then
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
End If
End If     'END3
Next t
End If       'END2

'---------------------------------------------------获取完路径
Savetime = Timer
While Timer < Savetime + 0.5
DoEvents
Wend
Dim b, m As Integer
b = 0
m = 0
Text3 = ""
'---------------------------------------
For e = 1 To h
'------------------------------------打开文件
Dim strFile  As String
    Dim intFile  As Integer
    Dim strData  As String
    strFile = Text2 & "\path" & e & ".txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile
Text5 = strData
Do
ab = InStr(Text5, "/")
Text5 = Mid(Text5, ab + 1)
Loop Until ab = 0
cb = InStr(Text5, "apk")
Text5 = Left(Text5, cb + 2)
Pa = Mid(strData, 9)
    strFile = Text2 & "\appname" & e & ".txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile
Label2 = "正在准备提取" & strData
Shell "cmd /c adb pull " & Pa

If Text5 = "" Then                 'IF A1
MsgBox strData & "无法被提取！点击确定提取下一个应用"
Text3 = Text3 & strData & "提取失败！"
'________________________________________________________________________可以提取
Else            'else A1

o = 0          '正在准备提取
Dim Filepath$
Dim yz
Do

Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
o = o + 1
Label2 = "已用时" & o & "秒 正在准备提取 " & strData
Loop Until Dir(App.Path & "\" & Text5) <> "" Or o > 15

If o > 14 Then
MsgBox strData & "超时仍未开始传输，似乎无法被提取，点击确定提取下一个应用"
Text3 = Text3 & strData & "可能失败！" & vbCrLf
End If
    'IF A2
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Do          '提取中
Label2 = "正在提取，请耐心等待  " & strData
yz = FileLen(App.Path & "\" & Text5)
Savetime = Timer
While Timer < Savetime + 0.5
DoEvents
Wend
Loop Until yz > 0

Label2 = strData & "提取成功！安装包保存在" & App.Path
Text3 = Text3 & strData & "提取成功！"

'-----------------------------------准备重命名
n = InStr(strData, ":")
If n > 0 Then                                'IF A3
z = Mid(strData, 1, n - 1)
ElseIf n = 0 Then
If Text5 = "base.apk" Then
z = "base" & e
Text4 = Text4 & strData & "命名为" & z & ".apk" & vbCrLf
Else
zs = InStr(Text5, ".apk")
z = Mid(Text5, 1, zs - 1)
End If
End If                 'End A3

o = 0
Do Until Dir(App.Path & "\" & Text5) <> "" Or o > 3
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
o = o + 1
Label2 = "等待重命名" & z
Loop
'-----------------------------------重命名
If Dir(App.Path & "\" & Text5) <> "" Then  'if A4
Name App.Path & "\" & Text5 As App.Path & "\" & z & ".apk"
Else
MsgBox z & "重命名失败"
End If                                                  'End A4
     'End A2
Text3 = Text3 & vbCrLf
End If            'End A1
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Next e

'___________________________________________________________________________________________
For d = 1 To h              '删除
Kill Text2 & "\appname" & d & ".txt"
Kill Text2 & "\path" & d & ".txt"
Next d

'-------------------------------------结果
Label2 = "提取完成"
MsgBox Text3 & "提取的应用已保存至" & App.Path & "，请及时移走它们！"
Text3 = ""
If Text4 <> "" Then
MsgBox "请注意，由于安装包没有名字：" & vbCrLf & Text4, 0, "重要提示！"
Text4 = ""
End If
'------------------------------------------
End If
Image1.Visible = 0
Timer1.Enabled = 0
Command11.Visible = 0
End If                            'End 1
Command2.Enabled = 1: Command3.Enabled = 1: Command4.Enabled = 1: Command10.Enabled = 1
End Sub

Private Sub Command2_Click()
Command2.Enabled = False: Command3.Enabled = False: Command4.Enabled = False
If List1.SelCount = 0 And List4.SelCount = 0 Then  'IF1
MsgBox "没有选中的项目！双击应用名称或勾选名称前的方框即可选中项目"
Else
Dim i, t, h, e As Integer
Dim y, z As String
h = 0
If g = 0 Then             'IF2
For i = 0 To List1.ListCount - 1
If List1.Selected(i) Then           'IF3
y = List1.List(i)
n = InStr(y, ":")
If n > 0 Then     'IF4
z = Mid(y, n + 1)
ElseIf n = 0 Then
z = y
End If          'END4
List3.AddItem y
h = h + 1
Shell "cmd /c adb shell pm disable-user " & z & ">" & Text2 & "\临时" & h & ".txt"
End If     'END3
Next i
ElseIf g = 1 Then
For t = 0 To List4.ListCount - 1
If List4.Selected(t) Then  'IF3.2
y = List4.List(t)
n = InStr(y, ":")
If n > 0 Then    'IF4.2
z = Mid(y, n + 1)
ElseIf n = 0 Then
z = y
End If          'E4.2
List3.AddItem y
h = h + 1
Shell "cmd /c adb shell pm disable-user " & z & ">" & Text2 & "\临时" & h & ".txt"
End If       'E3.2
Next t
End If       'END2
Dim Savetime As Single
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Dim b, m As Integer
b = 0
m = 0
For e = 1 To h
Dim strFile  As String
    Dim intFile  As Integer
    Dim strData  As String
    strFile = Text2 & "\临时" & e & ".txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile
If InStr(strData, "Package") = 0 Then      'IF5
b = b + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " 失败"
ElseIf InStr(strData, "Package") = 1 Then
m = m + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " 成功"
End If       'END5
Kill Text2 & "\临时" & e & ".txt"
Next e
MsgBox m & "个成功，" & b & "个失败！" & Text3
If Option2.Value = True Then Call Option2_Click
End If      'END1
List3.Clear
Text3 = ""
Command2.Enabled = 1: Command3.Enabled = 1: Command4.Enabled = 1
End Sub

Private Sub Command3_Click()
Command2.Enabled = False: Command3.Enabled = False: Command4.Enabled = False
If List1.SelCount = 0 And List4.SelCount = 0 Then  'IF1
MsgBox "没有选中的项目！双击应用名称或勾选名称前的方框即可选中项目"
Else
Dim i, t, h, e As Integer
Dim y, z As String
h = 0
If g = 0 Then             'IF2
For i = 0 To List1.ListCount - 1
If List1.Selected(i) Then           'IF3
y = List1.List(i)
n = InStr(y, ":")
If n > 0 Then     'IF4
z = Mid(y, n + 1)
ElseIf n = 0 Then
z = y
End If          'END4
List3.AddItem y
h = h + 1
Shell "cmd /c adb shell pm enable " & z & ">" & Text2 & "\临时" & h & ".txt"
End If     'END3
Next i
ElseIf g = 1 Then
For t = 0 To List4.ListCount - 1
If List4.Selected(t) Then  'IF3.2
y = List4.List(t)
n = InStr(y, ":")
If n > 0 Then    'IF4.2
z = Mid(y, n + 1)
ElseIf n = 0 Then
z = y
End If          'E4.2
List3.AddItem y
h = h + 1
Shell "cmd /c adb shell pm enable " & z & ">" & Text2 & "\临时" & h & ".txt"
End If       'E3.2
Next t
End If       'END2
Dim Savetime As Single
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Dim b, m As Integer
b = 0
m = 0
For e = 1 To h
Dim strFile  As String
    Dim intFile  As Integer
    Dim strData  As String
    strFile = Text2 & "\临时" & e & ".txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile
If InStr(strData, "Package") = 0 Then      'IF5
b = b + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " 失败"
ElseIf InStr(strData, "Package") = 1 Then
m = m + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " 成功"
End If       'END5
Kill Text2 & "\临时" & e & ".txt"
Next e
MsgBox m & "个成功，" & b & "个失败！" & Text3
If Option3.Value = True Then Call Option3_Click
End If      'END1
List3.Clear
Text3 = ""
Command2.Enabled = 1: Command3.Enabled = 1: Command4.Enabled = 1
End Sub

Private Sub Command4_Click()
Command2.Enabled = False: Command3.Enabled = False: Command4.Enabled = False
If List1.SelCount = 0 And List4.SelCount = 0 Then  'IF1
MsgBox "没有选中的项目！双击应用名称或勾选名称前的方框即可选中项目"
Else
K = MsgBox("你确定卸载选中的应用吗？强烈建议您卸载前先提取备份！", 1 + 48, "警告")
If K = 1 Then
Dim i, t, h, e As Integer
Dim y, z As String
Label2 = "正在批量卸载，请耐心等待..."
h = 0
If g = 0 Then             'IF2
For i = 0 To List1.ListCount - 1
If List1.Selected(i) Then           'IF3
y = List1.List(i)
n = InStr(y, ":")
If n > 0 Then     'IF4
z = Mid(y, n + 1)
ElseIf n = 0 Then
z = y
End If          'END4
List3.AddItem y
h = h + 1
Shell "cmd /c adb shell pm uninstall --user 0 " & z & ">" & Text2 & "\临时" & h & ".txt"
End If     'END3
Next i
ElseIf g = 1 Then
For t = 0 To List4.ListCount - 1
If List4.Selected(t) Then  'IF3.2
y = List4.List(t)
n = InStr(y, ":")
If n > 0 Then    'IF4.2
z = Mid(y, n + 1)
ElseIf n = 0 Then
z = y
End If          'E4.2
List3.AddItem y
h = h + 1
Shell "cmd /c adb shell pm uninstall --user 0 " & z & ">" & Text2 & "\临时" & h & ".txt"
End If       'E3.2
Next t
End If       'END2

Dim Savetime As Single
Savetime = Timer
While Timer < Savetime + h * 2
DoEvents
Wend

Dim b, m As Integer
b = 0
m = 0
For e = 1 To h
Dim strFile  As String
    Dim intFile  As Integer
    Dim strData  As String
    strFile = Text2 & "\临时" & e & ".txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile
If InStr(strData, "Su") = 0 Then      'IF5
b = b + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " 失败"
ElseIf InStr(strData, "Su") = 1 Then
m = m + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " 成功"
End If       'END5
Next e
Label2 = "卸载完成"
MsgBox m & "个成功，" & b & "个失败！" & Text3, 0, "结果判断可能有误，请以手机为准"
For e = 1 To h
If Dir(Text2 & "\临时" & e & ".txt") <> "" Then
'Kill Text2 & "\临时" & e & ".txt"
End If
Next e
If Option1.Value = True Then Call Option1_Click
If Option2.Value = True Then Call Option2_Click
If Option3.Value = True Then Call Option3_Click
End If

End If      'END1
List3.Clear
Text3 = ""
Command2.Enabled = 1: Command3.Enabled = 1: Command4.Enabled = 1
End Sub

Private Sub Command5_Click()
Form2.Show
End Sub

Private Sub Command6_Click()
Text1 = ""
Command6.Visible = 0
End Sub

Private Sub Command7_Click()
List4.Clear
List4.Visible = False
List1.Visible = 1
Command7.Visible = 0
g = 0
Frame3.Enabled = 1
Option1.ForeColor = &H8000000D: Option2.ForeColor = &H8000000D: Option3.ForeColor = &H8000000D
End Sub

Private Sub Command8_Click()
Form3.Show
End Sub

Private Sub Command9_Click()
Form5.Show
End Sub

Private Sub Form_Load()
Text2 = Form1.Text5
List1.AddItem "Loading..."
List1.AddItem "若您未连接网络，则无法获取中文"
Shell "cmd /c adb shell pm list packages" & ">" & Text2 & "\all.txt"
If Dir(Text2 & "\all.txt") <> "" Then '有！
Dim x, S, V, A, y As String
WebBrowser1.navigate "http://www.wwnote.xyz/ao/applist.html"
Savetime = Timer
While Timer < Savetime + 0.5
DoEvents
Wend
Form11.Show
i = 0
Do
x = WebBrowser1.Document.body.innerText
If x <> "" Then Exit Do
If i > 5 Then
Exit Do
MsgBox "加载超时，无法获取应用中文，请检查您的网络连接", 0 + 48, "获取中文失败"
End If
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
i = i + 1
Label2 = "正在获取中文名称，已用时" & i & "秒"
Loop

If Dir(Text2 & "\in.txt") <> "" Then
Else
Open Text2 & "\in.txt" For Append As #1 '没有新建
Close #1
End If
Open Text2 & "\in.txt" For Output As #1
Print #1, x
Close #1
Open Text2 & "\in.txt" For Input As #2
Do While Not EOF(2)
Line Input #2, A
List2.AddItem A
Loop
Close #2
List1.Clear
Open Text2 & "\all.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
j = 0
V = Mid(S, 9)      '手机应用
For i = 0 To List2.ListCount - 1  '网上列表
y = List2.List(i)
n = InStr(y, ":")
If n > 0 Then
c = Mid(y, n + 1)
Else
c = ""
End If
If c = V Then
List1.AddItem y
j = 1
Exit For
End If
Next i
If j = 0 Then List1.AddItem V
Loop
Close #1
End If
r = 1
g = 0
gg = 0

Label2 = "加载完毕！"
Command1.Enabled = True: Command2.Enabled = True: Command3.Enabled = True: Command4.Enabled = True: Command6.Enabled = True: List1.Enabled = True:
Command5.Enabled = True: Command8.Enabled = True: Command9.Enabled = True: Command10.Enabled = True: Check2.Enabled = True: Frame3.Visible = True
Option1.Enabled = 1
Option1.Value = True
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Label2 = "部分系统关键应用无法被停用卸载，如：手机管家，华为桌面，健康使用手机等！请谨慎操作！！！"
End Sub

Private Sub List1_Click()
Text1 = List1.Text
r = 2
Command6.Visible = True
End Sub

Private Sub Text1_Click()
If r = 1 Then Text1 = ""
r = 2
Text1.ForeColor = vbBlack
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click
End Sub
