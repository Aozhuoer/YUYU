VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form11 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "����Ӧ��"
   ClientHeight    =   8490
   ClientLeft      =   7650
   ClientTop       =   1275
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "΢���ź�"
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
      Caption         =   "�б�����ʾ��Ӧ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Caption         =   "ͣ�õ�Ӧ��"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���õ�Ӧ��"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ȫ��Ӧ��"
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
      Caption         =   "ȫѡ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�жϲ���"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��ȡ"
      Enabled         =   0   'False
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
      Caption         =   "Ӧ�÷���"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Caption         =   "����ϵͳӦ��"
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
         Caption         =   "��Ϊȫ��Ͱ"
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
         Caption         =   "�ȸ�ȫ��Ͱ"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      Caption         =   "���"
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
      Caption         =   "ж��"
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
      Caption         =   "����"
      Enabled         =   0   'False
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ͣ��"
      Enabled         =   0   'False
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F1E7&
      Caption         =   "����Ӧ��"
      Enabled         =   0   'False
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
      Text            =   "����������Ӧ����������Ӣ�İ���"
      Top             =   120
      Width           =   4335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "����ѡ�е�Ӧ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Caption         =   "ǿ�ҽ�����ж��ǰ����ȡ����Ӧ��"
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
         Name            =   "΢���ź�"
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
m = MsgBox("��ȷ��Ҫ�жϲ������˳�YUYU�����𣿿������δ֪�ĺ����", 1 + 48, "����")
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
V = Mid(S, 9)      '�ֻ�Ӧ��
For i = 0 To List2.ListCount - 1  '�����б�
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
Label2 = "������ϣ�"

End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Label2 = "���ڼ���..."
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
V = Mid(S, 9)      '�ֻ�Ӧ��
For i = 0 To List2.ListCount - 1  '�����б�
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
Label2 = "������ϣ�"
If c = "" Then MsgBox "��ǰû����ͣ�õ�Ӧ�ã�"
If Dir(Text2 & "\enable.txt") <> "" Then
Kill Text2 & "\enable.txt"
End If
ElseIf Option3.Value Then

List1.Clear

Open Text2 & "\all.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
j = 0
V = Mid(S, 9)      '�ֻ�Ӧ��
For i = 0 To List2.ListCount - 1  '�����б�
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
Label2 = "������ϣ�"
End If
End Sub


Private Sub Option3_Click()
If Option3.Value = True Then
Label2 = "���ڼ���..."
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
V = Mid(S, 9)      '�ֻ�Ӧ��
For i = 0 To List2.ListCount - 1  '�����б�
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
Label2 = "������ϣ�"
If List1.ListCount = 0 Then MsgBox "��ǰû����ͣ�õ�Ӧ�ã�"
If Dir(Text2 & "\disable.txt") <> "" Then
Kill Text2 & "\disable.txt"
End If
ElseIf Option3.Value Then

List1.Clear

Open Text2 & "\all.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
j = 0
V = Mid(S, 9)      '�ֻ�Ӧ��
For i = 0 To List2.ListCount - 1  '�����б�
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
Label2 = "������ϣ�"
End If
End Sub

Private Sub Timer1_Timer()
If Image1.Left > 9600 Then Image1.Left = -5780
Image1.Left = Image1.Left + 50
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, Url As Variant)

If (pDisp Is WebBrowser1.Object) Then
Label2 = "���ڻ�ȡ�������ƣ������ĵȴ�..."
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
MsgBox "û���ҵ��κν��"
Frame3.Enabled = 1
Option1.ForeColor = &H8000000D: Option2.ForeColor = &H8000000D: Option3.ForeColor = &H8000000D
End If
Command6.Visible = True

End Sub


Private Sub Command10_Click()
Command2.Enabled = False: Command3.Enabled = False: Command4.Enabled = False: Command10.Enabled = False
If List1.SelCount = 0 And List4.SelCount = 0 Then  'IF1
MsgBox "û��ѡ�е���Ŀ��˫��Ӧ�����ƻ�ѡ����ǰ�ķ��򼴿�ѡ����Ŀ"
'----------------------------------------------------------------------------------
Else
Shell "cmd /c adb decices"
Image1.Visible = True
Timer1.Enabled = True
Dim Savetime As Single
f = MsgBox("�Ƽ�ʹ���������Ӵ��䡣��ʹ�����������뱣���ֻ�������������ʧ���ʼ��ߣ�" & vbCrLf & "���ڣ����Ƿ�Ҫ��ʼ��ȡӦ�ã�", 4 + 48, "��ʾ")
If f = vbYes Then
Command11.Visible = True
Dim i, t, h, e As Integer
Dim y, z As String
h = 0
Label2 = "���ڻ�ȡӦ��·��..."
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
Open Text2 & "\appname" & h & ".txt" For Append As #1 'û���½�
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
Open Text2 & "\appname" & h & ".txt" For Append As #1 'û���½�
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

'---------------------------------------------------��ȡ��·��
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
'------------------------------------���ļ�
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
Label2 = "����׼����ȡ" & strData
Shell "cmd /c adb pull " & Pa

If Text5 = "" Then                 'IF A1
MsgBox strData & "�޷�����ȡ�����ȷ����ȡ��һ��Ӧ��"
Text3 = Text3 & strData & "��ȡʧ�ܣ�"
'________________________________________________________________________������ȡ
Else            'else A1

o = 0          '����׼����ȡ
Dim Filepath$
Dim yz
Do

Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
o = o + 1
Label2 = "����ʱ" & o & "�� ����׼����ȡ " & strData
Loop Until Dir(App.Path & "\" & Text5) <> "" Or o > 15

If o > 14 Then
MsgBox strData & "��ʱ��δ��ʼ���䣬�ƺ��޷�����ȡ�����ȷ����ȡ��һ��Ӧ��"
Text3 = Text3 & strData & "����ʧ�ܣ�" & vbCrLf
End If
    'IF A2
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Do          '��ȡ��
Label2 = "������ȡ�������ĵȴ�  " & strData
yz = FileLen(App.Path & "\" & Text5)
Savetime = Timer
While Timer < Savetime + 0.5
DoEvents
Wend
Loop Until yz > 0

Label2 = strData & "��ȡ�ɹ�����װ��������" & App.Path
Text3 = Text3 & strData & "��ȡ�ɹ���"

'-----------------------------------׼��������
n = InStr(strData, ":")
If n > 0 Then                                'IF A3
z = Mid(strData, 1, n - 1)
ElseIf n = 0 Then
If Text5 = "base.apk" Then
z = "base" & e
Text4 = Text4 & strData & "����Ϊ" & z & ".apk" & vbCrLf
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
Label2 = "�ȴ�������" & z
Loop
'-----------------------------------������
If Dir(App.Path & "\" & Text5) <> "" Then  'if A4
Name App.Path & "\" & Text5 As App.Path & "\" & z & ".apk"
Else
MsgBox z & "������ʧ��"
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
For d = 1 To h              'ɾ��
Kill Text2 & "\appname" & d & ".txt"
Kill Text2 & "\path" & d & ".txt"
Next d

'-------------------------------------���
Label2 = "��ȡ���"
MsgBox Text3 & "��ȡ��Ӧ���ѱ�����" & App.Path & "���뼰ʱ�������ǣ�"
Text3 = ""
If Text4 <> "" Then
MsgBox "��ע�⣬���ڰ�װ��û�����֣�" & vbCrLf & Text4, 0, "��Ҫ��ʾ��"
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
MsgBox "û��ѡ�е���Ŀ��˫��Ӧ�����ƻ�ѡ����ǰ�ķ��򼴿�ѡ����Ŀ"
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
Shell "cmd /c adb shell pm disable-user " & z & ">" & Text2 & "\��ʱ" & h & ".txt"
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
Shell "cmd /c adb shell pm disable-user " & z & ">" & Text2 & "\��ʱ" & h & ".txt"
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
    strFile = Text2 & "\��ʱ" & e & ".txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile
If InStr(strData, "Package") = 0 Then      'IF5
b = b + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " ʧ��"
ElseIf InStr(strData, "Package") = 1 Then
m = m + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " �ɹ�"
End If       'END5
Kill Text2 & "\��ʱ" & e & ".txt"
Next e
MsgBox m & "���ɹ���" & b & "��ʧ�ܣ�" & Text3
If Option2.Value = True Then Call Option2_Click
End If      'END1
List3.Clear
Text3 = ""
Command2.Enabled = 1: Command3.Enabled = 1: Command4.Enabled = 1
End Sub

Private Sub Command3_Click()
Command2.Enabled = False: Command3.Enabled = False: Command4.Enabled = False
If List1.SelCount = 0 And List4.SelCount = 0 Then  'IF1
MsgBox "û��ѡ�е���Ŀ��˫��Ӧ�����ƻ�ѡ����ǰ�ķ��򼴿�ѡ����Ŀ"
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
Shell "cmd /c adb shell pm enable " & z & ">" & Text2 & "\��ʱ" & h & ".txt"
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
Shell "cmd /c adb shell pm enable " & z & ">" & Text2 & "\��ʱ" & h & ".txt"
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
    strFile = Text2 & "\��ʱ" & e & ".txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile
If InStr(strData, "Package") = 0 Then      'IF5
b = b + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " ʧ��"
ElseIf InStr(strData, "Package") = 1 Then
m = m + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " �ɹ�"
End If       'END5
Kill Text2 & "\��ʱ" & e & ".txt"
Next e
MsgBox m & "���ɹ���" & b & "��ʧ�ܣ�" & Text3
If Option3.Value = True Then Call Option3_Click
End If      'END1
List3.Clear
Text3 = ""
Command2.Enabled = 1: Command3.Enabled = 1: Command4.Enabled = 1
End Sub

Private Sub Command4_Click()
Command2.Enabled = False: Command3.Enabled = False: Command4.Enabled = False
If List1.SelCount = 0 And List4.SelCount = 0 Then  'IF1
MsgBox "û��ѡ�е���Ŀ��˫��Ӧ�����ƻ�ѡ����ǰ�ķ��򼴿�ѡ����Ŀ"
Else
K = MsgBox("��ȷ��ж��ѡ�е�Ӧ����ǿ�ҽ�����ж��ǰ����ȡ���ݣ�", 1 + 48, "����")
If K = 1 Then
Dim i, t, h, e As Integer
Dim y, z As String
Label2 = "��������ж�أ������ĵȴ�..."
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
Shell "cmd /c adb shell pm uninstall --user 0 " & z & ">" & Text2 & "\��ʱ" & h & ".txt"
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
Shell "cmd /c adb shell pm uninstall --user 0 " & z & ">" & Text2 & "\��ʱ" & h & ".txt"
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
    strFile = Text2 & "\��ʱ" & e & ".txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile
If InStr(strData, "Su") = 0 Then      'IF5
b = b + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " ʧ��"
ElseIf InStr(strData, "Su") = 1 Then
m = m + 1
Text3 = Text3 & vbCrLf & List3.List(e - 1) & " �ɹ�"
End If       'END5
Next e
Label2 = "ж�����"
MsgBox m & "���ɹ���" & b & "��ʧ�ܣ�" & Text3, 0, "����жϿ������������ֻ�Ϊ׼"
For e = 1 To h
If Dir(Text2 & "\��ʱ" & e & ".txt") <> "" Then
'Kill Text2 & "\��ʱ" & e & ".txt"
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
List1.AddItem "����δ�������磬���޷���ȡ����"
Shell "cmd /c adb shell pm list packages" & ">" & Text2 & "\all.txt"
If Dir(Text2 & "\all.txt") <> "" Then '�У�
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
MsgBox "���س�ʱ���޷���ȡӦ�����ģ�����������������", 0 + 48, "��ȡ����ʧ��"
End If
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
i = i + 1
Label2 = "���ڻ�ȡ�������ƣ�����ʱ" & i & "��"
Loop

If Dir(Text2 & "\in.txt") <> "" Then
Else
Open Text2 & "\in.txt" For Append As #1 'û���½�
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
V = Mid(S, 9)      '�ֻ�Ӧ��
For i = 0 To List2.ListCount - 1  '�����б�
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

Label2 = "������ϣ�"
Command1.Enabled = True: Command2.Enabled = True: Command3.Enabled = True: Command4.Enabled = True: Command6.Enabled = True: List1.Enabled = True:
Command5.Enabled = True: Command8.Enabled = True: Command9.Enabled = True: Command10.Enabled = True: Check2.Enabled = True: Frame3.Visible = True
Option1.Enabled = 1
Option1.Value = True
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Label2 = "����ϵͳ�ؼ�Ӧ���޷���ͣ��ж�أ��磺�ֻ��ܼң���Ϊ���棬����ʹ���ֻ��ȣ����������������"
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
