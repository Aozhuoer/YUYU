VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form7 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�Զ���Ӧ���嵥"
   ClientHeight    =   5250
   ClientLeft      =   7725
   ClientTop       =   1260
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7080
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1695
      Left            =   10680
      TabIndex        =   15
      Top             =   3600
      Width           =   3375
      ExtentX         =   5953
      ExtentY         =   2990
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
      Location        =   ""
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   495
      Left            =   9360
      TabIndex        =   14
      Top             =   3480
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3360
      Left            =   7320
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   120
      Width           =   6615
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   3615
      Left            =   7320
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4080
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���"
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�޸�"
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0ED9C&
      Caption         =   "����"
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
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ֶ����"
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ɾ��"
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����ֻ���ͣ��Ӧ��"
      Height          =   375
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006DC725&
      Height          =   300
      Left            =   6240
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ȷ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006DC725&
      Height          =   300
      Left            =   5640
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005809FF&
      Height          =   300
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   960
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim u, t
Private Sub Command1_Click()
If u = 0 Then '����
Label1 = "�Ե�һ��..."
Call Command7_Click
t = 1
ElseIf u = 1 Then
m = MsgBox("��ǰ����Ӧ���嵥�����뽫���ǣ���ȷ��Ҫ������", 1 + 48, "���Ǿ���")
If m = 2 Then
Cancel = True
ElseIf m = 1 Then '����
Label1 = "�Ե�һ��..."
Call Command7_Click
Label1 = "�����������水ťֱ�ӹرմ��ڣ��򲻻Ḳ��֮ǰ���嵥"
t = 1
End If
End If
End Sub

Private Sub Command2_Click()
Dim i As Integer
If List1.SelCount = 1 Then
Text2 = List1.Text
Label1 = "���ڴ˴��޸�:"
Text2.Visible = 1: Label3.Visible = True: Label4.Visible = True
Command5.Enabled = False: Command6.Enabled = False: Command4.Enabled = False: Command3.Enabled = False: Command1.Enabled = False
ElseIf List1.SelCount > 1 Then
MsgBox "����ͬʱ�޸Ķ���", 0 + 48, "��ǰѡ���˶���"
ElseIf List1.SelCount = 0 Then
MsgBox "û��ѡ�е���Ŀ��"
End If
End Sub

Private Sub Command3_Click()
Dim i As Integer '�ж��б���Ƿ�ֻ��һ����Ŀ��ѡ��
If List1.SelCount = 0 Then
MsgBox "û��ѡ�е���Ŀ��"
ElseIf List1.SelCount = 1 Then
List1.RemoveItem List1.ListIndex
Label1 = "��ɾ��"
ElseIf List1.ListCount > 1 Then  'ɾ���б���е���ѡ�еĶ����Ŀ
For i = List1.ListCount - 1 To 0 Step -1  'ListCount�����б���е���Ŀ����
'ListCount-1���б�������һ����Ŀ��������
'�жϸ���Ŀ�Ƿ�ѡ�У�Selected()���ز���ֵ
If List1.Selected(i) Then  'ɾ��������Ϊi����Ŀ
List1.RemoveItem i
End If
Next
Label1 = "��ɾ��"
End If
t = 1
End Sub

Private Sub Command4_Click()
Dim n As String
n = InputBox("������Ӧ�ð���" & vbCrLf & vbCrLf & "����Ӧ�ð���ǰ��������ע�Ͳ���Ӣ��ð�ŷָ�" & vbCrLf & vbCrLf & "���磺�ǻ�����:com.huawei.search")
List1.AddItem n
Label1 = "�����"
Command2.Visible = 1: Command3.Visible = 1: Command6.Visible = 1: Command5.Visible = 1
t = 1
End Sub

Private Sub Command5_Click()
Dim i As Integer
Label1 = "���ڱ���..."
Text3 = List1.List(0)
For i = 1 To List1.ListCount - 1
Text3 = Text3 & vbCrLf & List1.List(i)
Next
If Dir(Text1 & "\List1.txt") <> "" Then Kill Text1 & "\List1.txt"
Open Text1 & "\List1.txt" For Append As #1 'û���½�
Close #1

Open Text1 & "\List1.txt" For Output As #1
Print #1, Text3
Close #1
Label1 = "�ѱ��棡"
MsgBox "����ɹ���"
t = 0
Form1.Command19.Enabled = 1: Form1.Command17.Enabled = 1
End Sub

Private Sub Command6_Click()
m = MsgBox("��ȷ��Ҫ����б���", 1 + 48, "����")
If m = 2 Then
Cancel = True
ElseIf m = 1 Then
List1.Clear
Label1 = "������б�"
End If
t = 1
End Sub

Private Sub Command7_Click()
Shell "cmd /c adb shell pm list packages -s -d" & ">" & Text1 & "\disable.txt"

Dim x, S, V, A, y As String
WebBrowser1.navigate "https://v.xiumi.us/board/v5/3JD5p/220492873"
Savetime = Timer
While Timer < Savetime + 2
DoEvents
Wend

x = WebBrowser1.Document.body.innerText
If Dir(Text1 & "\in.txt") <> "" Then
Else
Open Text1 & "\in.txt" For Append As #1 'û��ip�½�
Close #1
End If
Open Text1 & "\in.txt" For Output As #1
Print #1, x
Close #1
Open Text1 & "\in.txt" For Input As #2
Do While Not EOF(2)
Line Input #2, A
List2.AddItem A
Loop
Close #2
List1.Clear
Open Text1 & "\disable.txt" For Input As #1
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

If S = "" Then
MsgBox "û���κ���ͣ�õ�Ӧ�ã�����ȥͣ��Ӧ��"
Label1 = "û���κ���ͣ�õ�Ӧ�ã�����ȥͣ��Ӧ�ã�"
Else
Label1 = "�ѵ��룡"
Command2.Visible = True: Command3.Visible = True: Command6.Visible = True: Command5.Visible = True
End If
End Sub

Private Sub Label3_Click()
List1.List(List1.ListIndex) = Text2
Text2.Visible = 0: Label3.Visible = 0: Label4.Visible = 0
Command5.Enabled = 1: Command6.Enabled = 1: Command4.Enabled = 1: Command3.Enabled = 1: Command1.Enabled = 1
Label1 = "���޸�"
t = 1
End Sub

Private Sub Label4_Click()
Text2.Visible = 0: Label3.Visible = 0: Label4.Visible = 0
Command5.Enabled = 1: Command6.Enabled = 1: Command4.Enabled = 1: Command3.Enabled = 1: Command1.Enabled = 1
Label1 = "��ȡ��"
End Sub

Private Sub Form_Load()
Text1 = Form1.Text5
t = 0
If Dir(Text1 & "\List1.txt") <> "" Then '�У�
Label1 = "���ڵ������ϴα����Ӧ���嵥..."
Dim S As String
Dim iLine As Integer
iLine = 1
Open Text1 & "\List1.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
iLine = iLine + 1
List1.AddItem S
Loop
Close #1
u = 1
Label1 = "�ѵ������ϴα����Ӧ���嵥"
If S = "" Then
List1.Clear
u = 0
Kill Text1 & "\List1.txt"
End If
Else
u = 0
End If
If u = 0 Then
Command2.Visible = 0: Command3.Visible = 0: Command6.Visible = 0: Command5.Visible = 0
Label1 = "��Ŀǰû���趨Ӧ���嵥"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If t = 1 Then
m = MsgBox("�Ƿ񲻱��沢�˳���", vbExclamation + vbYesNo + vbDefaultButton2, "δ����")
If m = vbNo Then
Cancel = True
ElseIf m = vbYes Then
Unload Me
End If
End If
End Sub

