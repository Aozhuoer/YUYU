VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00FFFFFF&
   Caption         =   "��װӦ�õ��ֻ�"
   ClientHeight    =   4635
   ClientLeft      =   2790
   ClientTop       =   1275
   ClientWidth     =   4725
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4725
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3360
      Top             =   3960
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0ED9C&
      Caption         =   "��װ"
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
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   480
      Left            =   120
      TabIndex        =   1
      Text            =   "�ڴ������ļ�·��"
      Top             =   3000
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006DC725&
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   80
      Left            =   -4000
      Picture         =   "Form13.frx":74CA
      Stretch         =   -1  'True
      Top             =   2910
      Width           =   4000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���� D:\YUYU\666.apk"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   1920
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If InStr(Text2, ".apk") > 0 Then
Text1 = "����׼����װ..."
Command1.Enabled = 0: Timer1.Enabled = 1: Image1.Visible = True
If Dir(Text3 & "\apk.txt") <> "" Then Kill Text3 & "\apk.txt"
Shell "cmd /c adb install " & Text2 & ">" & Text3 & "\apk.txt"
Text1.Enabled = 0: Text2.Enabled = 0
Do Until Dir(Text3 & "\apk.txt") <> ""
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Text1 = "��������..."
Loop

Do
Text1 = "���ڰ�װ�����ĵȴ�..."
Savetime = Timer
While Timer < Savetime + 0.5
DoEvents
Wend
yz = FileLen(Text3 & "\apk.txt")
Loop Until yz > 0

i = 0
Open Text3 & "\apk.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
i = i + 1
If i = 2 Then Exit Do
Loop
Timer1.Enabled = 0: Image1.Visible = 0
If InStr(S, "Su") > 0 Then
MsgBox "��װ�ɹ���"
Else
MsgBox "�����޷��ж��Ƿ�ɹ��������ֻ��鿴"
End If
Close #1
Text1.Enabled = 1: Text2.Enabled = 1
Text1 = "��apk�ϵ��˴�" & vbCrLf & "�����·�����apk��ŵ�·��" & vbCrLf & "���ɰ�װӦ�õ��ֻ���"
Text2 = ""
Command1.Enabled = 1
ElseIf Text2 = "" Or InStr(Text2, ".apk") = 0 Then
MsgBox "û�а�װ����", 0 + 48, "�޷���ʼ"
End If
End Sub

Private Sub Form_Load()
Text3 = Form1.Text5
Text1 = vbCrLf & "��apk�ϵ��˴�" & vbCrLf & "�����·�����apk��ŵ�·��" & vbCrLf & "���ɰ�װӦ�õ��ֻ���"
End Sub


Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Data.GetFormat(vbCFFiles) = True Then '�ж��Ƿ�Ϊ�ļ�bai����
FileName = Data.Files(1) '����ļ���
Text1.Text = FileName
Text2.Text = Text1.Text
End If
End Sub

Private Sub Text2_Click()
Text2 = ""
End Sub

Private Sub Timer1_Timer()
If Image1.Left > 4965 Then Image1.Left = -4000
Image1.Left = Image1.Left + 30
End Sub
