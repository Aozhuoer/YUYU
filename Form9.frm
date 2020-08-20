VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FFFFFF&
   Caption         =   "设定IP"
   ClientHeight    =   2760
   ClientLeft      =   7830
   ClientTop       =   1935
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4230
   Begin VB.CommandButton Command2 
      BackColor       =   &H00F5F1E7&
      Caption         =   "保存"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "修改"
      Height          =   495
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   450
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "如何查看手机IP？"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1425
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "当前IP："
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
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "例如：192.168.1.2"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "若您的手机连接路由的IP地址有变化，可在此更改IP地址，然后按“重试”连接。"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Form1.Text1 = Text1
Kill Form1.Text5 & "\IP1.txt"
Command1.Enabled = True
Text1.Enabled = False
Command2.Enabled = False
Open Form1.Text5 & "\IP1.txt" For Append As #1
Print #1, Text1
Close #1
End Sub

Private Sub Form_Load()
Text1 = Form1.Text1
If Text1 = "" Then
Label1 = "请输入手机IP地址后保存，然后再次点击“建立无线连接”"
Label3 = "在此输入"
Text1.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
End If
End Sub

Private Sub Label4_Click()
Shell "explorer https://v.xiumi.us/board/v5/3JD5p/215117062"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command2_Click
End Sub
