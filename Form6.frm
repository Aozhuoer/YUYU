VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�ռ�����"
   ClientHeight    =   6765
   ClientLeft      =   2280
   ClientTop       =   2955
   ClientWidth     =   8760
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   8760
   Begin VB.TextBox Text6 
      Height          =   975
      Left            =   9120
      TabIndex        =   7
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   975
      Left            =   9120
      TabIndex        =   6
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   975
      Left            =   9120
      TabIndex        =   5
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   9120
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   9120
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�򿪡��ռ�����.txt��"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "�ռ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "�������䣺1483544237@qq.com"
      Top             =   6120
      Width           =   3735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�鿴��ϸ�̳�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   6485
      TabIndex        =   10
      Top             =   200
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5280
      TabIndex        =   9
      Top             =   5520
      Width           =   60
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ӵ�����صĽ��汻�㷢����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   8
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   120
      Picture         =   "Form6.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8460
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label2 = "�����ռ���Ϣ�����Ժ�..."
Shell "cmd /c adb shell getprop ro.product.model" & ">" & "d:\�ͺ�.txt"
Shell "cmd /c adb shell getprop ro.build.version.release" & ">" & "d:\�汾.txt"
Shell "cmd /c tasklist" & ">" & "d:\����.txt"
Dim Savetime As Single
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Dim strFile  As String
    Dim intFile  As Integer
    Dim strData  As String
    strFile = "d:\�ͺ�.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile
Text2.Text = strData
    Dim sData  As String
    strFile = "d:\�汾.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    sData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print sData
    Close intFile
Text2.Text = Text2.Text + vbCrLf + sData
    Dim stData  As String
    strFile = "d:\����.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    stData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print stData
    Close intFile
Text2.Text = Text2.Text + vbCrLf + stData

Text3.Text = Date
Text4.Text = VBA.Environ("computername")
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
For Each objItem In colItems
DoEvents
strosversion = objItem.version
Next
strosversion = Left(strosversion, 3)
Text5.Text = strosversion
Text6.Text = App.Path
Text2.Text = Text2.Text + vbCrLf + Text3.Text + vbCrLf + Text4.Text + vbCrLf + Text5.Text + vbCrLf + Text6.Text + vbCrLf + Form1.Text1
Dim strs As String
If Dir("d:\�ռ�����.txt") <> "" Then Kill "d:\�ռ�����.txt"
Open "d:\�ռ�����.txt" For Append As #1
     strs = Text2.Text
    Write #1, strs
    Close #1

Kill "d:\����.txt"
Kill "d:\�ͺ�.txt"
Kill "d:\�汾.txt"
MsgBox "�ռ���ɣ�"
Command2.Enabled = True
Label2 = ""
End Sub

Private Sub Command2_Click()
Shell "explorer d:\�ռ�����.txt"
End Sub

Private Sub Label3_Click()
Shell "explorer https://v.xiumi.us/board/v5/3JD5p/203664485"
End Sub
