VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ȸ�ȫ��Ͱ"
   ClientHeight    =   7950
   ClientLeft      =   7755
   ClientTop       =   1185
   ClientWidth     =   5115
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   5115
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5280
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command25 
      Caption         =   "25"
      Height          =   375
      Left            =   4320
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command24 
      Caption         =   "ж��"
      Height          =   375
      Left            =   3960
      TabIndex        =   29
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command23 
      Caption         =   "ж��"
      Height          =   375
      Left            =   3960
      TabIndex        =   28
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "ͣ��"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "ͣ��"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ͣ��"
      Height          =   375
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Caption         =   "ж��"
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "ж��"
      Height          =   375
      Left            =   3960
      TabIndex        =   18
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "ж��"
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "ж��"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "ͣ��"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "ͣ��"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "ͣ��"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "ͣ��"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ж��"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ж�غ�ֻ��ͨ���ָ��������ûָ�Ӧ�ã������ж�أ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   180
      Left            =   240
      TabIndex        =   33
      Top             =   7560
      Width           =   4680
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "����Ҳ��֪����ɶ��"
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   0
      TabIndex        =   32
      Top             =   6600
      Width           =   1620
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����ǰ���һ�£�"
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   120
      TabIndex        =   31
      Top             =   3840
      Width           =   1440
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ȸ�play����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ȸ�AR����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Google One Time Init"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ȸ豸�ݴ���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ȸ������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "play�̵�(�ȸ������³���)"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�� �� ȫ �� Ͱ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   405
      Left            =   960
      TabIndex        =   2
      Top             =   1605
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ж�غ󲿷������ȸ�����Ӧ����Ϸ�����޷�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   120
      Picture         =   "Form2.frx":74CA
      Top             =   0
      Width           =   4830
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.google.android.gms" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command10_Click()
Shell "cmd /c adb shell pm enable com.google.android.backuptransport" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command11_Click()
Shell "cmd /c adb shell pm enable com.google.android.partnersetup" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command12_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.android.vending" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command13_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.google.android.gsf" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command14_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.google.android.backuptransport" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command15_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.google.android.partnersetup" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command16_Click()
Shell "cmd /c adb shell pm disable-user com.google.ar.core" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command17_Click()
Shell "cmd /c adb shell pm disable-user com.google.android.onetimeinitializer" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command18_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.android.hwouc" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command19_Click()
Shell "cmd /c adb shell pm enable com.huawei.android.hwouc" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command2_Click()
Shell "cmd /c adb shell pm disable-user com.google.android.gms" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command20_Click()
Shell "cmd /c adb shell pm enable com.google.android.onetimeinitializer" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command21_Click()
Shell "cmd /c adb shell pm enable com.google.ar.core" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command22_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.android.hwouc" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command23_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.google.android.onetimeinitializer" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command24_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.google.ar.core" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command25_Click()
Dim Savetime As Single
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
  Dim strFile  As String
    Dim intFile  As Integer
    Dim strData  As String
    strFile = Text1 & "\YUYU��ʱ.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile

     If strData = "" Then
MsgBox "��Ӧ���޷���ͣ��ж�أ��������ֻ��޸�Ӧ��", 0 + 48, "ִ��ʧ��"
  Label1 = "ʧ��(T_T)"
 Else
        Label1 = "�ɹ�!"
        MsgBox "ִ�гɹ���" & strData, 0, "�ɹ�(�R���Q)"
        End If
Kill Text1 & "\YUYU��ʱ.txt"
End Sub

Private Sub Command3_Click()
Shell "cmd /c adb shell pm enable com.google.android.gms" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command4_Click()
Shell "cmd /c adb shell pm disable-user com.android.vending" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command5_Click()
Shell "cmd /c adb shell pm disable-user com.google.android.gsf" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command6_Click()
Shell "cmd /c adb shell pm disable-user com.google.android.backuptransport" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command7_Click()
Shell "cmd /c adb shell pm disable-user com.google.android.partnersetup" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command8_Click()
Shell "cmd /c adb shell pm enable com.android.vending" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Command9_Click()
Shell "cmd /c adb shell pm enable com.google.android.gsf" & ">" & Text1 & "\YUYU��ʱ.txt"
Call Command25_Click
End Sub

Private Sub Form_Load()
Text1 = Form1.Text5
Label5.Alignment = 2
Label1.Alignment = 2
End Sub

