VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "YUYU����v1.9"
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
      Caption         =   "���������ٶ�"
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
      Caption         =   "��װӦ�õ��ֻ�"
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
      Caption         =   "�鿴����Ӧ��"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�趨�ֻ�IP"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��ݲ���"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Caption         =   "�����Զ��嵥Ӧ��"
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
         Caption         =   "����������ͣ��Ӧ��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ͣ���Զ��嵥Ӧ��"
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
         Caption         =   "�Զ���Ӧ���嵥"
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
      Caption         =   "��������"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������������"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "������������"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "ǿ�ƿ���"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC671&
      Caption         =   "����"
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
      Caption         =   "�˳�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�˳���������adb"
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
      Caption         =   "�ռ�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��cmd����"
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
      Height          =   300
      Left            =   3480
      TabIndex        =   40
      Top             =   6480
      Width           =   1275
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�ҵ��ң�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�ᰲ"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      Caption         =   "���۾��ֲ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�������adb�˿�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��������ʧ�ܣ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�鿴��ϸ�̳�"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�����������ֵ������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�Ͽ���������"
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
      Caption         =   "������. . ."
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "���๦��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "����Ӧ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
MsgBox "���������ӽ�ʹ�������ڵ������豸�����ܿ��������豸����ȷ�Ͼ������豸���䰲װ���������ȫ���ٿ���", 1 + 48, "����"
Dim strFile  As String
Dim intFile  As Integer
Label4 = "���ڽ�����������. . ."
Shell "cmd /c adb devices"
Shell "cmd /c adb tcpip 5555"
If Text1 = "" Then  'F1
Form9.Show
Label4 = "û�����Ӽ�¼�������趨�ֻ�IP��ַ"
Else          'else
Label4 = Label4 & vbCrLf & "������������" & Text1
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
Label4 = "������������ʧ��!"
Call Command7_Click
ElseIf Text3.Text = Left(strDa, 2) Or Text4.Text = Left(strDa, 2) Then
Label4 = "�����������ӳɹ���" & vbCrLf & "���������ε������ߺ��������ԡ���ť"
Shell "cmd /c adb disconnect "
Command11.Enabled = False
Command11.Caption = "�밴 ���� ��ť"
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
Label4 = "�ȴ�adb��������. . . " & vbCrLf & "�����������������ӳɹ�Ŷ"
Savetime = Timer
While Timer < Savetime + 2
DoEvents
Wend

Label4 = Label4 & vbCrLf & "�����������. . . "
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
Label4 = Label4 & vbCrLf & "�����������. . . "
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
Label4 = "��������ʧ��!"
If InStr(strData, "-") = 0 Then '��������Ҳʧ��    'F2
Label4 = Label4 & vbCrLf & "��������Ҳʧ��!" & vbCrLf & strDa & "û�������ӵ��豸��������������������"
MsgBox "û�м�鵽�����ӵ��豸��", 0 + 48, "�豸δ����"
If Dir(Text5 & "\adb.txt") <> "" Then d = 1
ElseIf InStr(strData, "-") > 0 Then '�������ӳɹ�
Label4 = Label4 & vbCrLf & "�������ӳɹ���" & vbCrLf & "��ӭ�㣺" & strData
Command11.Enabled = 1
If InStr(Mid(sr, 3, 1), "1") > 0 Then Form11.Show
Call Command2_Click
End If    'end1
ElseIf Text3.Text = Left(strDa, 2) Or Text4.Text = Left(strDa, 2) Then
If InStr(strData, "-") = 0 Then  'f3
Label4 = "�������ӳɹ�������ͬʱ�����������ߣ�" & vbCrLf & "��Ҫ���������� �� �������Ϊ�������ӡ���" & vbCrLf & "��������ִ�в�����(��ʶ��������������ԡ�)"
ElseIf InStr(strData, "-") > 0 Then
Label4 = "�������ӳɹ���" & vbCrLf & strDa & "��ӭ�㣺" & strData
Call Command2_Click
Label7.Visible = True
If InStr(Mid(sr, 3, 1), "1") > 0 Then Form11.Show
End If  'end2
Command6.Caption = "��Ϊ��������"
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
m = MsgBox("ȷ��ͣ������Ӧ����" & n, 1, "ͣ��Ӧ��")
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
Label4 = "��ͣ�������Զ����嵥��Ӧ��"
MsgBox "��ͣ��" & n
Else
Label4 = "��ȡ��"
End If
Else
MsgBox "����δ�Զ���Ӧ���嵥��"
End If
End Sub

Private Sub Command18_Click()
Label4 = "���Ժ�..."
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
MsgBox "��ǰû����ͣ�õ�Ӧ�ã�"
Label4 = "��ǰû����ͣ�õ�Ӧ��"
Else
m = MsgBox("ȷ����������Ӧ����" & n, 1, "����������ͣ��Ӧ��")
If m = 1 Then
Open Text5 & "\disable.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, S
V = Mid(S, 9)
Shell "cmd /c adb shell pm enable " & V
Loop
Close #1
Label4 = "����������ͣ�õ�Ӧ��"
MsgBox "������" & n
Else
Label4 = "��ȡ��"
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
m = MsgBox("ȷ����������Ӧ����" & n, 1, "����Ӧ��")
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
Label4 = "�����������Զ����嵥��Ӧ��"
MsgBox "������" & n
Else
Label4 = "��ȡ��"
End If
Else
MsgBox "����δ�Զ���Ӧ���嵥��"
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

Private Sub Command21_Click() '����
d = 0
Command6.Enabled = False: Command11.Enabled = False: Command21.Enabled = False:
Image2.Left = -6240
Timer1.Interval = 28
Timer1.Enabled = True
Image2.Visible = True
Label4 = "�����������. . . "
Shell "cmd /c adb connect " & Text1 & ">" & Text5 & "\IPC.txt"
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Shell "cmd /c adb shell getprop ro.product.model" & ">" & Text5 & "\YUYU.txt"
Label4 = Label4 & vbCrLf & "�����������. . . "
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
Label4 = "��������ʧ��!"
If InStr(strData, "-") = 0 Then '��������Ҳʧ��    'F2
Label4 = Label4 & vbCrLf & "��������Ҳʧ��!" & vbCrLf & strDa & "û�������ӵ��豸���������鿴��ϸ�̳�"
Call Command7_Click
MsgBox "û�м�鵽�����ӵ��豸���������鿴��ϸ�̳�", 0 + 48, "�豸δ����"
d = 1
ElseIf InStr(strData, "-") > 0 Then '�������ӳɹ�
Label4 = Label4 & vbCrLf & "�������ӳɹ���" & vbCrLf & "��ӭ�㣺" & strData
Command11.Enabled = 1
If InStr(Mid(sr, 3, 1), "1") > 0 Then Form11.Show
Call Command2_Click
End If    'end1
ElseIf Text3.Text = Left(strDa, 2) Or Text4.Text = Left(strDa, 2) Then
If InStr(strData, "-") = 0 Then  'f3
Label4 = "�������ӳɹ�������ͬʱ�����������ߣ�" & vbCrLf & "��Ҫ���������߻�����ť��Ϊ�������ӣ�" & vbCrLf & "��������ִ�в�����"
Call Command7_Click
ElseIf InStr(strData, "-") > 0 Then
Label4 = "�������ӳɹ���" & vbCrLf & strDa & "��ӭ�㣺" & strData
Call Command2_Click
If InStr(Mid(sr, 3, 1), "1") > 0 Then Form11.Show
Label7.Visible = True
End If  'end2
Command6.Caption = "��Ϊ��������"
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
Label4 = "�ѽ���adb����"
Dim Savetime As Single
Savetime = Timer
While Timer < Savetime + 0.5
DoEvents
Wend
End
End If
End Sub

Private Sub Command6_Click() '��Ϊ����
d = 0
Command6.Caption = "������������"
Command6.Enabled = False: Command11.Enabled = False: Command21.Enabled = False:
Image2.Left = -6240
Timer1.Enabled = True
Image2.Visible = True
Dim Savetime As Single
Label4 = "��Ϊ��������. . . "
Shell "cmd /c adb disconnect " & Text1
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
Shell "cmd /c adb shell getprop ro.product.model" & ">" & Text5 & "\YUYU.txt"
Label4 = Label4 & vbCrLf & "�����������. . . "
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

If InStr(strData, "-") = 0 Then '��������ʧ��
Label4 = "��������ʧ��!" & vbCrLf & strDa & "û�м�鵽�������ӵ��豸"
Call Command7_Click
d = 1
MsgBox "û�м�鵽�����ӵ��豸���������鿴��ϸ�̳�", 0 + 48, "�豸δ����"
ElseIf InStr(strData, "-") > 0 Then '�������ӳɹ�
Label4 = "�������ӳɹ���" & vbCrLf & "��ӭ�㣺" & strData
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
MsgBox "����Ĭ�ϴ洢��" & p & "�������и���λ�á�", 0, "��ӭʹ��YUYU���֣�"
Shell "explorer https://c.xiumi.us/board/v5/3JD5p/215117062"
End If
Text5 = p

Shell "cmd /c adb devices" & ">" & Text5 & "\dev.txt" '����adb
WebBrowser1.navigate "http://www.wwnote.xyz/ao/check.html"
'WebBrowser1.navigate "https://v.xiumi.us/board/v5/3JD5p/245146051" '��ʱͣ��
If Dir(Text5 & "\set.txt") <> "" Then
Open Text5 & "\set.txt" For Input As #1
Input #1, sr
Close #1
Else
Open Text5 & "\set.txt" For Append As #1 'û��ip�½�
Close #1
Open Text5 & "\set.txt" For Output As #1 'д��

sr = "010"
Print #1, sr
Close #1
End If

If Dir(Text5 & "\IP1.txt") <> "" Then
Text1 = ""
Else
Open Text5 & "\IP1.txt" For Append As #1 'û��ip�½�
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
Line Input #1, strTemp    '����һ������
Text1 = strTemp
Close #1
End If
Call Command12_Click  '����

If InStr(Mid(sr, 1, 1), "1") > 0 Then
Call Command2_Click
Label4 = "Ĭ�����ӳɹ�ģʽ���ѽ�����ť"
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

If u > 1.9 And u < 3 Then                   '��ֵ������
Label14.ForeColor = vbRed
Label14 = "���°棡"
If InStr(Mid(sr, 2, 1), "1") > 0 Then
m = MsgBox("�°汾v" & u & "���������Ƿ�ȥ����", vbYesNo, "���°汾����")
If m = vbYes Then Shell "explorer https://v.xiumi.us/board/v5/3JD5p/228148533"
End If
End If
Dim v_Path As String, K As Long, e, v_Range
v_Path = Text5 & "\ad1"
WebBrowser1.Silent = True '�رս�������ֹ�ű�����
For Each e In WebBrowser1.Document.All
    If e.tagName = "IMG" Then
        Set v_Range = WebBrowser1.Document.body.createControlRange()
        v_Range.Add e
        v_Range.execCommand "Copy" '���Ƶ�������
        K = K + 1
        SavePicture Clipboard.GetData, v_Path & ".jpg" '���浽Ӳ��
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
m = MsgBox("�����Ҫ�˳���", vbExclamation + vbYesNo + vbDefaultButton2, "�˳�")
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
w = MsgBox("����ȷ�����Ѿ��Թ���Ա������б�����", 1 + 48)
If w = 1 Then
Randomize
Dim m As Integer
m = Int(Rnd * 200 + 7800)
Dim n As String
n = """" & m & """"
Shell "cmd /c setx /M ANDROID_ADB_SERVER_PORT " & n
MsgBox "�ѽ�adb�˿ڸ���Ϊ" & n & "������������ʹ����������Ч"
Label4 = "���������Ժ��ٳ������ӣ�"
Else
Label4 = "���Ҽ��Թ���Ա������б�����"
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
Label13 = "ע����������"
Savetime = Timer
While Timer < Savetime + 4
DoEvents
Wend
Label13 = "��cmd����"
End Sub

Private Sub Label14_Click()
Label4 = "���ڼ�����"

u = Mid(Text6, 1, 3)
If u > 1.9 Then           '��ֵ������
m = MsgBox("�°汾v" & u & "���������Ƿ�ȥ����", vbYesNo, "���°汾����")
Label4 = "���°汾��������"
Label14.ForeColor = vbRed
Label14 = "�а汾����"
If m = vbYes Then Shell "explorer https://v.xiumi.us/board/v5/3JD5p/228148533"
Else
MsgBox "�����°汾"
Label4 = "�����°汾"
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
Label7 = "���ڶϿ�"
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
Label4 = "�ѶϿ��������� " & strData
Label7.Enabled = True
Label7 = "�Ͽ���������"
End Sub



Private Sub Timer1_Timer()
Image2.Left = Image2.Left + 44
End Sub

