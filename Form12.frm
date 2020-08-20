VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00FFFFFF&
   Caption         =   "更改数据存储位置"
   ClientHeight    =   2595
   ClientLeft      =   7950
   ClientTop       =   4875
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   7170
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F1E7&
      Caption         =   "更改"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "G盘"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "F盘"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E盘"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "D盘"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "C盘"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   0
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005809FF&
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim u As Integer
Dim p, n
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Sub Command1_Click()
Dim FileSys As New FileSystemObject
Dim FolderObj As Folder
If u = 1 Then
If Option1.Value = True Then n = "c:\YUYU"
If Option2.Value = True Then n = "d:\YUYU"
If Option3.Value = True Then n = "e:\YUYU"
If Option4.Value = True Then n = "f:\YUYU"
If Option5.Value = True Then n = "g:\YUYU"
If n = p Then
MsgBox "请勿选择相同盘符"
Else
FileSys.CopyFolder p, n
FileSys.DeleteFolder p
Dim FSO As New FileSystemObject
If FSO.FolderExists("c:\YUYU") Then
Option1.Value = True
Label1 = "当前数据存储在C盘，请重启程序"
p = "c:\YUYU"
ElseIf FSO.FolderExists("d:\YUYU") Then
Option2.Value = True
Label1 = "当前数据存储在D盘，请重启程序"
p = "d:\YUYU"
ElseIf FSO.FolderExists("e:\YUYU") Then
Option3.Value = True
Label1 = "当前数据存储在E盘，请重启程序"
p = "e:\YUYU"
ElseIf FSO.FolderExists("f:\YUYU") Then
Option4.Value = True
Label1 = "当前数据存储在F盘，请重启程序"
p = "f:\YUYU"
ElseIf FSO.FolderExists("g:\YUYU") Then
Option5.Value = True
Label1 = "当前数据存储在G盘，请重启程序"
p = "g:\YUYU"
End If
Form1.Text5 = p
u = 0
Command1.Visible = 0
Label1 = Label1 & vbCrLf & p
End If
End If
End Sub

Private Sub Form_Load()
Dim strSave As String
    Dim drvName As String
    'Set the graphic mode to persistent
    Me.AutoRedraw = True
    'Create a buffer to store all the drives
    strSave = String(255, Chr$(0))
    'Get all the drives
    ret& = GetLogicalDriveStrings(255, strSave)
    'Extract the drives from the buffer and print them on the form
    For keer = 1 To 100
        If Left$(strSave, InStr(1, strSave, Chr$(0))) = Chr$(0) Then Exit For
        drvName = Left$(strSave, InStr(1, strSave, Chr$(0)) - 1)
        Select Case GetDriveType(drvName)
        Case 3
  If InStr(drvName, "C") > 0 Then
Option1.Visible = True
ElseIf InStr(drvName, "D") > 0 Then
Option2.Visible = True
ElseIf InStr(drvName, "E") > 0 Then
Option3.Visible = True
ElseIf InStr(drvName, "F") > 0 Then
Option4.Visible = True
ElseIf InStr(drvName, "G") > 0 Then
Option5.Visible = True
End If
           End Select
        strSave = Right$(strSave, Len(strSave) - InStr(1, strSave, Chr$(0)))
    Next keer
Dim FSO As New FileSystemObject
If FSO.FolderExists("c:\YUYU") Then
Option1.Value = True
Label1 = "当前数据存储在C盘，点击下方可更改"
p = "c:\YUYU"
ElseIf FSO.FolderExists("d:\YUYU") Then
Option2.Value = True
Label1 = "当前数据存储在D盘，点击下方可更改"
p = "d:\YUYU"
ElseIf FSO.FolderExists("e:\YUYU") Then
Option3.Value = True
Label1 = "当前数据存储在E盘，点击下方可更改"
p = "e:\YUYU"
ElseIf FSO.FolderExists("f:\YUYU") Then
Option4.Value = True
Label1 = "当前数据存储在F盘，点击下方可更改"
p = "f:\YUYU"
ElseIf FSO.FolderExists("g:\YUYU") Then
Option5.Value = True
Label1 = "当前数据存储在G盘，点击下方可更改"
p = "g:\YUYU"
End If
u = 0
Command1.Enabled = 0
Label1 = Label1 & vbCrLf & p
End Sub

Private Sub Option1_Click()
u = 1
Command1.Enabled = True
End Sub

Private Sub Option2_Click()
u = 1
Command1.Enabled = True
End Sub

Private Sub Option3_Click()
u = 1
Command1.Enabled = True
End Sub

Private Sub Option4_Click()
u = 1
Command1.Enabled = True
End Sub

Private Sub Option5_Click()
u = 1
Command1.Enabled = True
End Sub
