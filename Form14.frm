VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "µ÷ÕûÊÖ»ú¶¯»­ËÙ¶È"
   ClientHeight    =   3960
   ClientLeft      =   7605
   ClientTop       =   5145
   ClientWidth     =   3315
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3315
   Begin VB.VScrollBar VScroll3 
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   1920
      Width           =   300
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   1320
      Width           =   300
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   720
      Width           =   300
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C9A8FF&
      Caption         =   "È·¶¨"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      MaskColor       =   &H00C9A8FF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "»Ö¸´Ä¬ÈÏ(1x)"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0x(¹Ø±Õ)-10x"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "¶¯»­³ÌÐòÊ±³¤"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "¹ý¶É¶¯»­Ëõ·Å"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "´°¿Ú¶¯»­Ëõ·Å"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1260
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim af, bf, cf As String
Command1.Enabled = 0
If InStr(Text1, "x") > 0 Then
af = Mid(Text1, 1, InStr(Text1, "x") - 1)
ElseIf InStr(Text1, "x") = 0 Then
af = Text1
End If
If InStr(Text2, "x") > 0 Then
bf = Mid(Text2, 1, InStr(Text2, "x") - 1)
ElseIf InStr(Text2, "x") = 0 Then
bf = Text2
End If
If InStr(Text3, "x") > 0 Then
cf = Mid(Text3, 1, InStr(Text3, "x") - 1)
ElseIf InStr(Text3, "x") = 0 Then
cf = Text3
End If

Shell "cmd /c adb shell settings put global window_animation_scale " & af
Shell "cmd /c adb shell settings put global transition_animation_scale " & bf
Shell "cmd /c adb shell settings put global animator_duration_scale " & cf
MsgBox "Ö´ÐÐ³É¹¦"
Command1.Enabled = 1
End Sub

Private Sub Form_Load()
VScroll1.Min = 40
VScroll1.Max = 0
VScroll1.Value = 4
VScroll2.Min = 40
VScroll2.Max = 0
VScroll2.Value = 4
VScroll3.Min = 40
VScroll3.Max = 0
VScroll3.Value = 4
End Sub

Private Sub Label7_Click()
Shell "cmd /c adb shell settings put global window_animation_scale 1"
Shell "cmd /c adb shell settings put global transition_animation_scale 1"
Shell "cmd /c adb shell settings put global animator_duration_scale 1"
MsgBox "ÒÑ»Ö¸´Ä¬ÈÏ"
VScroll3.Value = 4: VScroll1.Value = 4: VScroll2.Value = 4
End Sub


Private Sub VScroll1_Change()
AA = VScroll1.Value
Text1 = AA * 0.25 & "x"
End Sub

Private Sub VScroll2_Change()
BB = VScroll2.Value
Text2 = BB * 0.25 & "x"
End Sub

Private Sub VScroll3_Change()
CC = VScroll3.Value
Text3 = CC * 0.25 & "x"
End Sub
