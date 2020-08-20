VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "华为全家桶"
   ClientHeight    =   8235
   ClientLeft      =   7755
   ClientTop       =   1185
   ClientWidth     =   10695
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10695
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   10920
      TabIndex        =   92
      Text            =   "Text1"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command52 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command70 
      Caption         =   "70"
      Height          =   375
      Left            =   9960
      TabIndex        =   87
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command69 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command68 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command67 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command66 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command65 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command64 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command63 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command62 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton Command61 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command59 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command57 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command55 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Command54 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command53 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command51 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command50 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command49 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command48 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command47 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command46 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command45 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command43 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command42 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command41 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command40 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command39 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command38 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command37 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9480
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   8400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4320
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "停用"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "启用"
      Height          =   375
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "快应用中心"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   91
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "（停用对打电话无影响）"
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   5520
      TabIndex        =   86
      Top             =   5640
      Width           =   1980
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "（HMS Core）"
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   240
      TabIndex        =   85
      Top             =   5640
      Width           =   1080
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为主题"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   75
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为RCS服务"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   74
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "智能提醒"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   73
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为浏览器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   69
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "华为全家桶 部分应用卸载后只能通过恢复出厂设置恢复应用，请谨慎卸载！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   68
      Top             =   7800
      Width           =   10215
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "天际通数据服务"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   30
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "天际通"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   29
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "讯飞语音引擎"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   28
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "语音助手"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   27
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为智能建议"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为智慧引擎"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   25
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为AI引擎"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "精品推荐"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为推送服务"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为移动服务"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为应用市场"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "文件管理(云空间)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "运动健康"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为视频"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为钱包"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "华为音乐"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "百度输入法华为版"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   930
      Left            =   480
      Picture         =   "Form3.frx":74CA
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFile  As String
Dim intFile  As Integer
Dim strData  As String
Dim e As Integer
Private Sub Command1_Click()
Shell "cmd /c adb shell pm disable-user com.baidu.input_huawei" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command10_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.hwid" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command11_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.android.pushagent" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command12_Click()
Shell "cmd /c adb shell pm enable com.huawei.hifolder" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command13_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.hifolder" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command14_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.hifolder" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command15_Click()
e = 1
Shell "cmd /c adb shell pm enable com.android.mediacenter" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
If e = 2 Then
Shell "cmd /c adb shell pm enable com.huawei.music" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End If
End Sub

Private Sub Command16_Click()
Shell "cmd /c adb shell pm enable com.huawei.wallet" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command17_Click()
Shell "cmd /c adb shell pm enable com.huawei.himovie" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command18_Click()
Shell "cmd /c adb shell pm enable com.huawei.health" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command19_Click()
Shell "cmd /c adb shell pm enable com.huawei.hidisk" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command2_Click()
Shell "cmd /c adb shell pm enable com.baidu.input_huawei" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command20_Click()
Shell "cmd /c adb shell pm enable com.huawei.appmarket" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command21_Click()
Shell "cmd /c adb shell pm enable com.huawei.android.pushagent" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command22_Click()
Shell "cmd /c adb shell pm enable com.huawei.hwid" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command23_Click()
Shell "cmd /c adb shell pm enable com.huawei.recsys" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command24_Click()
Shell "cmd /c adb shell pm enable com.huawei.hiai" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command25_Click()
Shell "cmd /c adb shell pm enable com.huawei.pengine" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command26_Click()
Shell "cmd /c adb shell pm enable com.huawei.vassistant" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command27_Click()
Shell "cmd /c adb shell pm enable com.iflytek.speechsuite" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command28_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.fastapp" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command29_Click()
Shell "cmd /c adb shell pm enable com.huawei.hiskytone" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command3_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.baidu.input_huawei" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command30_Click()
Shell "cmd /c adb shell pm enable com.huawei.skytone" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command31_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.recsys" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command32_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.hiai" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command33_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.pengine" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command34_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.vassistant" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command35_Click()
Shell "cmd /c adb shell pm disable-user com.iflytek.speechsuite" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub



Private Sub Command36_Click()
Shell "cmd /c adb shell pm enable com.huawei.fastapp" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command37_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.hiskytone" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command38_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.skytone" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command39_Click()
e = 1
Shell "cmd /c adb shell pm uninstall --user 0 com.android.mediacenter" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
If e = 2 Then
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.music" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End If
End Sub

Private Sub Command4_Click()
e = 1
Shell "cmd /c adb shell pm disable-user com.android.mediacenter" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
If e = 2 Then
Shell "cmd /c adb shell pm disable-user com.huawei.music" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End If
End Sub


Private Sub Command40_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.wallet" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command41_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.himovie" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command42_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.health" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command43_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.hidisk" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command44_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.appmarket" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command45_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.hwid" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command46_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.android.pushagent" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command47_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.recsys" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command48_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.hiai" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command49_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.pengine" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command5_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.wallet" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub


Private Sub Command50_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.vassistant" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command51_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.iflytek.speechsuite" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command52_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.fastapp" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command53_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.hiskytone" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command54_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.skytone" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command55_Click()
Shell "cmd /c adb shell pm disable-user com.android.browser" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command56_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.hicloud" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command57_Click()
Shell "cmd /c adb shell pm enable com.android.browser" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command58_Click()
Shell "cmd /c adb shell pm enable com.huawei.hicloud" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command59_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.android.browser" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command6_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.himovie" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command60_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.hicloud" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command61_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.tips" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command62_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.android.thememanager" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command63_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.rcsserviceapplication" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command64_Click()
Shell "cmd /c adb shell pm enable com.huawei.android.thememanager" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command65_Click()
Shell "cmd /c adb shell pm enable com.huawei.tips" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command66_Click()
Shell "cmd /c adb shell pm enable com.huawei.rcsserviceapplication" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command67_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.android.thememanager" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command68_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.tips" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command69_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.rcsserviceapplication" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command7_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.health" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub


Private Sub Command70_Click()
Dim Savetime As Single
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
  Dim strFile  As String
    Dim intFile  As Integer
    Dim strData  As String
    strFile = Text1 & "\YUYU临时.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile

     If strData = "" Then
If e = 0 Then
MsgBox "该应用无法被停用卸载，或您的手机无该应用", 0 + 48, "执行失败"
 Label19 = "失败(T_T)"
 ElseIf e = 1 Then
 Label19 = "尝试新华为音乐包名"
 e = 2
 ElseIf e = 2 Then
 MsgBox "2个华为音乐包名均无效，或您的手机无该应用", 0 + 48, "执行失败"
 Label19 = "失败(T_T)"
 e = 0
 End If
 Else
       Label19 = "成功!"
        MsgBox "执行成功：" & strData, 0, "成功(≧▽≦)"
End If
Kill Text1 & "\YUYU临时.txt"
End Sub

Private Sub Command8_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.hidisk" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Command9_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.appmarket" & ">" & Text1 & "\YUYU临时.txt"
Call Command70_Click
End Sub

Private Sub Form_Load()
Text1 = Form1.Text5
e = 0
Label19.Alignment = 2
End Sub


