VERSION 5.00
Begin VB.Form Form5 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "其他系统应用"
   ClientHeight    =   8385
   ClientLeft      =   7755
   ClientTop       =   1185
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   10260
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   98
      Text            =   "Text2"
      Top             =   8520
      Width           =   1695
   End
   Begin VB.CommandButton Command59 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command52 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command50 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command76 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command75 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command74 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005809FF&
      Height          =   615
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   90
      Top             =   7560
      Width           =   6135
   End
   Begin VB.CommandButton Command73 
      Caption         =   "73"
      Height          =   255
      Left            =   9360
      TabIndex        =   89
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command71 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command70 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command69 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command68 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command67 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command66 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command63 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command62 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command61 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command60 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command58 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command57 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command56 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command55 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command54 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command53 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command51 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command49 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command48 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command47 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command46 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command45 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command44 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command43 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command42 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command41 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command40 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command39 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   6840
      Width           =   735
   End
   Begin VB.CommandButton Command38 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command37 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command36 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton Command35 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command34 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   6240
      Width           =   735
   End
   Begin VB.CommandButton Command33 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command32 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command31 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command30 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command29 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command28 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command27 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command26 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command25 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command23 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton Command16 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command15 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "停用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "卸载"
      Height          =   375
      Left            =   4200
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "启用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label32 
      BackColor       =   &H00FFFFFF&
      Caption         =   "智能检测"
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
      Left            =   120
      TabIndex        =   91
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "（可能无效）"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   0
      TabIndex        =   88
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "省电精灵"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   87
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "（游戏空间）"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   0
      TabIndex        =   86
      Top             =   6000
      Width           =   1080
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "应用助手"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFFFF&
      Caption         =   "查找手机"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   28
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "图库"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   27
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "相机"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   26
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "日历"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   25
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      Caption         =   "计算器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   24
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "用户改进体验计划"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "K歌特效"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   22
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "视频编辑"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   21
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "杂志锁屏"
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
      Left            =   5280
      TabIndex        =   20
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "畅联通话"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   19
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "天气"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   18
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "智慧视觉"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   17
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "银联服务组件"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "升级向导"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "录音机"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "电子邮件"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
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
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "智能助手(负一屏)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "下拉搜索"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "（停用后重启手机桌面图标即消失）"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      Top             =   6720
      Width           =   2880
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SIM卡应用"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   1
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "系统更新"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.android.hwouc" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command10_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.intelligent" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command11_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.bd" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command12_Click()
Shell "cmd /c adb shell pm disable-user com.android.email" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command13_Click()
Shell "cmd /c adb shell pm disable-user com.android.soundrecorder" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command14_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.android.hwupgradeguide" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command15_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.gameassistant" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command16_Click()
Shell "cmd /c adb shell pm disable-user com.unionpay.tsmservice" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command17_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.scanner" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command18_Click()
Shell "cmd /c adb shell pm enable com.android.stk" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command19_Click()
Shell "cmd /c adb shell pm enable com.huawei.search" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command2_Click()
Shell "cmd /c adb shell pm enable com.huawei.android.hwouc" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command20_Click()
Shell "cmd /c adb shell pm enable com.huawei.intelligent" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command21_Click()
Shell "cmd /c adb shell pm enable com.huawei.powergenie" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command22_Click()
Shell "cmd /c adb shell pm enable com.android.soundrecorder" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command23_Click()
Shell "cmd /c adb shell pm enable com.android.email" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command24_Click()
Shell "cmd /c adb shell pm enable com.huawei.bd" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command25_Click()
Shell "cmd /c adb shell pm enable com.huawei.scanner" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command26_Click()
Shell "cmd /c adb shell pm enable com.unionpay.tsmservice" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command27_Click()
Shell "cmd /c adb shell pm enable com.huawei.gameassistant" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command28_Click()
Shell "cmd /c adb shell pm enable com.huawei.android.hwupgradeguide" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command29_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.android.soundrecorder" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command3_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.android.hwouc" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command30_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.android.email" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command31_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.bd" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command32_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.search" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command33_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.intelligent" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command34_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.android.stk" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command35_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.powergenie" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command36_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.unionpay.tsmservice" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command37_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.gameassistant" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command38_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.android.hwupgradeguide" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command39_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.scanner" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command4_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.powergenie" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command40_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.android.totemweather" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command41_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.android.karaoke" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command42_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.videoeditor" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command43_Click()
Shell "cmd /c adb shell pm disable-user com.android.calendar" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command44_Click()
Shell "cmd /c adb shell pm disable-user com.android.calculator2" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command45_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.android.totemweather" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command46_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.videoeditor" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command47_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.android.karaoke" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command48_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.camera" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command49_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.hwvoipservice" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command5_Click()
Shell "cmd /c adb shell pm disable-user com.android.stk" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command50_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.android.keyguard" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command51_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.android.findmyphone" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub


Private Sub Command52_Click()
Shell "cmd /c adb shell pm enable com.android.keyguard" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command53_Click()
Shell "cmd /c adb shell pm disable-user com.android.gallery3d" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command54_Click()
Shell "cmd /c adb shell pm enable com.android.calculator2" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command55_Click()
Shell "cmd /c adb shell pm enable com.huawei.android.totemweather" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command56_Click()
Shell "cmd /c adb shell pm enable com.huawei.android.karaoke" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command57_Click()
Shell "cmd /c adb shell pm enable com.huawei.videoeditor" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command58_Click()
Shell "cmd /c adb shell pm enable com.huawei.hwvoipservice" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command59_Click()
Shell "cmd /c adb shell pm disable-user com.android.keyguard" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command6_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.search" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command60_Click()
Shell "cmd /c adb shell pm enable com.android.calendar" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command61_Click()
Shell "cmd /c adb shell pm enable com.android.gallery3d" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command62_Click()
Shell "cmd /c adb shell pm enable com.huawei.camera" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command63_Click()
Shell "cmd /c adb shell pm enable com.huawei.android.findmyphone" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command66_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.android.calendar" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command67_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.android.calculator2" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command68_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.android.gallery3d" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command69_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.camera" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub


Private Sub Command70_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.hwvoipservice" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command71_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.android.findmyphone" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub


Private Sub Command73_Click()
Dim Savetime As Single
Savetime = Timer
While Timer < Savetime + 1
DoEvents
Wend
  Dim strFile  As String
    Dim intFile  As Integer
    Dim strData  As String
    strFile = Text2 & "\YUYU临时.txt"
    intFile = FreeFile
    Open strFile For Input As intFile
    strData = StrConv(InputB(FileLen(strFile), intFile), vbUnicode)
    Debug.Print strData
    Close intFile

     If strData = "" Then
MsgBox "该应用无法被停用卸载，或您的手机无该应用", 0 + 48, "执行失败"
 Text1.Text = "失败(T_T)"
 Else
        Text1.Text = "成功!"
        MsgBox "执行成功：" & strData, 0, "成功(≧▽≦)"
        End If
Kill Text2 & "\YUYU临时.txt"
End Sub

Private Sub Command74_Click()
Shell "cmd /c adb shell pm disable-user com.huawei.hwdetectrepair" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command75_Click()
Shell "cmd /c adb shell pm enable com.huawei.hwdetectrepair" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Command76_Click()
Shell "cmd /c adb shell pm uninstall --user 0 com.huawei.hwdetectrepair" & ">" & Text2 & "\YUYU临时.txt"
Call Command73_Click
End Sub

Private Sub Form_Load()
Text2 = Form1.Text5
Text1 = "系统应用请谨慎操作！卸载后只能通过恢复出厂设置恢复应用!" & vbCrLf & "健康使用手机，华为桌面，手机管家无法停用卸载，故不加入"
End Sub

