VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   " 百度°音乐"
   ClientHeight    =   7770
   ClientLeft      =   5475
   ClientTop       =   2505
   ClientWidth     =   4470
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "百度音乐"
   MaxButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   4470
   Begin VB.PictureBox Picture23 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   10
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   360
      ScaleWidth      =   4440
      TabIndex        =   75
      Top             =   6120
      Width           =   4440
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   3360
         TabIndex        =   79
         Top             =   45
         Width           =   855
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   2320
         TabIndex        =   78
         Top             =   45
         Width           =   855
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   1260
         TabIndex        =   77
         Top             =   45
         Width           =   855
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   240
         TabIndex        =   76
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3690
      Left            =   10
      Picture         =   "Form1.frx":3E1C
      ScaleHeight     =   3690
      ScaleWidth      =   4425
      TabIndex        =   57
      Top             =   2430
      Width           =   4430
      Begin VB.CommandButton Command1 
         Caption         =   "登录"
         Height          =   360
         Left            =   1680
         TabIndex        =   58
         ToolTipText     =   "登录百度音乐"
         Top             =   3160
         Width           =   1095
      End
      Begin VB.Label Label59 
         Caption         =   "1"
         Height          =   255
         Left            =   3480
         TabIndex        =   88
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label58 
         Caption         =   "1"
         Height          =   255
         Left            =   3480
         TabIndex        =   87
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label57 
         Caption         =   "1"
         Height          =   375
         Left            =   2160
         TabIndex        =   86
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label56 
         Caption         =   "1"
         Height          =   255
         Left            =   1080
         TabIndex        =   85
         Top             =   2760
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label55 
         Caption         =   "1"
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Label19"
         Height          =   255
         Left            =   3360
         TabIndex        =   68
         Top             =   3360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   255
         Left            =   3360
         TabIndex        =   67
         Top             =   3000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label46 
         Caption         =   "Label46"
         Height          =   495
         Left            =   3720
         TabIndex        =   66
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label42 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label41 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   960
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label40 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label39 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label45 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label43 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2650
      Picture         =   "Form1.frx":6CE9
      ScaleHeight     =   405
      ScaleWidth      =   1770
      TabIndex        =   51
      Top             =   2040
      Width           =   1770
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1405
      Left            =   10
      Picture         =   "Form1.frx":7305
      ScaleHeight     =   1410
      ScaleWidth      =   4440
      TabIndex        =   0
      Top             =   0
      Width           =   4440
      Begin VB.PictureBox Picture20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2760
         Picture         =   "Form1.frx":E257
         ScaleHeight     =   270
         ScaleWidth      =   1575
         TabIndex        =   71
         Top             =   0
         Width           =   1575
         Begin VB.PictureBox Picture22 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   0
            Picture         =   "Form1.frx":E84F
            ScaleHeight     =   270
            ScaleWidth      =   375
            TabIndex        =   80
            Top             =   0
            Width           =   375
         End
         Begin VB.PictureBox Picture28 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1080
            Picture         =   "Form1.frx":EBF9
            ScaleHeight     =   270
            ScaleWidth      =   495
            TabIndex        =   74
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.PictureBox Picture26 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   730
            Picture         =   "Form1.frx":F056
            ScaleHeight     =   270
            ScaleWidth      =   360
            TabIndex        =   73
            Top             =   0
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.PictureBox Picture24 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   360
            Picture         =   "Form1.frx":F3B8
            ScaleHeight     =   270
            ScaleWidth      =   375
            TabIndex        =   72
            Top             =   0
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3600
         Picture         =   "Form1.frx":F718
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   52
         Top             =   840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Picture21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4000
         Picture         =   "Form1.frx":11EBD
         ScaleHeight     =   330
         ScaleWidth      =   375
         TabIndex        =   5
         ToolTipText     =   "搜索"
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000011&
         Height          =   230
         Left            =   1370
         MousePointer    =   3  'I-Beam
         TabIndex        =   4
         Text            =   "搜索 歌曲、歌手、专辑"
         Top             =   430
         Width           =   2640
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3600
         Picture         =   "Form1.frx":12271
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   840
         Width           =   480
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3600
         Picture         =   "Form1.frx":149A5
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         ToolTipText     =   "打开文件"
         Top             =   840
         Width           =   480
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000002&
         X1              =   0
         X2              =   4440
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "音乐你的生活"
         Height          =   180
         Left            =   960
         TabIndex        =   70
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "百度音乐_"
         Height          =   180
         Left            =   120
         TabIndex        =   69
         Top             =   120
         Width           =   810
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00F97364&
         X1              =   0
         X2              =   4440
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 登录"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   180
         Left            =   2040
         MouseIcon       =   "Form1.frx":17256
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   120
         Width           =   450
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hello_"
         ForeColor       =   &H00004040&
         Height          =   180
         Left            =   960
         MouseIcon       =   "Form1.frx":173A8
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   120
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   865
         Left            =   210
         MouseIcon       =   "Form1.frx":174FA
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         ToolTipText     =   "更换头像"
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "百度音乐_音乐你的生活"
         Height          =   390
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   2325
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "         "
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   4455
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2160
         TabIndex        =   7
         Top             =   1160
         Width           =   105
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         Height          =   180
         Left            =   2280
         TabIndex        =   6
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0FF&
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Width           =   825
      End
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3950
      Picture         =   "Form1.frx":17804
      ScaleHeight     =   315
      ScaleWidth      =   510
      TabIndex        =   56
      ToolTipText     =   "打开乐库"
      Top             =   1710
      Width           =   510
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3950
      Picture         =   "Form1.frx":19840
      ScaleHeight     =   300
      ScaleWidth      =   510
      TabIndex        =   55
      ToolTipText     =   "打开歌词"
      Top             =   1410
      Width           =   510
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3950
      Picture         =   "Form1.frx":1B7AE
      ScaleHeight     =   315
      ScaleWidth      =   510
      TabIndex        =   54
      Top             =   1710
      Width           =   510
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3950
      Picture         =   "Form1.frx":1BB67
      ScaleHeight     =   300
      ScaleWidth      =   510
      TabIndex        =   53
      Top             =   1410
      Width           =   510
   End
   Begin VB.PictureBox Picture16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1800
      Picture         =   "Form1.frx":1BEEF
      ScaleHeight     =   405
      ScaleWidth      =   1335
      TabIndex        =   50
      ToolTipText     =   "切换到‘我的收藏’"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.PictureBox Picture15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1320
      Picture         =   "Form1.frx":1E631
      ScaleHeight     =   405
      ScaleWidth      =   1770
      TabIndex        =   49
      Top             =   2040
      Width           =   1770
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   30
      Picture         =   "Form1.frx":1ECFE
      ScaleHeight     =   390
      ScaleWidth      =   1320
      TabIndex        =   48
      ToolTipText     =   "切换到‘播放列表’"
      Top             =   2050
      Width           =   1320
   End
   Begin VB.PictureBox Picture17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   3120
      Picture         =   "Form1.frx":21147
      ScaleHeight     =   405
      ScaleWidth      =   1320
      TabIndex        =   47
      ToolTipText     =   "切换到‘随便听听’"
      Top             =   2040
      Width           =   1320
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   30
      Picture         =   "Form1.frx":236E0
      ScaleHeight     =   390
      ScaleWidth      =   1320
      TabIndex        =   46
      Top             =   2050
      Width           =   1320
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   3120
      Picture         =   "Form1.frx":23C13
      ScaleHeight     =   405
      ScaleWidth      =   1320
      TabIndex        =   45
      Top             =   2040
      Width           =   1320
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   1800
      Picture         =   "Form1.frx":24194
      ScaleHeight     =   405
      ScaleWidth      =   1335
      TabIndex        =   44
      Top             =   2040
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   30
      Picture         =   "Form1.frx":24796
      ScaleHeight     =   405
      ScaleWidth      =   1770
      TabIndex        =   43
      Top             =   2040
      Width           =   1770
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   10
      Picture         =   "Form1.frx":24DBD
      ScaleHeight     =   1305
      ScaleWidth      =   4455
      TabIndex        =   13
      Top             =   6480
      Width           =   4455
      Begin VB.Timer Timer7 
         Interval        =   10
         Left            =   3360
         Top             =   1080
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H8000000C&
         Height          =   300
         Left            =   600
         MousePointer    =   3  'I-Beam
         TabIndex        =   82
         Top             =   40
         Width           =   3375
      End
      Begin VB.Timer Timer6 
         Interval        =   100
         Left            =   2760
         Top             =   480
      End
      Begin VB.Timer Timer5 
         Interval        =   10
         Left            =   2400
         Top             =   480
      End
      Begin VB.Timer Timer4 
         Interval        =   300
         Left            =   2040
         Top             =   480
      End
      Begin VB.Timer Timer3 
         Interval        =   4000
         Left            =   1680
         Top             =   480
      End
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   1320
         Top             =   480
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   960
         Top             =   480
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   3840
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3360
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label60 
         Caption         =   "label60"
         Height          =   255
         Left            =   1800
         TabIndex        =   89
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   4005
         TabIndex        =   83
         Top             =   65
         Width           =   255
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   255
         TabIndex        =   81
         Top             =   75
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   0
         Picture         =   "Form1.frx":2DBD5
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4425
      End
      Begin VB.Label Label37 
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label36 
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label35 
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label34 
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label33 
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label22 
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "百度音乐_音乐你的生活"
         Height          =   180
         Left            =   210
         TabIndex        =   14
         Top             =   1020
         Width           =   1890
      End
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   1230
      Picture         =   "Form1.frx":2E626
      Stretch         =   -1  'True
      Top             =   1640
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   920
      Picture         =   "Form1.frx":2E9BD
      Stretch         =   -1  'True
      Top             =   1640
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   4450
      X2              =   4450
      Y1              =   0
      Y2              =   7320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      BorderWidth     =   2
      X1              =   15
      X2              =   0
      Y1              =   0
      Y2              =   7800
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   735
      Left            =   10
      TabIndex        =   42
      Top             =   1305
      Width           =   3960
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6985
      _cy             =   1296
   End
   Begin VB.Label Label32 
      BackColor       =   &H00FFC0FF&
      Height          =   345
      Left            =   10
      TabIndex        =   41
      Top             =   5760
      Width           =   11805
   End
   Begin VB.Label Label31 
      Height          =   345
      Left            =   10
      TabIndex        =   40
      Top             =   5400
      Width           =   11805
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFC0FF&
      Height          =   345
      Left            =   10
      TabIndex        =   39
      Top             =   5040
      Width           =   11805
   End
   Begin VB.Label Label29 
      Height          =   345
      Left            =   10
      TabIndex        =   38
      Top             =   4680
      Width           =   11805
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFC0FF&
      Height          =   345
      Left            =   10
      TabIndex        =   37
      Top             =   4320
      Width           =   11805
   End
   Begin VB.Label Label27 
      Height          =   345
      Left            =   10
      TabIndex        =   36
      Top             =   3960
      Width           =   11805
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFC0FF&
      Height          =   345
      Left            =   10
      TabIndex        =   35
      Top             =   3600
      Width           =   11805
   End
   Begin VB.Label Label25 
      Height          =   345
      Left            =   10
      TabIndex        =   34
      Top             =   3240
      Width           =   11805
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFC0FF&
      Height          =   345
      Left            =   10
      TabIndex        =   33
      Top             =   2880
      Width           =   11805
   End
   Begin VB.Label Label23 
      Height          =   345
      Left            =   10
      TabIndex        =   32
      Top             =   2520
      Width           =   11805
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0FF&
      Height          =   345
      Left            =   360
      TabIndex        =   31
      Top             =   5760
      Width           =   11805
   End
   Begin VB.Label Label9 
      Height          =   345
      Left            =   360
      TabIndex        =   30
      Top             =   5400
      Width           =   11805
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0FF&
      Height          =   345
      Left            =   360
      TabIndex        =   29
      Top             =   5040
      Width           =   11805
   End
   Begin VB.Label Label7 
      Height          =   345
      Left            =   360
      TabIndex        =   28
      Top             =   4680
      Width           =   11805
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0FF&
      Height          =   345
      Left            =   360
      TabIndex        =   27
      Top             =   4320
      Width           =   11805
   End
   Begin VB.Label Label5 
      Height          =   345
      Left            =   360
      TabIndex        =   26
      Top             =   3960
      Width           =   11805
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0FF&
      Height          =   345
      Left            =   360
      TabIndex        =   25
      Top             =   3600
      Width           =   11805
   End
   Begin VB.Label Label3 
      Height          =   345
      Left            =   360
      TabIndex        =   24
      Top             =   3240
      Width           =   11805
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Height          =   345
      Left            =   360
      TabIndex        =   23
      Top             =   2880
      Width           =   11805
   End
   Begin VB.Label Label1 
      Height          =   345
      Left            =   360
      TabIndex        =   22
      Top             =   2520
      Width           =   11805
   End
   Begin VB.Label Label11 
      Caption         =   "       1.        2.        3.        4.        5.        6.        7.        8.        9.       10."
      ForeColor       =   &H80000011&
      Height          =   3735
      Left            =   15
      TabIndex        =   21
      Top             =   2400
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, d, e, f, g, h, i, j As String
Dim k, l, ma, n, o, p, q, r, s, t As String
Dim d1, d2, d3, d4, d5, d6, d7, d8, d9, d10 As String
Dim e1, e2, e3, e4 As String
Dim m As String
Dim simble As String
Dim pic As String
Dim urlstring As String
Dim mouse_x, mouse_y As String




Private Sub addbendi_Click()


End Sub

Private Sub addwangluo_Click()
Form4.Show 1
End Sub

Private Sub clean_Click()
Label1.Caption = ""

Label2.Caption = ""

Label3.Caption = ""

Label4.Caption = ""

Label5.Caption = ""

Label6.Caption = ""

Label7.Caption = ""

Label8.Caption = ""

Label9.Caption = ""

Label10.Caption = ""

a = ""
b = ""
c = ""
d = ""
e = ""
f = ""
g = ""
h = ""
i = ""
j = ""
End Sub

Private Sub Command1_Click()
Form7.Show 1
End Sub

Private Sub danqu_Click()
   Form3.Option4.Value = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = True
End Sub

Private Sub delete_Click()
On Error Resume Next
If Label1.BackColor = &HFF8080 Then
   Label1.Caption = ""
ElseIf Label2.BackColor = &HFF8080 Then
   Label2.Caption = ""
ElseIf Label3.BackColor = &HFF8080 Then
   Label3.Caption = ""
ElseIf Label4.BackColor = &HFF8080 Then
   Label4.Caption = ""
ElseIf Label5.BackColor = &HFF8080 Then
   Label5.Caption = ""
ElseIf Label6.BackColor = &HFF8080 Then
   Label6.Caption = ""
ElseIf Label7.BackColor = &HFF8080 Then
   Label7.Caption = ""
ElseIf Label8.BackColor = &HFF8080 Then
   Label8.Caption = ""
ElseIf Label9.BackColor = &HFF8080 Then
   Label9.Caption = ""
ElseIf Label10.BackColor = &HFF8080 Then
   Label10.Caption = ""
End If
'我的收藏
If Label23.BackColor = vbRed Then
   Label23.Caption = ""
ElseIf Label24.BackColor = vbRed Then
   Label24.Caption = ""
ElseIf Label25.BackColor = vbRed Then
   Label25.Caption = ""
ElseIf Label26.BackColor = vbRed Then
   Label26.Caption = ""
ElseIf Label27.BackColor = vbRed Then
   Label27.Caption = ""
ElseIf Label28.BackColor = vbRed Then
   Label28.Caption = ""
ElseIf Label29.BackColor = vbRed Then
   Label29.Caption = ""
ElseIf Label30.BackColor = vbRed Then
   Label30.Caption = ""
ElseIf Label31.BackColor = vbRed Then
   Label31.Caption = ""
ElseIf Label32.BackColor = vbRed Then
   Label32.Caption = ""
End If

If Label1.Caption = "" Then
   m = a
   Label1.Caption = Label2.Caption
   a = b
   Label2.Caption = Label3.Caption
   b = c
   Label3.Caption = Label4.Caption
   c = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 1 Then
   WindowsMediaPlayer1.URL = a
   Label16.Caption = "正在播放：" & Label1.Caption
   End If
ElseIf Label2.Caption = "" Then
   m = b
   Label2.Caption = Label3.Caption
   b = c
   Label3.Caption = Label4.Caption
   c = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 2 Then
   WindowsMediaPlayer1.URL = b
   Label16.Caption = "正在播放：" & Label2.Caption
   End If
ElseIf Label3.Caption = "" Then
   m = c
   Label3.Caption = Label4.Caption
   c = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 3 Then
   WindowsMediaPlayer1.URL = c
   Label16.Caption = "正在播放：" & Label3.Caption
   End If
ElseIf Label4.Caption = "" Then
   m = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 4 Then
   WindowsMediaPlayer1.URL = d
   Label16.Caption = "正在播放：" & Label4.Caption
   End If
ElseIf Label5.Caption = "" Then
   m = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 5 Then
   WindowsMediaPlayer1.URL = e
   Label16.Caption = "正在播放：" & Label5.Caption
   End If
ElseIf Label6.Caption = "" Then
   m = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 6 Then
   WindowsMediaPlayer1.URL = f
   Label16.Caption = "正在播放：" & Label6.Caption
   End If
ElseIf Label7.Caption = "" Then
   m = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 7 Then
   WindowsMediaPlayer1.URL = g
   Label16.Caption = "正在播放：" & Label7.Caption
   End If
ElseIf Label8.Caption = "" Then
   m = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 8 Then
   WindowsMediaPlayer1.URL = h
   Label16.Caption = "正在播放：" & Label8.Caption
   End If
ElseIf Label9.Caption = "" Then
   m = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 9 Then
   WindowsMediaPlayer1.URL = i
   Label16.Caption = "正在播放：" & Label9.Caption
   End If
End If
   
'我的收藏
If Label23.Caption = "" Then
   m = k
   Label23.Caption = Label24.Caption
   k = l
   Label24.Caption = Label25.Caption
   l = ma
   Label25.Caption = Label26.Caption
   ma = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 11 Then
   WindowsMediaPlayer1.URL = k
   Label16.Caption = "正在播放：" & Label23.Caption
   End If
ElseIf Label24.Caption = "" Then
   m = l
   Label24.Caption = Label25.Caption
   l = ma
   Label25.Caption = Label26.Caption
   ma = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 12 Then
   WindowsMediaPlayer1.URL = l
   Label16.Caption = "正在播放：" & Label24.Caption
   End If
ElseIf Label25.Caption = "" Then
   m = ma
   Label25.Caption = Label26.Caption
   ma = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 13 Then
   WindowsMediaPlayer1.URL = ma
   Label16.Caption = "正在播放：" & Label25.Caption
   End If
ElseIf Label26.Caption = "" Then
   m = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 14 Then
   WindowsMediaPlayer1.URL = n
   Label16.Caption = "正在播放：" & Label26.Caption
   End If
ElseIf Label27.Caption = "" Then
   m = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 15 Then
   WindowsMediaPlayer1.URL = o
   Label16.Caption = "正在播放：" & Label27.Caption
   End If
ElseIf Label28.Caption = "" Then
   m = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 16 Then
   WindowsMediaPlayer1.URL = p
   Label16.Caption = "正在播放：" & Label28.Caption
   End If
ElseIf Label29.Caption = "" Then
   m = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 17 Then
   WindowsMediaPlayer1.URL = q
   Label16.Caption = "正在播放：" & Label29.Caption
   End If
ElseIf Label30.Caption = "" Then
   m = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 18 Then
   WindowsMediaPlayer1.URL = r
   Label16.Caption = "正在播放：" & Label30.Caption
   End If
ElseIf Label31.Caption = "" Then
   m = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = 19 Then
   WindowsMediaPlayer1.URL = s
   Label16.Caption = "正在播放：" & Label31.Caption
   End If
End If
   

   
   
   
End Sub

Private Sub Form_Load()
On Error Resume Next

If Dir("c:\百度音乐播放器", vbDirectory) = "" Then
   MkDir ("c:\百度音乐播放器")
End If

CommonDialog1.Filter = "音乐文件 （*.mp3;*.wma）|*.mp3;*.wma;"
CommonDialog2.Filter = "图片文件 (*.jpg;*.gif;*.bmp;*.jpeg;*.jpe;)|*.jpg;*.gif;*.bmp;*.jpeg;*.jpe;)"


Form1.BorderStyle = 0 - none
Picture6.Visible = True
Picture9.Visible = False
Picture7.Visible = True
Picture10.Visible = False
Picture8.Visible = True
Picture11.Visible = False
Picture12.Visible = False
Picture19.Visible = False

Label12.Caption = "00:00"
Timer1.Enabled = False
Timer2.Enabled = True
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False

Picture13.Visible = False
Picture14.Visible = False
Picture2.Visible = True
Picture15.Visible = False
Picture18.Visible = False


Image2.Visible = False
Combo1.Visible = False
Label53.Visible = False
Label54.Visible = False

Label16.Caption = "百度音乐_音乐你的生活"
Label16.Alignment = 2

Text1.Visible = True


Picture21.Visible = False
Label20.Visible = False


Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
Label26.Visible = False
Label27.Visible = False
Label28.Visible = False
Label29.Visible = False
Label30.Visible = False
Label31.Visible = False
Label32.Visible = False


If Dir("c:\百度音乐播放器\播放列表.m3u") <> "" And Dir("c:\百度音乐播放器\播放列表.txt") = "" Then
   Name "c:\百度音乐播放器\播放列表.m3u" As "c:\百度音乐播放器\播放列表.txt"
ElseIf Dir("c:\百度音乐播放器\播放列表.m3u") = "" And Dir("c:\百度音乐播放器\播放列表.txt") = "" Then
   Open "c:\百度音乐播放器\播放列表.txt" For Output As #10
   Close #10
End If

If Dir("c:\百度音乐播放器\system setup.ini") <> "" And Dir("c:\百度音乐播放器\system setup.txt") = "" Then
   Name "c:\百度音乐播放器\system setup.ini" As "c:\百度音乐播放器\system setup.txt"
ElseIf Dir("c:\百度音乐播放器\system setup.ini") = "" And Dir("c:\百度音乐播放器\system setup.txt") = "" Then
   Open "c:\百度音乐播放器\system setup.txt" For Output As #18
   Close #18
End If


If Dir("c:\百度音乐播放器\播放列表.txt") <> "" Then
   Open "c:\百度音乐播放器\播放列表.txt" For Input As #8
     If Not EOF(8) Then
       Line Input #8, d1
       Label1.Caption = d1
       Line Input #8, a
       Line Input #8, d2
       Label2.Caption = d2
       Line Input #8, b
       Line Input #8, d3
       Label3.Caption = d3
       Line Input #8, c
       Line Input #8, d4
       Label4.Caption = d4
       Line Input #8, d
       Line Input #8, d5
       Label5.Caption = d5
       Line Input #8, e
       Line Input #8, d6
       Label6.Caption = d6
       Line Input #8, f
       Line Input #8, d7
       Label7.Caption = d7
       Line Input #8, g
       Line Input #8, d8
       Label8.Caption = d8
       Line Input #8, h
       Line Input #8, d9
       Label9.Caption = d9
       Line Input #8, i
       Line Input #8, d10
       Label10.Caption = d10
       Line Input #8, j
     End If
   Close #8
  
End If
       
       
If Dir("c:\百度音乐播放器\system setup.txt") <> "" Then
   Open "c:\百度音乐播放器\system setup.txt" For Input As #11
      If Not EOF(11) Then
        Line Input #11, e1
        Line Input #11, e2
        Line Input #11, e3
        Line Input #11, e4
      End If
      If e1 <> "" And e2 <> "" And e3 <> "" And e4 <> "" Then
        Form3.Option1.Value = e1
        Form3.Option2.Value = e2
        Form3.Option3.Value = e3
        Form3.Option4.Value = e4

      End If
   Close #11
ElseIf Dir("c:\百度音乐播放器\system setup.txt") = "" Then
     Form3.Option1.Value = True
     Form3.Option2.Value = False
     Form3.Option3.Value = False
     Form3.Option4.Value = False
     
     Form10.shunxu.Checked = True
     Form10.shunxuxunhuan.Checked = False
     Form10.shuiji.Checked = False
     Form10.danqu.Checked = False
     
    Form10.shunxu2.Checked = True
    Form10.shunxuxunhuan2.Checked = False
    Form10.shuiji2.Checked = False
    Form10.danqu2.Checked = False
End If

If Form3.Option1.Value = False And Form3.Option2.Value = False And Form3.Option3.Value = False And Form3.Option4.Value = False Then
   Form3.Option1.Value = True
     Form3.Option2.Value = False
     Form3.Option3.Value = False
     Form3.Option4.Value = False
     
     Form10.shunxu.Checked = True
     Form10.shunxuxunhuan.Checked = False
     Form10.shuiji.Checked = False
     Form10.danqu.Checked = False
     
    Form10.shunxu2.Checked = True
    Form10.shunxuxunhuan2.Checked = False
    Form10.shuiji2.Checked = False
    Form10.danqu2.Checked = False
 End If



If Form3.Option1.Value = True Then

   
   Form10.shunxu.Checked = True
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = True
    Form10.shunxuxunhuan2.Checked = False
    Form10.shuiji2.Checked = False
    Form10.danqu2.Checked = False
ElseIf Form3.Option2.Value = True Then
  
   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = True
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = False
    Form10.shunxuxunhuan2.Checked = True
    Form10.shuiji2.Checked = False
    Form10.danqu2.Checked = False
ElseIf Form3.Option3.Value = True Then
   
   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = True
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = False
    Form10.shunxuxunhuan2.Checked = False
    Form10.shuiji2.Checked = True
    Form10.danqu2.Checked = False
ElseIf Form3.Option4.Value = True Then

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = True
   
   Form10.shunxu2.Checked = False
    Form10.shunxuxunhuan2.Checked = False
    Form10.shuiji2.Checked = False
    Form10.danqu2.Checked = True
End If
   
  If Label22.Caption = "" Then
      Form10.load.Enabled = True
      Form10.unload.Enabled = False
   ElseIf Label22.Caption <> "" Then
      Form10.load.Enabled = False
      Form10.unload.Enabled = True
   End If
   
   If Picture2.Visible = True Then
        If Label2.Caption = "" Then
           Image3.Visible = False
           Image4.Visible = False
        ElseIf Label2.Caption <> "" Then
           Image3.Visible = True
           Image4.Visible = True
        End If
ElseIf Picture15.Visible = True Then
        If Label24.Caption = "" Then
           Image3.Visible = False
           Image4.Visible = False
        ElseIf Label24.Caption <> "" Then
           Image3.Visible = True
           Image4.Visible = True
        End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If 3945 < X And X < 4425 And 1410 < Y And Y < 1710 Then
   Picture7.Visible = False
   Picture10.Visible = True
Else
   Picture7.Visible = True
   Picture10.Visible = False
End If

If 3945 < X And X < 4425 And 1710 < Y And Y < 1980 Then
   Picture8.Visible = False
   Picture11.Visible = True
Else
   Picture8.Visible = True
   Picture11.Visible = False
End If

If 0 <= X And X <= 1325 And 2055 <= Y And Y <= 2400 And Picture2.Visible = False Then
  Picture13.Visible = False
  Picture14.Visible = True
ElseIf 1345 <= X And X <= 3120 And 2055 <= Y And Y <= 2400 And Picture15.Visible = False Then
  Picture3.Visible = False
  Picture16.Visible = True
ElseIf 3125 <= X And X <= 4425 And 2055 <= Y And Y <= 2400 And Picture18.Visible = False Then
  Picture4.Visible = False
  Picture17.Visible = True
Else
  If Picture2.Visible = False Then
     Picture13.Visible = True
     Picture14.Visible = False
  End If
  If Picture15.Visible = False Then
     Picture3.Visible = True
     Picture16.Visible = False
  End If
  If Picture18.Visible = False Then
     Picture4.Visible = True
     Picture17.Visible = False
  End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)


If Dir("c:\百度音乐播放器\", vbDirectory) <> "" Then
   Open "c:\百度音乐播放器\播放列表.txt" For Output As #9
   Print #9, Label1.Caption
   Print #9, a
   Print #9, Label2.Caption
   Print #9, b
   Print #9, Label3.Caption
   Print #9, c
   Print #9, Label4.Caption
   Print #9, d
   Print #9, Label5.Caption
   Print #9, e
   Print #9, Label6.Caption
   Print #9, f
   Print #9, Label7.Caption
   Print #9, g
   Print #9, Label8.Caption
   Print #9, h
   Print #9, Label9.Caption
   Print #9, i
   Print #9, Label10.Caption
   Print #9, j
   Close #9
End If



If Label22.Caption <> "" Then
     Open "c:\百度音乐播放器\" & Label22.Caption & "的收藏.txt" For Output As #2
   Print #2, Label23.Caption
   Print #2, k
   Print #2, Label24.Caption
   Print #2, l
   Print #2, Label25.Caption
   Print #2, ma
   Print #2, Label26.Caption
   Print #2, n
   Print #2, Label27.Caption
   Print #2, o
   Print #2, Label28.Caption
   Print #2, p
   Print #2, Label29.Caption
   Print #2, q
   Print #2, Label30.Caption
   Print #2, r
   Print #2, Label31.Caption
   Print #2, s
   Print #2, Label32.Caption
   Print #2, t
   Print #2, pic
   Close #2
     
   Name "c:\百度音乐播放器\" & Label22.Caption & "的收藏.txt" As "c:\百度音乐播放器\" & Label22.Caption & "的收藏.m3u"
End If
  



 Open "c:\百度音乐播放器\system setup.txt" For Output As #1
 Print #1, Form3.Option1.Value
 Print #1, Form3.Option2.Value
 Print #1, Form3.Option3.Value
 Print #1, Form3.Option4.Value
 Close #1


If Dir("c:\百度音乐播放器\播放列表.m3u") = "" And Dir("c:\百度音乐播放器\播放列表.txt") <> "" Then
   Name "c:\百度音乐播放器\播放列表.txt" As "c:\百度音乐播放器\播放列表.m3u"
ElseIf Dir("c:\百度音乐播放器\播放列表.m3u") = "" And Dir("c:\百度音乐播放器\播放列表.txt") = "" Then
   Open "c:\百度音乐播放器\播放列表.txt" For Output As #10
   Close #10
   Name "c:\百度音乐播放器\播放列表.txt" As "c:\百度音乐播放器\播放列表.m3u"
End If

If Dir("c:\百度音乐播放器\system setup.txt") <> "" And Dir("c:\百度音乐播放器\system setup.ini") = "" Then
   Name "c:\百度音乐播放器\system setup.txt" As "c:\百度音乐播放器\system setup.ini"
ElseIf Dir("c:\百度音乐播放器\system setup.txt") = "" And Dir("c:\百度音乐播放器\system setup.ini") = "" Then
   Open "c:\百度音乐播放器\system setup.txt" For Output As #18
   Close #18
   Name "c:\百度音乐播放器\system setup.txt" As "c:\百度音乐播放器\system setup.ini"
End If

unload Form2
unload Form3
unload Form4
unload Form5

unload Form7
unload Form8
unload Form9

unload Form1

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  

   CommonDialog2.CancelError = True
   On Error GoTo errline
   CommonDialog2.ShowOpen
   'CommonDialog2.Filter = "图片文件 (*.jpg;*.jpeg;*.jpe;*.bmp;*.gif)|*.jpg;*.jpeg;*.jpe;*.bmp;*.gif)"
   Image1.Picture = LoadPicture(CommonDialog2.FileName)
   Form3.Image1.Picture = LoadPicture(CommonDialog2.FileName)
   
     pic = CommonDialog2.FileName
  
   MousePointer = 1
   
     Text1.Enabled = False
   If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
   End If
   
Exit Sub

errline:
   MousePointer = 1
     Text1.Enabled = False
  If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
  End If
   Exit Sub
End If

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label53.BorderStyle = 0
Label54.BorderStyle = 0
End Sub

Private Sub Image3_Click()

                        If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                                If Label10 <> "" Then
                                                                    WindowsMediaPlayer1.URL = j
                                                                    Label16.Caption = "正在播放：" & Label10.Caption
                                                                    simble = 10
                                                                ElseIf Label10 = "" And Label9 <> "" Then
                                                                    WindowsMediaPlayer1.URL = i
                                                                    Label16.Caption = "正在播放：" & Label9.Caption
                                                                    simble = 9
                                                                ElseIf Label9 = "" And Label8 <> "" Then
                                                                    WindowsMediaPlayer1.URL = h
                                                                    Label16.Caption = "正在播放：" & Label8.Caption
                                                                    simble = 8
                                                                ElseIf Label8 = "" And Label7 <> "" Then
                                                                    WindowsMediaPlayer1.URL = g
                                                                    Label16.Caption = "正在播放：" & Label7.Caption
                                                                    simble = 7
                                                                ElseIf Label7 = "" And Label6 <> "" Then
                                                                    WindowsMediaPlayer1.URL = f
                                                                    Label16.Caption = "正在播放：" & Label6.Caption
                                                                    simble = 6
                                                                ElseIf Label6 = "" And Label5 <> "" Then
                                                                    WindowsMediaPlayer1.URL = e
                                                                    Label16.Caption = "正在播放：" & Label5.Caption
                                                                    simble = 5
                                                                ElseIf Label5 = "" And Label4 <> "" Then
                                                                    WindowsMediaPlayer1.URL = d
                                                                    Label16.Caption = "正在播放：" & Label4.Caption
                                                                    simble = 4
                                                                ElseIf Label4 = "" And Label3 <> "" Then
                                                                    WindowsMediaPlayer1.URL = c
                                                                    Label16.Caption = "正在播放：" & Label3.Caption
                                                                    simble = 3
                                                                ElseIf Label3 = "" And Label2 <> "" Then
                                                                    WindowsMediaPlayer1.URL = b
                                                                    Label16.Caption = "正在播放：" & Label2.Caption
                                                                    simble = 2
                                                                ElseIf Label2 = "" And Label1 <> "" Then
                                                                    WindowsMediaPlayer1.URL = a
                                                                    Label16.Caption = "正在播放：" & Label1.Caption
                                                                    simble = 1
                                                                End If
                      ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                                       WindowsMediaPlayer1.URL = a
                                                                       Label16.Caption = "正在播放：" & Label1.Caption
                                                                       simble = 1
                       ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                        WindowsMediaPlayer1.URL = b
                                                        Label16.Caption = "正在播放：" & Label2.Caption
                                                        simble = 2
                       ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                        WindowsMediaPlayer1.URL = c
                                                        Label16.Caption = "正在播放：" & Label3.Caption
                                                        simble = 3
                    ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                                        WindowsMediaPlayer1.URL = d
                                                        Label16.Caption = "正在播放：" & Label4.Caption
                                                        simble = 4
                 ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
                                                        WindowsMediaPlayer1.URL = e
                                                        Label16.Caption = "正在播放：" & Label5.Caption
                                                        simble = 5
                         ElseIf WindowsMediaPlayer1.URL = g And simble = 7 Then
                                                            WindowsMediaPlayer1.URL = f
                                                            Label16.Caption = "正在播放：" & Label6.Caption
                                                            simble = 6
                         ElseIf WindowsMediaPlayer1.URL = h And simble = 8 Then
                                                                WindowsMediaPlayer1.URL = g
                                                                Label16.Caption = "正在播放：" & Label7.Caption
                                                                simble = 7
                           ElseIf WindowsMediaPlayer1.URL = i And simble = 9 Then
                                                                WindowsMediaPlayer1.URL = h
                                                                Label16.Caption = "正在播放：" & Label8.Caption
                                                                simble = 8
                         ElseIf WindowsMediaPlayer1.URL = j And simble = 10 Then
                                                                    WindowsMediaPlayer1.URL = i
                                                                    Label16.Caption = "正在播放：" & Label9.Caption
                                                                    simble = 9
  
                        

'我的收藏——————————————————————————————————————————顺序循环播放—————————我的收藏———————

                         ElseIf WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                                    WindowsMediaPlayer1.URL = t
                                                                    Label16.Caption = "正在播放：" & Label32.Caption
                                                                    simble = 20
                        ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                                    WindowsMediaPlayer1.URL = k
                                                                    Label16.Caption = "正在播放：" & Label23.Caption
                                                                    simble = 11
                                ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                                        WindowsMediaPlayer1.URL = l
                                                                        Label16.Caption = "正在播放：" & Label24.Caption
                                                                        simble = 12
                            ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                                            WindowsMediaPlayer1.URL = ma
                                                                        Label16.Caption = "正在播放：" & Label25.Caption
                                                                        simble = 13
                             ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                                                        WindowsMediaPlayer1.URL = n
                                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                                        simble = 14
                      ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
                                                                        WindowsMediaPlayer1.URL = o
                                                                        Label16.Caption = "正在播放：" & Label27.Caption
                                                                        simble = 15
                       ElseIf WindowsMediaPlayer1.URL = q And simble = 17 Then
                                                                        WindowsMediaPlayer1.URL = p
                                                                        Label16.Caption = "正在播放：" & Label28.Caption
                                                                        simble = 16
                           ElseIf WindowsMediaPlayer1.URL = r And simble = 18 Then
                                                                            WindowsMediaPlayer1.URL = q
                                                                            Label16.Caption = "正在播放：" & Label29.Caption
                                                                            simble = 17
                        ElseIf WindowsMediaPlayer1.URL = s And simble = 19 Then
                                                                        WindowsMediaPlayer1.URL = r
                                                                        Label16.Caption = "正在播放：" & Label30.Caption
                                                                        simble = 18
                        ElseIf WindowsMediaPlayer1.URL = t And simble = 20 Then
                                                                        WindowsMediaPlayer1.URL = s
                                                                        Label16.Caption = "正在播放：" & Label31.Caption
                                                                        simble = 19
                        End If
                        
       Timer5.Enabled = True
End Sub

Private Sub Image4_Click()

                        If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                    If Label2.Caption <> "" Then
                                                                    WindowsMediaPlayer1.URL = b
                                                                    Label16.Caption = "正在播放：" & Label2.Caption
                                                                    simble = 2
        
                                                     Else
                                                                        WindowsMediaPlayer1.URL = a
                                                                    Label16.Caption = "正在播放：" & Label1.Caption
                                                                    simble = 1

                                                    End If
                      ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                 If Label3.Caption <> "" Then
                                                                       WindowsMediaPlayer1.URL = c
                                                                       Label16.Caption = "正在播放：" & Label3.Caption
                                                                       simble = 3
                                                                    
                                                  Else
                                                                            WindowsMediaPlayer1.URL = a
                                                                       Label16.Caption = "正在播放：" & Label1.Caption
                                                                       simble = 1
                                                                      
                                                 End If
                       ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                  If Label4.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = d
                                                        Label16.Caption = "正在播放：" & Label4.Caption
                                                        simble = 4
     
                                                  Else
                                                             WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                 End If
                       ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                If Label5.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = e
                                                        Label16.Caption = "正在播放：" & Label5.Caption
                                                        simble = 5
   
                                                 Else
                                                             WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                 End If
                    ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                                 If Label6.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = f
                                                        Label16.Caption = "正在播放：" & Label6.Caption
                                                        simble = 6
     
                                                Else
                                                             WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                End If
                 ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
                                                 If Label7.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = g
                                                        Label16.Caption = "正在播放：" & Label7.Caption
                                                        simble = 7
      
                                               Else
                                                            WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                End If
                         ElseIf WindowsMediaPlayer1.URL = g And simble = 7 Then
                                                 If Label8.Caption <> "" Then
                                                            WindowsMediaPlayer1.URL = h
                                                            Label16.Caption = "正在播放：" & Label8.Caption
                                                            simble = 8
   
                                                    Else
                                                                 WindowsMediaPlayer1.URL = a
                                                            Label16.Caption = "正在播放：" & Label1.Caption
                                                            simble = 1
     
                                                 End If
                         ElseIf WindowsMediaPlayer1.URL = h And simble = 8 Then
                                                 If Label9.Caption <> "" Then
                                                                WindowsMediaPlayer1.URL = i
                                                                Label16.Caption = "正在播放：" & Label9.Caption
                                                                simble = 9
    
                                                    Else
                                                                     WindowsMediaPlayer1.URL = a
                                                                Label16.Caption = "正在播放：" & Label1.Caption
                                                                simble = 1
 
          
                                                     End If
                           ElseIf WindowsMediaPlayer1.URL = i And simble = 9 Then
                                                     If Label10.Caption <> "" Then
                                                                WindowsMediaPlayer1.URL = j
                                                                Label16.Caption = "正在播放：" & Label10.Caption
                                                                simble = 10
   
                                                    Else
                                                                         WindowsMediaPlayer1.URL = a
                                                                    Label16.Caption = "正在播放：" & Label1.Caption
                                                                    simble = 1
                                                                
                                                    End If
                         ElseIf WindowsMediaPlayer1.URL = j And simble = 10 Then
                                                                    WindowsMediaPlayer1.URL = a
                                                                    Label16.Caption = "正在播放：" & Label1.Caption
                                                                    simble = 1
  
                        

'我的收藏——————————————————————————————————————————顺序循环播放—————————我的收藏———————

                         ElseIf WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                  If Label24.Caption <> "" Then
                                                                    WindowsMediaPlayer1.URL = l
                                                                    Label16.Caption = "正在播放：" & Label24.Caption
                                                                    simble = 12
                    
       
                                                 Else
                                                                    WindowsMediaPlayer1.URL = k
                                                                    Label16.Caption = "正在播放：" & Label23.Caption
                                                                    simble = 11
        
                                                   End If

    
                        ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                    If Label25.Caption <> "" Then
                                                                    WindowsMediaPlayer1.URL = ma
                                                                    Label16.Caption = "正在播放：" & Label25.Caption
                                                                    simble = 13
        
       
                                                     Else
                                                                    WindowsMediaPlayer1.URL = k
                                                                    Label16.Caption = "正在播放：" & Label23.Caption
                                                                    simble = 11
       
                                                    End If

      
                                ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                     If Label26.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = n
                                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                                        simble = 14
        
  
                                                        Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        
   
                                                        End If

      
                            ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                    If Label27.Caption <> "" Then
                                                                            WindowsMediaPlayer1.URL = o
                                                                        Label16.Caption = "正在播放：" & Label27.Caption
                                                                        simble = 15
        
  
                                                    Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        

                                                    End If

      
                             ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                                         If Label28.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = p
                                                                        Label16.Caption = "正在播放：" & Label28.Caption
                                                                        simble = 16

                                                         Else
                                                                            WindowsMediaPlayer1.URL = k
                                                                            Label16.Caption = "正在播放：" & Label23.Caption
                                                                            simble = 11
        

                                                        End If

      
                             ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
                                                          If Label29.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = q
                                                                        Label16.Caption = "正在播放：" & Label29.Caption
                                                                        simble = 17
        
 
                                                        Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        

                                                        End If

      
                       ElseIf WindowsMediaPlayer1.URL = q And simble = 17 Then
                                                         If Label30.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = r
                                                                        Label16.Caption = "正在播放：" & Label30.Caption
                                                                        simble = 18
        
    
                                                        Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11

                                                        End If

      
                           ElseIf WindowsMediaPlayer1.URL = r And simble = 18 Then
                                                    If Label31.Caption <> "" Then
                                                                            WindowsMediaPlayer1.URL = s
                                                                            Label16.Caption = "正在播放：" & Label31.Caption
                                                                            simble = 19
        
      
                                                    Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        

                                                    End If

      
                             ElseIf WindowsMediaPlayer1.URL = s And simble = 19 Then
                                                    If Label32.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = t
                                                                        Label16.Caption = "正在播放：" & Label32.Caption
                                                                        simble = 20
        
   
                                                    Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        
    
                                                    End If



                        ElseIf WindowsMediaPlayer1.URL = t And simble = 20 Then
      
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                                        simble = 11
        
 
                        End If
                        
                   Timer5.Enabled = True
End Sub

Private Sub Label1_Click()
   Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label1_DblClick()
If Label1.Caption <> "" Then
WindowsMediaPlayer1.URL = a
Label16.Caption = "正在播放：" & Label1.Caption
simble = 1

Timer1.Enabled = True

End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label1.Caption <> "" Then

        If Button = 1 Then
           Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then
           Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
            
              Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

           PopupMenu Form10.right
           
    End If
ElseIf Label1.Caption = "" Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
           
End If

       
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label10_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label10_DblClick()
If Label10.Caption <> "" Then
WindowsMediaPlayer1.URL = j
Label16.Caption = "正在播放：" & Label10.Caption
simble = 10

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label10.Caption <> "" Then
        If Button = 1 Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFF8080
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 13
           Label10.ForeColor = RGB(255, 255, 255)
    ElseIf Button = 2 Then
 
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFF8080

           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 13
           Label10.ForeColor = RGB(255, 255, 255)
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
           PopupMenu Form10.right
   
   End If
   
ElseIf Label10.Caption = "" Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
 End If
   
End Sub



Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label11_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
End Sub

Private Sub Label12_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label13_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label14_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   mouse_x = X
   mouse_y = Y
End If

End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label20.ForeColor = &H4040&
Label21.ForeColor = &H808000

Picture22.Visible = False
Picture24.Visible = False
Picture26.Visible = False
Picture28.Visible = False

If Button = 1 Then
   Form1.Left = Form1.Left + X - mouse_x
   Form1.Top = Form1.Top + Y - mouse_y
   
   Form2.Left = Form1.Left + 4470
   Form2.Top = Form1.Top
End If
End Sub

Private Sub Label16_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label2_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label2_DblClick()
If Label2.Caption <> "" Then
WindowsMediaPlayer1.URL = b
Label16.Caption = "正在播放：" & Label2.Caption
simble = 2

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label2.Caption <> "" Then
        If Button = 1 Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFF8080
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 13
           Label2.ForeColor = RGB(255, 255, 255)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then

           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFF8080
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF

           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 13
           Label2.ForeColor = RGB(255, 255, 255)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
   Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
           PopupMenu Form10.right
    End If
ElseIf Label2.Caption = "" Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
 End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label20_Click()
Form3.Show

If Label22.Caption <> "" Then
   Form3.Picture5.Top = 840
   Form3.Picture6.Top = 840
   Form3.Picture7.Top = 840
   Form3.Label4.Visible = True
   Form3.Label4.Caption = "Hello_" & Label22.Caption
   
   With Form3
                .Label2.Visible = False
                .Label13.Visible = False
                .Option1.Visible = False
               .Option2.Visible = False
               .Option3.Visible = False
               .Option4.Visible = False
               .Check4.Visible = False
               
               .Label1.Visible = False
             .Label3.Visible = False
       .Label4.Visible = True
    

        .Label2.Visible = False
       .Label13.Visible = False
        .Option1.Visible = False
        .Option2.Visible = False
        .Option3.Visible = False
        .Option4.Visible = False
        .Check4.Visible = False
        
        
        .Label5.Visible = True
        .Label6.Visible = True
        .Check1.Visible = True
        .Check2.Visible = True
        .Image1.Visible = True
        .Picture15.Visible = True
        .Label12.Visible = True
        .Label8.Visible = True
        .Label9.Visible = True
        .Label10.Visible = True
        .Label14.Visible = True
       .Text1.Visible = True
      
        
        
       .Picture9.Visible = False
           End With
End If
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.ForeColor = vbRed
End Sub

Private Sub Label21_Click()
If Label60.Caption <> "0" Then
  Form7.Timer1.Enabled = True
End If
Form7.Show 1
End Sub

Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label21.ForeColor = vbRed
End Sub

Private Sub Label23_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label23_DblClick()

If Label23.Caption <> "" Then
WindowsMediaPlayer1.URL = k
Label16.Caption = "正在播放：" & Label23.Caption
simble = 11

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label23_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label23.Caption <> "" Then

        If Button = 1 Then
           Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then
           Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
            
              Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

           PopupMenu Form10.right
           
    End If
ElseIf Label23.Caption = "" Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
           
End If
End Sub

Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label24_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label24_DblClick()
If Label24.Caption <> "" Then
WindowsMediaPlayer1.URL = l
Label16.Caption = "正在播放：" & Label24.Caption
simble = 12

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label24.Caption <> "" Then

        If Button = 1 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = vbRed
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 13
           Label24.ForeColor = RGB(255, 255, 255)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = vbRed
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 13
           Label24.ForeColor = RGB(255, 255, 255)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
            
              Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

           PopupMenu Form10.right
           
    End If
ElseIf Label24.Caption = "" Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
           
End If
End Sub

Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label25_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label25_DblClick()
If Label25.Caption <> "" Then
WindowsMediaPlayer1.URL = ma
Label16.Caption = "正在播放：" & Label25.Caption
simble = 13

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label25.Caption <> "" Then

        If Button = 1 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = vbRed
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 13
           Label25.ForeColor = RGB(255, 255, 255)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then
            Label23.BackColor = &H8000000F
           Label25.BackColor = vbRed
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 13
           Label25.ForeColor = RGB(255, 255, 255)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
            
              Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

           PopupMenu Form10.right
           
    End If
ElseIf Label25.Caption = "" Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
           
End If
End Sub

Private Sub Label25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label26_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label26_DblClick()
If Label26.Caption <> "" Then
WindowsMediaPlayer1.URL = n
Label16.Caption = "正在播放：" & Label26.Caption
simble = 14

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label26.Caption <> "" Then

        If Button = 1 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = vbRed
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 13
           Label26.ForeColor = RGB(255, 255, 255)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = vbRed
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 13
           Label26.ForeColor = RGB(255, 255, 255)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
            
              Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

           PopupMenu Form10.right
           
    End If
ElseIf Label26.Caption = "" Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
           
End If
End Sub

Private Sub Label26_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label27_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label27_DblClick()
If Label27.Caption <> "" Then
WindowsMediaPlayer1.URL = o
Label16.Caption = "正在播放：" & Label27.Caption
simble = 15

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label27_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label27.Caption <> "" Then

        If Button = 1 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = vbRed
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 13
           Label27.ForeColor = RGB(255, 255, 255)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = vbRed
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 13
           Label27.ForeColor = RGB(255, 255, 255)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
            
              Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

           PopupMenu Form10.right
           
    End If
ElseIf Label27.Caption = "" Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
           
End If
End Sub

Private Sub Label27_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label28_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label28_DblClick()
If Label28.Caption <> "" Then
WindowsMediaPlayer1.URL = p
Label16.Caption = "正在播放：" & Label28.Caption
simble = 16

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label28.Caption <> "" Then

        If Button = 1 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = vbRed
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 13
           Label28.ForeColor = RGB(255, 255, 255)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = vbRed
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 13
           Label28.ForeColor = RGB(255, 255, 255)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
            
              Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

           PopupMenu Form10.right
           
    End If
ElseIf Label28.Caption = "" Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
           
End If
End Sub

Private Sub Label28_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label29_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label29_DblClick()
If Label29.Caption <> "" Then
WindowsMediaPlayer1.URL = q
Label16.Caption = "正在播放：" & Label29.Caption
simble = 17

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label29_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label29.Caption <> "" Then

        If Button = 1 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = vbRed
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 13
           Label29.ForeColor = RGB(255, 255, 255)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = vbRed
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 13
           Label29.ForeColor = RGB(255, 255, 255)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
            
              Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

           PopupMenu Form10.right
           
    End If
ElseIf Label29.Caption = "" Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
           
End If
End Sub

Private Sub Label29_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label3_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label3_DblClick()
If Label3.Caption <> "" Then
WindowsMediaPlayer1.URL = c
Label16.Caption = "正在播放：" & Label3.Caption
simble = 3

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label3.Caption <> "" Then
        If Button = 1 Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &HFF8080
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
                           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 13
           Label3.ForeColor = RGB(255, 255, 255)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then

           Label1.BackColor = &H8000000F
           Label3.BackColor = &HFF8080
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
            
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 13
           Label3.ForeColor = RGB(255, 255, 255)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
           PopupMenu Form10.right
    End If
ElseIf Label3.Caption = "" Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
 End If
    
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label30_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label30_DblClick()
If Label30.Caption <> "" Then
WindowsMediaPlayer1.URL = r
Label16.Caption = "正在播放：" & Label30.Caption
simble = 18

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label30_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label30.Caption <> "" Then

        If Button = 1 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = vbRed
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 13
           Label30.ForeColor = RGB(255, 255, 255)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = vbRed
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 13
           Label30.ForeColor = RGB(255, 255, 255)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
            
              Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

           PopupMenu Form10.right
           
    End If
ElseIf Label30.Caption = "" Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
           
End If
End Sub

Private Sub Label30_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label31_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label31_DblClick()
If Label31.Caption <> "" Then
WindowsMediaPlayer1.URL = s
Label16.Caption = "正在播放：" & Label31.Caption
simble = 19

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label31_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label31.Caption <> "" Then

        If Button = 1 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = vbRed

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 13
           Label31.ForeColor = RGB(255, 255, 255)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then
            Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = vbRed

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 13
           Label31.ForeColor = RGB(255, 255, 255)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
            
              Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

           PopupMenu Form10.right
           
    End If
ElseIf Label31.Caption = "" Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
           
End If
End Sub

Private Sub Label31_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label32_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label32_DblClick()
If Label32.Caption <> "" Then
WindowsMediaPlayer1.URL = t
Label16.Caption = "正在播放：" & Label32.Caption
simble = 20

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label32_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label32.Caption <> "" Then

        If Button = 1 Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = vbRed
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 13
           Label32.ForeColor = RGB(255, 255, 255)
    ElseIf Button = 2 Then
            Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = vbRed
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 13
           Label32.ForeColor = RGB(255, 255, 255)
            
              Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

           PopupMenu Form10.right
           
    End If
ElseIf Label32.Caption = "" Then
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
           
End If
End Sub

Private Sub Label32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label36_Change()
            If Form1.Label1.BackColor = &HFF8080 Then
                a = Label36.Caption
            ElseIf Form1.Label2.BackColor = &HFF8080 Then
                b = Label36.Caption
            ElseIf Form1.Label3.BackColor = &HFF8080 Then
                c = Label36.Caption
            ElseIf Form1.Label4.BackColor = &HFF8080 Then
               d = Label36.Caption
            ElseIf Form1.Label5.BackColor = &HFF8080 Then
                e = Label36.Caption
            ElseIf Form1.Label6.BackColor = &HFF8080 Then
                f = Label36.Caption
            ElseIf Form1.Label7.BackColor = &HFF8080 Then
                g = Label36.Caption
            ElseIf Form1.Label8.BackColor = &HFF8080 Then
                h = Label36.Caption
            ElseIf Form1.Label9.BackColor = &HFF8080 Then
                i = Label36.Caption
            ElseIf Form1.Label10.BackColor = &HFF8080 Then
               j = Label36.Caption
                
            ElseIf Form1.Label23.BackColor = vbRed Then
                k = Label36.Caption
            ElseIf Form1.Label24.BackColor = vbRed Then
                l = Label36.Caption
            ElseIf Form1.Label25.BackColor = vbRed Then
               ma = Label36.Caption
            ElseIf Form1.Label26.BackColor = vbRed Then
               n = Label36.Caption
            ElseIf Form1.Label27.BackColor = vbRed Then
               o = Label36.Caption
            ElseIf Form1.Label28.BackColor = vbRed Then
              p = Label36.Caption
            ElseIf Form1.Label29.BackColor = vbRed Then
              q = Label36.Caption
            ElseIf Form1.Label30.BackColor = vbRed Then
               r = Label36.Caption
            ElseIf Form1.Label31.BackColor = vbRed Then
               s = Label36.Caption
            ElseIf Form1.Label32.BackColor = vbRed Then
               t = Label36.Caption
            End If
End Sub

Private Sub Label39_Change()
If WindowsMediaPlayer1.playState <> wmppsPaused Then
    If Label46.Caption = "1" Then
       WindowsMediaPlayer1.URL = a
       Label16.Caption = "正在播放：" & Label1.Caption
       simble = "1"
       Timer1.Enabled = True
       Timer2.Enabled = True
    ElseIf Label46.Caption = "2" Then
       WindowsMediaPlayer1.URL = b
       Label16.Caption = "正在播放：" & Label2.Caption
       simble = "2"
       Timer1.Enabled = True
        Timer2.Enabled = True
    ElseIf Label46.Caption = "3" Then
       WindowsMediaPlayer1.URL = c
       Label16.Caption = "正在播放：" & Label3.Caption
       simble = "3"
       Timer1.Enabled = True
        Timer2.Enabled = True
    ElseIf Label46.Caption = "4" Then
       WindowsMediaPlayer1.URL = d
       Label16.Caption = "正在播放：" & Label4.Caption
       simble = "4"
       Timer1.Enabled = True
        Timer2.Enabled = True
    ElseIf Label46.Caption = "5" Then
       WindowsMediaPlayer1.URL = e
       Label16.Caption = "正在播放：" & Label5.Caption
       simble = "5"
       Timer1.Enabled = True
        Timer2.Enabled = True
    ElseIf Label46.Caption = "6" Then
       WindowsMediaPlayer1.URL = f
       Label16.Caption = "正在播放：" & Label6.Caption
       simble = "6"
       Timer1.Enabled = True
        Timer2.Enabled = True
    ElseIf Label46.Caption = "7" Then
       WindowsMediaPlayer1.URL = g
       Label16.Caption = "正在播放：" & Label7.Caption
       simble = "7"
       Timer1.Enabled = True
        Timer2.Enabled = True
    ElseIf Label46.Caption = "8" Then
       WindowsMediaPlayer1.URL = h
       Label16.Caption = "正在播放：" & Label8.Caption
       simble = "8"
       Timer1.Enabled = True
        Timer2.Enabled = True
    ElseIf Label46.Caption = "9" Then
       WindowsMediaPlayer1.URL = i
       Label16.Caption = "正在播放：" & Label9.Caption
       simble = "9"
       Timer1.Enabled = True
        Timer2.Enabled = True
    ElseIf Label46.Caption = "10" Then
       WindowsMediaPlayer1.URL = j
       Label16.Caption = "正在播放：" & Label10.Caption
       simble = "10"
       Timer1.Enabled = True
        Timer2.Enabled = True
    End If
    '我的收藏
    If Label46.Caption = "11" Then
       WindowsMediaPlayer1.URL = k
       Label16.Caption = "正在播放：" & Label23.Caption
       simble = "11"
       Timer1.Enabled = True
       Timer2.Enabled = True
    ElseIf Label46.Caption = "12" Then
       WindowsMediaPlayer1.URL = l
       Label16.Caption = "正在播放：" & Label24.Caption
       simble = "12"
       Timer1.Enabled = True
       Timer2.Enabled = True
    ElseIf Label46.Caption = "13" Then
       WindowsMediaPlayer1.URL = ma
       Label16.Caption = "正在播放：" & Label25.Caption
       simble = "13"
       Timer1.Enabled = True
       Timer2.Enabled = True
    ElseIf Label46.Caption = "14" Then
       WindowsMediaPlayer1.URL = n
       Label16.Caption = "正在播放：" & Label26.Caption
       simble = "14"
       Timer1.Enabled = True
       Timer2.Enabled = True
    ElseIf Label46.Caption = "15" Then
       WindowsMediaPlayer1.URL = l
       Label16.Caption = "正在播放：" & Label27.Caption
       simble = "15"
       Timer1.Enabled = True
       Timer2.Enabled = True
    ElseIf Label46.Caption = "16" Then
       WindowsMediaPlayer1.URL = l
       Label16.Caption = "正在播放：" & Label28.Caption
       simble = "16"
       Timer1.Enabled = True
       Timer2.Enabled = True
    ElseIf Label46.Caption = "17" Then
       WindowsMediaPlayer1.URL = l
       Label16.Caption = "正在播放：" & Label29.Caption
       simble = "17"
       Timer1.Enabled = True
       Timer2.Enabled = True
    ElseIf Label46.Caption = "18" Then
       WindowsMediaPlayer1.URL = l
       Label16.Caption = "正在播放：" & Label30.Caption
       simble = "18"
       Timer1.Enabled = True
       Timer2.Enabled = True
    ElseIf Label46.Caption = "19" Then
       WindowsMediaPlayer1.URL = l
       Label16.Caption = "正在播放：" & Label31.Caption
       simble = "19"
       Timer1.Enabled = True
       Timer2.Enabled = True
    ElseIf Label46.Caption = "20" Then
       WindowsMediaPlayer1.URL = l
       Label16.Caption = "正在播放：" & Label32.Caption
       simble = "20"
       Timer1.Enabled = True
       Timer2.Enabled = True
    End If
Else
 WindowsMediaPlayer1.Controls.play
End If
 

If simble = "1" Then
           Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "2" Then
 Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFF8080
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 13
           Label2.ForeColor = RGB(255, 255, 255)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "3" Then
Label1.BackColor = &H8000000F
           Label3.BackColor = &HFF8080
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
                           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 13
           Label3.ForeColor = RGB(255, 255, 255)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "4" Then
 Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFF8080
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 13
           Label4.ForeColor = RGB(255, 255, 255)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "5" Then
Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &HFF8080
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 13
           Label5.ForeColor = RGB(255, 255, 255)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "6" Then
 Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFF8080
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 13
           Label6.ForeColor = RGB(255, 255, 255)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "7" Then
Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &HFF8080
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 13
           Label7.ForeColor = RGB(255, 255, 255)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "8" Then
Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFF8080
           Label10.BackColor = &HFFC0FF
                        
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 13
           Label8.ForeColor = RGB(255, 255, 255)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "9" Then
Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &HFF8080

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 13
           Label9.ForeColor = RGB(255, 255, 255)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "10" Then
 Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFF8080
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 13
           Label10.ForeColor = RGB(255, 255, 255)
ElseIf simble = "11" Then
Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "12" Then
 Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = vbRed
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 13
           Label24.ForeColor = RGB(255, 255, 255)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "13" Then
 Label23.BackColor = &H8000000F
           Label25.BackColor = vbRed
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 13
           Label25.ForeColor = RGB(255, 255, 255)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "14" Then
Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = vbRed
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 13
           Label26.ForeColor = RGB(255, 255, 255)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "15" Then
Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = vbRed
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 13
           Label27.ForeColor = RGB(255, 255, 255)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "16" Then
 Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = vbRed
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 13
           Label28.ForeColor = RGB(255, 255, 255)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "17" Then
Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = vbRed
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 13
           Label29.ForeColor = RGB(255, 255, 255)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "18" Then
Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = vbRed
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 13
           Label30.ForeColor = RGB(255, 255, 255)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "19" Then
 Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = vbRed

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 13
           Label31.ForeColor = RGB(255, 255, 255)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "20" Then
Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = vbRed
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 13
           Label32.ForeColor = RGB(255, 255, 255)
End If
Timer3.Enabled = True
End Sub

Private Sub Label4_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label4_DblClick()
If Label4.Caption <> "" Then
WindowsMediaPlayer1.URL = d
Label16.Caption = "正在播放：" & Label4.Caption
simble = 4

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label4.Caption <> "" Then
        If Button = 1 Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFF8080
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 13
           Label4.ForeColor = RGB(255, 255, 255)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then

           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFF8080
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF

           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 13
           Label4.ForeColor = RGB(255, 255, 255)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
           PopupMenu Form10.right
    End If
ElseIf Label4.Caption = "" Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
 End If
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label40_Change()
On Error Resume Next
If Label46.Caption = "1" Then
   Label1.Caption = ""
ElseIf Label46.Caption = "2" Then
   Label2.Caption = ""
ElseIf Label46.Caption = "3" Then
   Label3.Caption = ""
ElseIf Label46.Caption = "4" Then
   Label4.Caption = ""
ElseIf Label46.Caption = "5" Then
   Label5.Caption = ""
ElseIf Label46.Caption = "6" Then
   Label6.Caption = ""
ElseIf Label46.Caption = "7" Then
   Label7.Caption = ""
ElseIf Label46.Caption = "8" Then
   Label8.Caption = ""
ElseIf Label46.Caption = "9" Then
   Label9.Caption = ""
ElseIf Label46.Caption = "10" Then
   Label10.Caption = ""
End If
'我的收藏
If Label46.Caption = "11" Then
   Label23.Caption = ""
ElseIf Label46.Caption = "12" Then
   Label24.Caption = ""
ElseIf Label46.Caption = "13" Then
   Label25.Caption = ""
ElseIf Label46.Caption = "14" Then
   Label26.Caption = ""
ElseIf Label46.Caption = "15" Then
   Label27.Caption = ""
ElseIf Label46.Caption = "16" Then
   Label28.Caption = ""
ElseIf Label46.Caption = "17" Then
   Label29.Caption = ""
ElseIf Label46.Caption = "18" Then
   Label30.Caption = ""
ElseIf Label46.Caption = "19" Then
   Label31.Caption = ""
ElseIf Label46.Caption = "20" Then
   Label32.Caption = ""
End If

If Form1.Label1.Caption = "" Then
   m = a
   Form1.Label1.Caption = Label2.Caption
   a = b
   Label2.Caption = Label3.Caption
   b = c
   Label3.Caption = Label4.Caption
   c = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "1" Then
     If Label1.Caption <> "" Then
                WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
     Else
        WindowsMediaPlayer1.URL = ""
        Label16.Caption = "百度音乐_音乐你的生活"
        Label16.Alignment = 2
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label2.Caption = "" Then
   m = b
   Label2.Caption = Label3.Caption
   b = c
   Label3.Caption = Label4.Caption
   c = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "2" Then
       If Label2.Caption <> "" Then
                WindowsMediaPlayer1.URL = b
                Label16.Caption = "正在播放：" & Label2.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
           Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label3.Caption = "" Then
   m = c
   Label3.Caption = Label4.Caption
   c = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "3" Then
       If Label3.Caption <> "" Then
                WindowsMediaPlayer1.URL = c
                Label16.Caption = "正在播放：" & Label3.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
           Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label4.Caption = "" Then
   m = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "4" Then
       If Label4.Caption <> "" Then
                WindowsMediaPlayer1.URL = d
                Label16.Caption = "正在播放：" & Label4.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
           Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label5.Caption = "" Then
   m = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "5" Then
       If Label5.Caption <> "" Then
                WindowsMediaPlayer1.URL = e
                Label16.Caption = "正在播放：" & Label5.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
            Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label6.Caption = "" Then
   m = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "6" Then
       If Label6.Caption <> "" Then
                WindowsMediaPlayer1.URL = f
                Label16.Caption = "正在播放：" & Label6.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
          Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label7.Caption = "" Then
   m = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "7" Then
       If Label7.Caption <> "" Then
                WindowsMediaPlayer1.URL = g
                Label16.Caption = "正在播放：" & Label7.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
          Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label8.Caption = "" Then
   m = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "8" Then
       If Label8.Caption <> "" Then
                WindowsMediaPlayer1.URL = h
                Label16.Caption = "正在播放：" & Label8.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
           Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label9.Caption = "" Then
   m = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "9" Then
       If Label9.Caption <> "" Then
                WindowsMediaPlayer1.URL = i
                Label16.Caption = "正在播放：" & Label9.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
           Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
End If
'我的收藏————————————————————————————————————————————————————
If Form1.Label23.Caption = "" Then
   m = k
   Label23.Caption = Label24.Caption
   k = l
   Label24.Caption = Label25.Caption
   l = ma
   Label25.Caption = Label26.Caption
   ma = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "11" Then
       If Label23.Caption <> "" Then
                WindowsMediaPlayer1.URL = k
                Label16.Caption = "正在播放：" & Label23.Caption
       Else
                 WindowsMediaPlayer1.URL = ""
                 Label16.Caption = "百度音乐_音乐你的生活"
                 Label16.Alignment = 2
                   Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label24.Caption = "" Then
   m = l
   Label24.Caption = Label25.Caption
   l = ma
   Label25.Caption = Label26.Caption
   ma = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "12" Then
     If Label24.Caption <> "" Then
            WindowsMediaPlayer1.URL = l
            Label16.Caption = "正在播放：" & Label24.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                   Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label25.Caption = "" Then
   m = ma
   Label25.Caption = Label26.Caption
   ma = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "13" Then
     If Label25.Caption <> "" Then
            WindowsMediaPlayer1.URL = ma
            Label16.Caption = "正在播放：" & Label25.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                  Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label26.Caption = "" Then
   m = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "14" Then
     If Label26.Caption <> "" Then
            WindowsMediaPlayer1.URL = n
            Label16.Caption = "正在播放：" & Label26.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                   Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label27.Caption = "" Then
   m = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "15" Then
     If Label27.Caption <> "" Then
            WindowsMediaPlayer1.URL = o
            Label16.Caption = "正在播放：" & Label27.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                   Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label28.Caption = "" Then
   m = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "16" Then
     If Label28.Caption <> "" Then
            WindowsMediaPlayer1.URL = p
            Label16.Caption = "正在播放：" & Label28.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                   Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label29.Caption = "" Then
   m = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "17" Then
     If Label29.Caption <> "" Then
            WindowsMediaPlayer1.URL = q
            Label16.Caption = "正在播放：" & Label29.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                   Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label30.Caption = "" Then
   m = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "18" Then
     If Label30.Caption <> "" Then
            WindowsMediaPlayer1.URL = r
            Label16.Caption = "正在播放：" & Label30.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                  Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label31.Caption = "" Then
   m = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "19" Then
     If Label31.Caption <> "" Then
            WindowsMediaPlayer1.URL = s
            Label16.Caption = "正在播放：" & Label31.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                  Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
End If


If Label1.Caption = "" Then
   Label1.BackColor = &H8000000F
   Label1.FontSize = 9
   Label1.ForeColor = RGB(0, 0, 0)
ElseIf Label2.Caption = "" Then
  Label2.BackColor = &HFFC0FF
  Label2.FontSize = 9
  Label2.ForeColor = RGB(0, 0, 0)
ElseIf Label3.Caption = "" Then
  Label3.BackColor = &H8000000F
  Label3.FontSize = 9
  Label3.ForeColor = RGB(0, 0, 0)
ElseIf Label4.Caption = "" Then
  Label4.BackColor = &HFFC0FF
  Label4.FontSize = 9
  Label4.ForeColor = RGB(0, 0, 0)
ElseIf Label5.Caption = "" Then
  Label5.BackColor = &H8000000F
  Label5.FontSize = 9
  Label5.ForeColor = RGB(0, 0, 0)
ElseIf Label6.Caption = "" Then
  Label6.BackColor = &HFFC0FF
  Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
ElseIf Label7.Caption = "" Then
  Label7.BackColor = &H8000000F
  Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
ElseIf Label8.Caption = "" Then
  Label8.BackColor = &HFFC0FF
  Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
ElseIf Label9.Caption = "" Then
  Label9.BackColor = &H8000000F
  Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
ElseIf Label10.Caption = "" Then
  Label10.BackColor = &HFFC0FF
  Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
End If
'my count________________________________________________________
If Label23.Caption = "" Then
  Label23.BackColor = &H8000000F
  Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
ElseIf Label24.Caption = "" Then
Label24.BackColor = &HFFC0FF
Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
ElseIf Label25.Caption = "" Then
Label25.BackColor = &H8000000F
 Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
ElseIf Label26.Caption = "" Then
 Label26.BackColor = &HFFC0FF
 Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
ElseIf Label27.Caption = "" Then
Label27.BackColor = &H8000000F
Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
ElseIf Label28.Caption = "" Then
Label28.BackColor = &HFFC0FF
Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
ElseIf Label29.Caption = "" Then
Label29.BackColor = &H8000000F
Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
ElseIf Label30.Caption = "" Then
Label30.BackColor = &HFFC0FF
Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
ElseIf Label31.Caption = "" Then
Label31.BackColor = &H8000000F
Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
ElseIf Label32.Caption = "" Then
Label32.BackColor = &HFFC0FF
 Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)

End If
           

End Sub

Private Sub Label41_Change()

Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0

Label53.BorderStyle = 0
Label54.BorderStyle = 0


CommonDialog1.CancelError = True
On Error GoTo errline
CommonDialog1.ShowOpen
'CommonDialog1.Filter = "音乐文件 （*.mp3;*.wma）|*.mp3;*.wma"
WindowsMediaPlayer1.URL = CommonDialog1.FileName
Label16.Caption = "正在播放：" & CommonDialog1.FileTitle

If Label1.Caption = "" Then
   Label1.Caption = " " & CommonDialog1.FileTitle
   a = CommonDialog1.FileName
   simble = 1
           Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf Label1.Caption <> "" And Label2.Caption = "" Then
   Label2.Caption = " " & CommonDialog1.FileTitle
   b = CommonDialog1.FileName
   simble = 2
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFF8080
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 13
           Label2.ForeColor = RGB(255, 255, 255)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption = "" Then
   Label3.Caption = " " & CommonDialog1.FileTitle
   c = CommonDialog1.FileName
   simble = 3
    Label1.BackColor = &H8000000F
           Label3.BackColor = &HFF8080
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
                           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 13
           Label3.ForeColor = RGB(255, 255, 255)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption = "" Then
   Label4.Caption = " " & CommonDialog1.FileTitle
   d = CommonDialog1.FileName
   simble = 4
   Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFF8080
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 13
           Label4.ForeColor = RGB(255, 255, 255)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption = "" Then
   Label5.Caption = " " & CommonDialog1.FileTitle
   e = CommonDialog1.FileName
   simble = 5
   Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &HFF8080
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 13
           Label5.ForeColor = RGB(255, 255, 255)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption = "" Then
   Label6.Caption = " " & CommonDialog1.FileTitle
   f = CommonDialog1.FileName
   simble = 6
   Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFF8080
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 13
           Label6.ForeColor = RGB(255, 255, 255)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption = "" Then
   Label7.Caption = " " & CommonDialog1.FileTitle
   g = CommonDialog1.FileName
   simble = 7
   Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &HFF8080
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 13
           Label7.ForeColor = RGB(255, 255, 255)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption = "" Then
   Label8.Caption = " " & CommonDialog1.FileTitle
   h = CommonDialog1.FileName
   simble = 8
   Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFF8080
           Label10.BackColor = &HFFC0FF
                        
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 13
           Label8.ForeColor = RGB(255, 255, 255)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption <> "" And Label9.Caption = "" Then
   Label9.Caption = " " & CommonDialog1.FileTitle
   i = CommonDialog1.FileName
   simble = 9
   Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &HFF8080

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 13
           Label9.ForeColor = RGB(255, 255, 255)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption <> "" And Label9.Caption <> "" And Label10.Caption = "" Then
   Label10.Caption = " " & CommonDialog1.FileTitle
   j = CommonDialog1.FileName
   simble = 10
   Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFF8080
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 13
           Label10.ForeColor = RGB(255, 255, 255)
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption <> "" And Label9.Caption <> "" And Label10.Caption <> "" Then
           Label10.Caption = " " & CommonDialog1.FileTitle
         j = CommonDialog1.FileName
         simble = 10
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFF8080
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 13
           Label10.ForeColor = RGB(255, 255, 255)
End If
 Picture9.Visible = False
 Picture6.Visible = True
 Timer1.Enabled = True
 Timer2.Enabled = True
 Picture19.Visible = False
 
   Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
   
   
   Label23.Visible = False
   Label24.Visible = False
   Label25.Visible = False
   Label26.Visible = False
   Label27.Visible = False
   Label28.Visible = False
   Label29.Visible = False
   Label30.Visible = False
   Label31.Visible = False
   Label32.Visible = False
   
   Picture13.Visible = False
   Picture2.Visible = True
   Picture15.Visible = False
   Picture3.Visible = True
Exit Sub

errline:
Picture9.Visible = False
 Picture6.Visible = True
 
   Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
Exit Sub
End Sub

Private Sub Label42_Change()
On Error Resume Next
If Label22.Caption <> "" Then
 If Label46.Caption = "1" Then
    If Label23.Caption = "" Then
       Label23.Caption = Label1.Caption
       k = a
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label1.Caption
       l = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label1.Caption
       ma = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label1.Caption
       n = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label1.Caption
       o = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label1.Caption
       p = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label1.Caption
       q = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label1.Caption
       r = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label1.Caption
       s = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label1.Caption
       t = a
    End If
    

  
ElseIf Label46.Caption = "2" Then
   If Label23.Caption = "" Then
       Label23.Caption = Label2.Caption
       k = b
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label2.Caption
       l = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label2.Caption
       ma = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label2.Caption
       n = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label2.Caption
       o = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label2.Caption
       p = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label2.Caption
       q = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label2.Caption
       r = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label2.Caption
       s = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label2.Caption
       t = b
    End If
    

ElseIf Label46.Caption = "3" Then
   If Label23.Caption = "" Then
       Label23.Caption = Label3.Caption
       k = c
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label3.Caption
       l = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label3.Caption
       ma = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label3.Caption
       n = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label3.Caption
       o = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label3.Caption
       p = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label3.Caption
       q = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label3.Caption
       r = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label3.Caption
       s = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label3.Caption
       t = c
    End If
ElseIf Label46.Caption = "4" Then
   If Label23.Caption = "" Then
       Label23.Caption = Label4.Caption
       k = d
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label4.Caption
       l = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label4.Caption
       ma = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label4.Caption
       n = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label4.Caption
       o = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label4.Caption
       p = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label4.Caption
       q = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label4.Caption
       r = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label4.Caption
       s = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label4.Caption
       t = d
    End If
ElseIf Label46.Caption = "5" Then
   If Label23.Caption = "" Then
       Label23.Caption = Label5.Caption
       k = e
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label5.Caption
       l = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label5.Caption
       ma = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label5.Caption
       n = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label5.Caption
       o = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label5.Caption
       p = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label5.Caption
       q = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label5.Caption
       r = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label5.Caption
       s = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label5.Caption
       t = e
    End If
ElseIf Label46.Caption = "6" Then
   If Label23.Caption = "" Then
       Label23.Caption = Label6.Caption
       k = f
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label6.Caption
       l = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label6.Caption
       ma = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label6.Caption
       n = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label6.Caption
       o = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label6.Caption
       p = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label6.Caption
       q = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label6.Caption
       r = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label6.Caption
       s = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label6.Caption
       t = f
    End If
ElseIf Label46.Caption = "7" Then
   If Label23.Caption = "" Then
       Label23.Caption = Label7.Caption
       k = g
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label7.Caption
       l = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label7.Caption
       ma = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label7.Caption
       n = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label7.Caption
       o = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label7.Caption
       p = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label7.Caption
       q = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label7.Caption
       r = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label7.Caption
       s = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label7.Caption
       t = g
    End If
ElseIf Label46.Caption = "8" Then
   If Label23.Caption = "" Then
       Label23.Caption = Label8.Caption
       k = h
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label8.Caption
       l = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label8.Caption
       ma = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label8.Caption
       n = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label8.Caption
       o = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label8.Caption
       p = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label8.Caption
       q = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label8.Caption
       r = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label8.Caption
       s = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label8.Caption
       t = h
    End If
ElseIf Label46.Caption = "9" Then
   If Label23.Caption = "" Then
       Label23.Caption = Label9.Caption
       k = i
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label9.Caption
       l = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label9.Caption
       ma = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label9.Caption
       n = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label9.Caption
       o = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label9.Caption
       p = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label9.Caption
       q = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label9.Caption
       r = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label9.Caption
       s = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label9.Caption
       t = i
    End If
ElseIf Label46.Caption = "10" Then
   If Label23.Caption = "" Then
       Label23.Caption = Label10.Caption
       k = j
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label10.Caption
       l = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label10.Caption
       ma = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label10.Caption
       n = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label10.Caption
       o = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label10.Caption
       p = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label10.Caption
       q = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label10.Caption
       r = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label10.Caption
       s = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label10.Caption
       t = j
    End If
End If

  Label23.Visible = True
  Label24.Visible = True
  Label25.Visible = True
  Label26.Visible = True
  Label27.Visible = True
  Label28.Visible = True
  Label29.Visible = True
  Label30.Visible = True
  Label31.Visible = True
  Label32.Visible = True
  
  Picture2.Visible = False
  Picture13.Visible = True
  Picture3.Visible = False
  Picture15.Visible = True
  
            Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf Label22.Caption = "" Then
    MsgBox "请登录后再添加到‘我的收藏’", , "草哥提示"
End If
End Sub


Private Sub Label43_Change()
            If Form1.Label46.Caption = "1" Then
                Form1.Label35.Caption = a
                Form1.Label37.Caption = 1
            ElseIf Form1.Label46.Caption = "1" Then
                Form1.Label35.Caption = b
                Form1.Label37.Caption = 2
            ElseIf Form1.Label46.Caption = "1" Then
               Form1.Label35.Caption = c
               Form1.Label37.Caption = 3
            ElseIf Form1.Label46.Caption = "1" Then
               Form1.Label35.Caption = d
               Form1.Label37.Caption = 4
            ElseIf Form1.Label46.Caption = "1" Then
                Form1.Label35.Caption = e
                Form1.Label37.Caption = 5
            ElseIf Form1.Label46.Caption = "1" Then
                Form1.Label35.Caption = f
                Form1.Label37.Caption = 6
            ElseIf Form1.Label46.Caption = "1" Then
                Form1.Label35.Caption = g
                Form1.Label37.Caption = 7
            ElseIf Form1.Label46.Caption = "1" Then
                Form1.Label35.Caption = h
                Form1.Label37.Caption = 8
            ElseIf Form1.Label46.Caption = "1" Then
                Form1.Label35.Caption = i
                Form1.Label37.Caption = 9
            ElseIf Form1.Label46.Caption = "1" Then
                Form1.Label35.Caption = j
                Form1.Label37.Caption = 10
           '我的收藏————————————————————————————————————————————
            ElseIf Form1.Label46.Caption = "1" Then
               Form1.Label35.Caption = k
               Form1.Label37.Caption = 11
            ElseIf Form1.Label46.Caption = "1" Then
               Form1.Label35.Caption = l
               Form1.Label37.Caption = 12
            ElseIf Form1.Label46.Caption = "1" Then
               Form1.Label35.Caption = ma
               Form1.Label37.Caption = 13
            ElseIf Form1.Label46.Caption = "1" Then
               Form1.Label35.Caption = n
               Form1.Label37.Caption = 14
            ElseIf Form1.Label46.Caption = "1" Then
               Form1.Label35.Caption = o
               Form1.Label37.Caption = 15
            ElseIf Form1.Label46.Caption = "1" Then
               Form1.Label35.Caption = p
               Form1.Label37.Caption = 16
            ElseIf Form1.Label46.Caption = "1" Then
               Form1.Label35.Caption = q
               Form1.Label37.Caption = 17
            ElseIf Form1.Label46.Caption = "1" Then
              Form1.Label35.Caption = r
              Form1.Label37.Caption = 18
            ElseIf Form1.Label46.Caption = "1" Then
               Form1.Label35.Caption = s
               Form1.Label37.Caption = 19
            ElseIf Form1.Label46.Caption = "1" Then
               Form1.Label35.Caption = t
               Form1.Label37.Caption = 20
            End If

Form5.Show 1
End Sub

Private Sub Label44_Change()
If Label22.Caption <> "" Then
   Open "c:\百度音乐播放器\" & Label22.Caption & "播放列表.txt" For Output As #2
   Print #2, Label1.Caption
   Print #2, a
   Print #2, Label2.Caption
   Print #2, b
   Print #2, Label3.Caption
   Print #2, c
   Print #2, Label4.Caption
   Print #2, d
   Print #2, Label5.Caption
   Print #2, e
   Print #2, Label6.Caption
   Print #2, f
   Print #2, Label7.Caption
   Print #2, g
   Print #2, Label8.Caption
   Print #2, h
   Print #2, Label9.Caption
   Print #2, i
   Print #2, Label10.Caption
   Print #2, j

   Print #2, pic

   Close #2
Else
  MsgBox "未登录，不可保存"
End If

End Sub

Private Sub Label45_Change()
If Picture2.Visible = True Then
        Label1.Caption = ""
        Label2.Caption = ""
        Label3.Caption = ""
        Label4.Caption = ""
        Label5.Caption = ""
        Label6.Caption = ""
        Label7.Caption = ""
        Label8.Caption = ""
        Label9.Caption = ""
        Label10.Caption = ""
        a = ""
        b = ""
        c = ""
        d = ""
        e = ""
        f = ""
        g = ""
        h = ""
        i = ""
        j = ""
ElseIf Picture15.Visible = True Then
        Label23.Caption = ""
        Label24.Caption = ""
        Label25.Caption = ""
        Label26.Caption = ""
        Label27.Caption = ""
        Label28.Caption = ""
        Label29.Caption = ""
        Label30.Caption = ""
        Label31.Caption = ""
        Label32.Caption = ""
        k = ""
        l = ""
        ma = ""
        n = ""
        o = ""
        p = ""
        q = ""
        r = ""
        s = ""
        t = ""
End If

    Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
           
           
           Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub Label47_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   mouse_x = X
   mouse_y = Y
End If
End Sub

Private Sub Label47_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.ForeColor = &H4040&
Label21.ForeColor = &H808000
MousePointer = 1

Picture22.Visible = False
Picture24.Visible = False
Picture26.Visible = False
Picture28.Visible = False

If Button = 1 Then
   Form1.Left = Form1.Left + X - mouse_x
   Form1.Top = Form1.Top + Y - mouse_y
   
   Form2.Left = Form1.Left + 4470
   Form2.Top = Form1.Top
End If
End Sub

Private Sub Label48_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   mouse_x = X
   mouse_y = Y
End If
End Sub

Private Sub Label48_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.ForeColor = &H4040&
Label21.ForeColor = &H808000
MousePointer = 1

Picture22.Visible = False
Picture24.Visible = False
Picture26.Visible = False
Picture28.Visible = False

If Button = 1 Then
   Form1.Left = Form1.Left + X - mouse_x
   Form1.Top = Form1.Top + Y - mouse_y
   
   Form2.Left = Form1.Left + 4470
   Form2.Top = Form1.Top
End If
End Sub

Private Sub Label49_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label49.BorderStyle = 1
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0

If Button = 1 Then
   PopupMenu Form10.addmusic, 0, 40, 6480
End If
End Sub

Private Sub Label49_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label49.BorderStyle = 1
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
Label53.BorderStyle = 0
Label54.BorderStyle = 0
End Sub

Private Sub Label5_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label5_DblClick()
If Label5.Caption <> "" Then
WindowsMediaPlayer1.URL = e
Label16.Caption = "正在播放：" & Label5.Caption
simble = 5

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label5.Caption <> "" Then
        If Button = 1 Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &HFF8080
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 13
           Label5.ForeColor = RGB(255, 255, 255)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then

           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &HFF8080
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF

           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 13
           Label5.ForeColor = RGB(255, 255, 255)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
           PopupMenu Form10.right
    End If
ElseIf Label5.Caption = "" Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
 End If
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label50_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label49.BorderStyle = 0
Label50.BorderStyle = 1
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0

If Button = 1 Then
   PopupMenu Form10.delete2, 0, 1290, 6480
End If
End Sub

Private Sub Label50_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label49.BorderStyle = 0
Label50.BorderStyle = 1
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
Label53.BorderStyle = 0
Label54.BorderStyle = 0
End Sub

Private Sub Label51_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label49.BorderStyle = 0
Label50.BorderStyle = 0
Label51.BorderStyle = 1
Label52.BorderStyle = 0

If Image2.Visible = False Then
Image2.Visible = True
Label53.Visible = True
Label54.Visible = True
Combo1.Visible = True
ElseIf Image2.Visible = True Then
Image2.Visible = False
Label53.Visible = False
Label54.Visible = False
Combo1.Visible = False
End If
End Sub

Private Sub Label51_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label49.BorderStyle = 0
Label50.BorderStyle = 0
Label51.BorderStyle = 1
Label52.BorderStyle = 0
Label53.BorderStyle = 0
Label54.BorderStyle = 0
End Sub

Private Sub Label52_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 1

If Button = 1 Then
   PopupMenu Form10.model, 0, 2500, 6480
End If
End Sub

Private Sub Label52_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 1
Label53.BorderStyle = 0
Label54.BorderStyle = 0
End Sub

Private Sub Label53_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   urlstring = "http://mp3.baidu.com/m?f=ms&tn=baidump3&ct=134217728&lf=&rn=&word=" & Combo1.Text & "&lm=0"
   Form2.Visible = True
   Form2.WebBrowser1.Navigate (urlstring)
   Combo1.AddItem Combo1.Text
   
  Text1.Enabled = False
  If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
  End If
End If

End Sub

Private Sub Label53_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label53.BorderStyle = 1
End Sub

Private Sub Label54_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Label53.Visible = False
Label54.Visible = False
Combo1.Visible = False
End Sub

Private Sub Label54_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label54.BorderStyle = 1
End Sub

Private Sub Label55_Change()
   CommonDialog2.CancelError = True
   On Error GoTo errline
   CommonDialog2.ShowOpen
   'CommonDialog2.Filter = "图片文件 (*.jpg;*.jpeg;*.jpe;*.bmp;*.gif)|*.jpg;*.jpeg;*.jpe;*.bmp;*.gif)"
   Image1.Picture = LoadPicture(CommonDialog2.FileName)
   Form3.Image1.Picture = LoadPicture(CommonDialog2.FileName)
   
     pic = CommonDialog2.FileName
  
   MousePointer = 1
   
     Text1.Enabled = False
   If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
   End If
   
Exit Sub

errline:
   MousePointer = 1
     Text1.Enabled = False
  If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
  End If
   Exit Sub

End Sub

Private Sub Label56_Change()
On Error Resume Next
         If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                                If Label10 <> "" Then
                                                                    WindowsMediaPlayer1.URL = j
                                                                    Label16.Caption = "正在播放：" & Label10.Caption
                                                                    simble = 10
                                                                ElseIf Label10 = "" And Label9 <> "" Then
                                                                    WindowsMediaPlayer1.URL = i
                                                                    Label16.Caption = "正在播放：" & Label9.Caption
                                                                    simble = 9
                                                                ElseIf Label9 = "" And Label8 <> "" Then
                                                                    WindowsMediaPlayer1.URL = h
                                                                    Label16.Caption = "正在播放：" & Label8.Caption
                                                                    simble = 8
                                                                ElseIf Label8 = "" And Label7 <> "" Then
                                                                    WindowsMediaPlayer1.URL = g
                                                                    Label16.Caption = "正在播放：" & Label7.Caption
                                                                    simble = 7
                                                                ElseIf Label7 = "" And Label6 <> "" Then
                                                                    WindowsMediaPlayer1.URL = f
                                                                    Label16.Caption = "正在播放：" & Label6.Caption
                                                                    simble = 6
                                                                ElseIf Label6 = "" And Label5 <> "" Then
                                                                    WindowsMediaPlayer1.URL = e
                                                                    Label16.Caption = "正在播放：" & Label5.Caption
                                                                    simble = 5
                                                                ElseIf Label5 = "" And Label4 <> "" Then
                                                                    WindowsMediaPlayer1.URL = d
                                                                    Label16.Caption = "正在播放：" & Label4.Caption
                                                                    simble = 4
                                                                ElseIf Label4 = "" And Label3 <> "" Then
                                                                    WindowsMediaPlayer1.URL = c
                                                                    Label16.Caption = "正在播放：" & Label3.Caption
                                                                    simble = 3
                                                                ElseIf Label3 = "" And Label2 <> "" Then
                                                                    WindowsMediaPlayer1.URL = b
                                                                    Label16.Caption = "正在播放：" & Label2.Caption
                                                                    simble = 2
                                                                ElseIf Label2 = "" And Label1 <> "" Then
                                                                    WindowsMediaPlayer1.URL = a
                                                                    Label16.Caption = "正在播放：" & Label1.Caption
                                                                    simble = 1
                                                                End If
                      ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                                       WindowsMediaPlayer1.URL = a
                                                                       Label16.Caption = "正在播放：" & Label1.Caption
                                                                       simble = 1
                       ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                        WindowsMediaPlayer1.URL = b
                                                        Label16.Caption = "正在播放：" & Label2.Caption
                                                        simble = 2
                       ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                        WindowsMediaPlayer1.URL = c
                                                        Label16.Caption = "正在播放：" & Label3.Caption
                                                        simble = 3
                    ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                                        WindowsMediaPlayer1.URL = d
                                                        Label16.Caption = "正在播放：" & Label4.Caption
                                                        simble = 4
                 ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
                                                        WindowsMediaPlayer1.URL = e
                                                        Label16.Caption = "正在播放：" & Label5.Caption
                                                        simble = 5
                         ElseIf WindowsMediaPlayer1.URL = g And simble = 7 Then
                                                            WindowsMediaPlayer1.URL = f
                                                            Label16.Caption = "正在播放：" & Label6.Caption
                                                            simble = 6
                         ElseIf WindowsMediaPlayer1.URL = h And simble = 8 Then
                                                                WindowsMediaPlayer1.URL = g
                                                                Label16.Caption = "正在播放：" & Label7.Caption
                                                                simble = 7
                           ElseIf WindowsMediaPlayer1.URL = i And simble = 9 Then
                                                                WindowsMediaPlayer1.URL = h
                                                                Label16.Caption = "正在播放：" & Label8.Caption
                                                                simble = 8
                         ElseIf WindowsMediaPlayer1.URL = j And simble = 10 Then
                                                                    WindowsMediaPlayer1.URL = i
                                                                    Label16.Caption = "正在播放：" & Label9.Caption
                                                                    simble = 9
  
                        

'我的收藏——————————————————————————————————————————顺序循环播放—————————我的收藏———————

                         ElseIf WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                                    WindowsMediaPlayer1.URL = t
                                                                    Label16.Caption = "正在播放：" & Label32.Caption
                                                                    simble = 20
                        ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                                    WindowsMediaPlayer1.URL = k
                                                                    Label16.Caption = "正在播放：" & Label23.Caption
                                                                    simble = 11
                                ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                                        WindowsMediaPlayer1.URL = l
                                                                        Label16.Caption = "正在播放：" & Label24.Caption
                                                                        simble = 12
                            ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                                            WindowsMediaPlayer1.URL = ma
                                                                        Label16.Caption = "正在播放：" & Label25.Caption
                                                                        simble = 13
                             ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                                                        WindowsMediaPlayer1.URL = n
                                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                                        simble = 14
                      ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
                                                                        WindowsMediaPlayer1.URL = o
                                                                        Label16.Caption = "正在播放：" & Label27.Caption
                                                                        simble = 15
                       ElseIf WindowsMediaPlayer1.URL = q And simble = 17 Then
                                                                        WindowsMediaPlayer1.URL = p
                                                                        Label16.Caption = "正在播放：" & Label28.Caption
                                                                        simble = 16
                           ElseIf WindowsMediaPlayer1.URL = r And simble = 18 Then
                                                                            WindowsMediaPlayer1.URL = q
                                                                            Label16.Caption = "正在播放：" & Label29.Caption
                                                                            simble = 17
                        ElseIf WindowsMediaPlayer1.URL = s And simble = 19 Then
                                                                        WindowsMediaPlayer1.URL = r
                                                                        Label16.Caption = "正在播放：" & Label30.Caption
                                                                        simble = 18
                        ElseIf WindowsMediaPlayer1.URL = t And simble = 20 Then
                                                                        WindowsMediaPlayer1.URL = s
                                                                        Label16.Caption = "正在播放：" & Label31.Caption
                                                                        simble = 19
                        End If
                        
       Timer5.Enabled = True
End Sub

Private Sub Label57_Change()
On Error Resume Next
     If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                    If Label2.Caption <> "" Then
                                                                    WindowsMediaPlayer1.URL = b
                                                                    Label16.Caption = "正在播放：" & Label2.Caption
                                                                    simble = 2
        
                                                     Else
                                                                        WindowsMediaPlayer1.URL = a
                                                                    Label16.Caption = "正在播放：" & Label1.Caption
                                                                    simble = 1

                                                    End If
                      ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                 If Label3.Caption <> "" Then
                                                                       WindowsMediaPlayer1.URL = c
                                                                       Label16.Caption = "正在播放：" & Label3.Caption
                                                                       simble = 3
                                                                    
                                                  Else
                                                                            WindowsMediaPlayer1.URL = a
                                                                       Label16.Caption = "正在播放：" & Label1.Caption
                                                                       simble = 1
                                                                      
                                                 End If
                       ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                  If Label4.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = d
                                                        Label16.Caption = "正在播放：" & Label4.Caption
                                                        simble = 4
     
                                                  Else
                                                             WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                 End If
                       ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                If Label5.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = e
                                                        Label16.Caption = "正在播放：" & Label5.Caption
                                                        simble = 5
   
                                                 Else
                                                             WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                 End If
                    ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                                 If Label6.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = f
                                                        Label16.Caption = "正在播放：" & Label6.Caption
                                                        simble = 6
     
                                                Else
                                                             WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                End If
                 ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
                                                 If Label7.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = g
                                                        Label16.Caption = "正在播放：" & Label7.Caption
                                                        simble = 7
      
                                               Else
                                                            WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                End If
                         ElseIf WindowsMediaPlayer1.URL = g And simble = 7 Then
                                                 If Label8.Caption <> "" Then
                                                            WindowsMediaPlayer1.URL = h
                                                            Label16.Caption = "正在播放：" & Label8.Caption
                                                            simble = 8
   
                                                    Else
                                                                 WindowsMediaPlayer1.URL = a
                                                            Label16.Caption = "正在播放：" & Label1.Caption
                                                            simble = 1
     
                                                 End If
                         ElseIf WindowsMediaPlayer1.URL = h And simble = 8 Then
                                                 If Label9.Caption <> "" Then
                                                                WindowsMediaPlayer1.URL = i
                                                                Label16.Caption = "正在播放：" & Label9.Caption
                                                                simble = 9
    
                                                    Else
                                                                     WindowsMediaPlayer1.URL = a
                                                                Label16.Caption = "正在播放：" & Label1.Caption
                                                                simble = 1
 
          
                                                     End If
                           ElseIf WindowsMediaPlayer1.URL = i And simble = 9 Then
                                                     If Label10.Caption <> "" Then
                                                                WindowsMediaPlayer1.URL = j
                                                                Label16.Caption = "正在播放：" & Label10.Caption
                                                                simble = 10
   
                                                    Else
                                                                         WindowsMediaPlayer1.URL = a
                                                                    Label16.Caption = "正在播放：" & Label1.Caption
                                                                    simble = 1
                                                                
                                                    End If
                         ElseIf WindowsMediaPlayer1.URL = j And simble = 10 Then
                                                                    WindowsMediaPlayer1.URL = a
                                                                    Label16.Caption = "正在播放：" & Label1.Caption
                                                                    simble = 1
  
                        

'我的收藏——————————————————————————————————————————顺序循环播放—————————我的收藏———————

                         ElseIf WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                  If Label24.Caption <> "" Then
                                                                    WindowsMediaPlayer1.URL = l
                                                                    Label16.Caption = "正在播放：" & Label24.Caption
                                                                    simble = 12
                    
       
                                                 Else
                                                                    WindowsMediaPlayer1.URL = k
                                                                    Label16.Caption = "正在播放：" & Label23.Caption
                                                                    simble = 11
        
                                                   End If

    
                        ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                    If Label25.Caption <> "" Then
                                                                    WindowsMediaPlayer1.URL = ma
                                                                    Label16.Caption = "正在播放：" & Label25.Caption
                                                                    simble = 13
        
       
                                                     Else
                                                                    WindowsMediaPlayer1.URL = k
                                                                    Label16.Caption = "正在播放：" & Label23.Caption
                                                                    simble = 11
       
                                                    End If

      
                                ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                     If Label26.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = n
                                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                                        simble = 14
        
  
                                                        Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        
   
                                                        End If

      
                            ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                    If Label27.Caption <> "" Then
                                                                            WindowsMediaPlayer1.URL = o
                                                                        Label16.Caption = "正在播放：" & Label27.Caption
                                                                        simble = 15
        
  
                                                    Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        

                                                    End If

      
                             ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                                         If Label28.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = p
                                                                        Label16.Caption = "正在播放：" & Label28.Caption
                                                                        simble = 16

                                                         Else
                                                                            WindowsMediaPlayer1.URL = k
                                                                            Label16.Caption = "正在播放：" & Label23.Caption
                                                                            simble = 11
        

                                                        End If

      
                             ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
                                                          If Label29.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = q
                                                                        Label16.Caption = "正在播放：" & Label29.Caption
                                                                        simble = 17
        
 
                                                        Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        

                                                        End If

      
                       ElseIf WindowsMediaPlayer1.URL = q And simble = 17 Then
                                                         If Label30.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = r
                                                                        Label16.Caption = "正在播放：" & Label30.Caption
                                                                        simble = 18
        
    
                                                        Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11

                                                        End If

      
                           ElseIf WindowsMediaPlayer1.URL = r And simble = 18 Then
                                                    If Label31.Caption <> "" Then
                                                                            WindowsMediaPlayer1.URL = s
                                                                            Label16.Caption = "正在播放：" & Label31.Caption
                                                                            simble = 19
        
      
                                                    Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        

                                                    End If

      
                             ElseIf WindowsMediaPlayer1.URL = s And simble = 19 Then
                                                    If Label32.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = t
                                                                        Label16.Caption = "正在播放：" & Label32.Caption
                                                                        simble = 20
        
   
                                                    Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        
    
                                                    End If



                        ElseIf WindowsMediaPlayer1.URL = t And simble = 20 Then
      
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                                        simble = 11
        
 
                        End If
                        
                   Timer5.Enabled = True
End Sub

Private Sub Label58_Change()
WindowsMediaPlayer1.Controls.pause
End Sub

Private Sub Label59_Change()
WindowsMediaPlayer1.Controls.stop
End Sub

Private Sub Label6_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label6_DblClick()
If Label6.Caption <> "" Then
WindowsMediaPlayer1.URL = f
Label16.Caption = "正在播放：" & Label6.Caption
simble = 6

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label6.Caption <> "" Then
        If Button = 1 Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFF8080
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 13
           Label6.ForeColor = RGB(255, 255, 255)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then

           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFF8080
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF

           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 13
           Label6.ForeColor = RGB(255, 255, 255)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
           PopupMenu Form10.right
    End If
ElseIf Label6.Caption = "" Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
 End If
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label7_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label7_DblClick()
If Label7.Caption <> "" Then
WindowsMediaPlayer1.URL = g
Label16.Caption = "正在播放：" & Label7.Caption
simble = 7
Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label7.Caption <> "" Then
        If Button = 1 Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &HFF8080
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 13
           Label7.ForeColor = RGB(255, 255, 255)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then

           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &HFF8080
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF

           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 13
           Label7.ForeColor = RGB(255, 255, 255)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
           PopupMenu Form10.right
    End If
ElseIf Label7.Caption = "" Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
 End If
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label8_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label8_DblClick()
If Label8.Caption <> "" Then
WindowsMediaPlayer1.URL = h
Label16.Caption = "正在播放：" & Label8.Caption
simble = 8

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label8.Caption <> "" Then
        If Button = 1 Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFF8080
           Label10.BackColor = &HFFC0FF
                        
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 13
           Label8.ForeColor = RGB(255, 255, 255)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then

           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFF8080
           Label10.BackColor = &HFFC0FF

           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 13
           Label8.ForeColor = RGB(255, 255, 255)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
           PopupMenu Form10.right
    End If
ElseIf Label8.Caption = "" Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
 End If
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Label9_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Label9_DblClick()
If Label9.Caption <> "" Then
WindowsMediaPlayer1.URL = i
Label16.Caption = "正在播放：" & Label9.Caption
simble = 9

Timer1.Enabled = True
Timer2.Enabled = True
End If
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label9.Caption <> "" Then
        If Button = 1 Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &HFF8080

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 13
           Label9.ForeColor = RGB(255, 255, 255)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
    ElseIf Button = 2 Then

           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &HFF8080

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF

           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 13
           Label9.ForeColor = RGB(255, 255, 255)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
           PopupMenu Form10.right
    End If
ElseIf Label9.Caption = "" Then
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
 End If
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub mycount_Click()


  
On Error Resume Next
If Label1.BackColor = &HFF8080 Then
    If Label23.Caption = "" Then
       Label23.Caption = Label1.Caption
       k = a
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label1.Caption
       l = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label1.Caption
       ma = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label1.Caption
       n = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label1.Caption
       o = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label1.Caption
       p = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label1.Caption
       q = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label1.Caption
       r = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label1.Caption
       s = a
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label1.Caption
       t = a
    End If
    

  
ElseIf Label2.BackColor = &HFF8080 Then
   If Label23.Caption = "" Then
       Label23.Caption = Label2.Caption
       k = b
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label2.Caption
       l = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label2.Caption
       ma = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label2.Caption
       n = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label2.Caption
       o = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label2.Caption
       p = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label2.Caption
       q = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label2.Caption
       r = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label2.Caption
       s = b
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label2.Caption
       t = b
    End If
    

ElseIf Label3.BackColor = &HFF8080 Then
   If Label23.Caption = "" Then
       Label23.Caption = Label3.Caption
       k = c
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label3.Caption
       l = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label3.Caption
       ma = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label3.Caption
       n = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label3.Caption
       o = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label3.Caption
       p = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label3.Caption
       q = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label3.Caption
       r = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label3.Caption
       s = c
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label3.Caption
       t = c
    End If
ElseIf Label4.BackColor = &HFF8080 Then
   If Label23.Caption = "" Then
       Label23.Caption = Label4.Caption
       k = d
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label4.Caption
       l = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label4.Caption
       ma = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label4.Caption
       n = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label4.Caption
       o = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label4.Caption
       p = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label4.Caption
       q = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label4.Caption
       r = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label4.Caption
       s = d
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label4.Caption
       t = d
    End If
ElseIf Label5.BackColor = &HFF8080 Then
   If Label23.Caption = "" Then
       Label23.Caption = Label5.Caption
       k = e
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label5.Caption
       l = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label5.Caption
       ma = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label5.Caption
       n = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label5.Caption
       o = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label5.Caption
       p = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label5.Caption
       q = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label5.Caption
       r = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label5.Caption
       s = e
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label5.Caption
       t = e
    End If
ElseIf Label6.BackColor = &HFF8080 Then
   If Label23.Caption = "" Then
       Label23.Caption = Label6.Caption
       k = f
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label6.Caption
       l = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label6.Caption
       ma = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label6.Caption
       n = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label6.Caption
       o = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label6.Caption
       p = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label6.Caption
       q = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label6.Caption
       r = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label6.Caption
       s = f
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label6.Caption
       t = f
    End If
ElseIf Label7.BackColor = &HFF8080 Then
   If Label23.Caption = "" Then
       Label23.Caption = Label7.Caption
       k = g
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label7.Caption
       l = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label7.Caption
       ma = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label7.Caption
       n = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label7.Caption
       o = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label7.Caption
       p = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label7.Caption
       q = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label7.Caption
       r = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label7.Caption
       s = g
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label7.Caption
       t = g
    End If
ElseIf Label8.BackColor = &HFF8080 Then
   If Label23.Caption = "" Then
       Label23.Caption = Label8.Caption
       k = h
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label8.Caption
       l = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label8.Caption
       ma = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label8.Caption
       n = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label8.Caption
       o = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label8.Caption
       p = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label8.Caption
       q = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label8.Caption
       r = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label8.Caption
       s = h
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label8.Caption
       t = h
    End If
ElseIf Label9.BackColor = &HFF8080 Then
   If Label23.Caption = "" Then
       Label23.Caption = Label9.Caption
       k = i
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label9.Caption
       l = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label9.Caption
       ma = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label9.Caption
       n = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label9.Caption
       o = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label9.Caption
       p = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label9.Caption
       q = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label9.Caption
       r = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label9.Caption
       s = i
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label9.Caption
       t = i
    End If
ElseIf Label10.BackColor = &HFF8080 Then
   If Label23.Caption = "" Then
       Label23.Caption = Label10.Caption
       k = j
    ElseIf Label23.Caption <> "" And Label24.Caption = "" Then
       Label24.Caption = Label10.Caption
       l = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
       Label25.Caption = Label10.Caption
       ma = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
       Label26.Caption = Label10.Caption
       n = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
       Label27.Caption = Label10.Caption
       o = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
       Label28.Caption = Label10.Caption
       p = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
       Label29.Caption = Label10.Caption
       q = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
       Label30.Caption = Label10.Caption
       r = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
       Label31.Caption = Label10.Caption
       s = j
    ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
       Label32.Caption = Label10.Caption
       t = j
    End If
End If

  Label23.Visible = True
  Label24.Visible = True
  Label25.Visible = True
  Label26.Visible = True
  Label27.Visible = True
  Label28.Visible = True
  Label29.Visible = True
  Label30.Visible = True
  Label31.Visible = True
  Label32.Visible = True
  
  Picture2.Visible = False
  Picture13.Visible = True
  Picture3.Visible = False
  Picture15.Visible = True
  
            Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub Picture1_Click()

If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label20.ForeColor = &H4040&
Label21.ForeColor = &H808000


If 3600 <= X And X <= 4080 And 840 <= Y And Y <= 1320 Then
   Picture6.Visible = False
   Picture9.Visible = True
ElseIf 4000 <= X And X <= 4375 And 360 <= Y And Y <= 690 Then
   Picture21.Visible = True
ElseIf 1320 <= X And X <= 3990 And 375 <= Y And Y <= 675 Then
   Text1.Enabled = True
Else
   Picture6.Visible = ture
   Picture9.Visible = False
   Picture7.Visible = True
   Picture10.Visible = False
   Picture21.Visible = False
   
   

End If



End Sub

Private Sub Picture10_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
  
End Sub

Private Sub Picture11_Click()

  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

If Form2.Visible = False Then
   Form2.Visible = True
   Form2.WebBrowser1.Navigate ("http://mp3.baidu.com")
End If
End Sub

Private Sub Picture13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = True
Picture16.Visible = False
Picture4.Visible = True
Picture17.Visible = False
End Sub

Private Sub Picture14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)



If Button = 1 Then
 Picture14.Visible = False
 Picture2.Visible = True
 Picture3.Left = 1800
 Picture3.Visible = True
 Picture16.Left = 1800
 Picture15.Visible = False
 Picture4.Visible = True
 Picture18.Visible = False
 

 
   Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

  If Label22.Caption <> "" Then
  
                Picture19.Visible = False
  
                Label23.Visible = False
                Label24.Visible = False
                Label25.Visible = False
                Label26.Visible = False
                Label27.Visible = False
                Label28.Visible = False
                Label29.Visible = False
                Label30.Visible = False
                Label31.Visible = False
                Label32.Visible = False


            Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
         
     ElseIf Label22.Caption = "" Then
         Picture19.Visible = False
                Label23.Visible = False
                Label24.Visible = False
                Label25.Visible = False
                Label26.Visible = False
                Label27.Visible = False
                Label28.Visible = False
                Label29.Visible = False
                Label30.Visible = False
                Label31.Visible = False
                Label32.Visible = False


            Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
     End If
End If
End Sub

Private Sub Picture14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

If Picture15.Visible = False Then
Picture3.Visible = True
Picture16.Visible = False
End If
If Picture18.Visible = False Then
Picture4.Visible = True
Picture17.Visible = False
End If
End Sub


Private Sub Picture15_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Picture15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Picture13.Visible = True
Picture14.Visible = False
Picture4.Visible = True
Picture17.Visible = False
End Sub

Private Sub Picture16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

If Button = 1 Then
  Picture16.Visible = False
  Picture15.Visible = True
  Picture13.Visible = True
  Picture2.Visible = False
  Picture4.Visible = True
  Picture18.Visible = False
  

  
    Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If

If Label22.Caption <> "" Then

            Picture19.Visible = False

            Label23.Visible = True
            Label24.Visible = True
            Label25.Visible = True
            Label26.Visible = True
            Label27.Visible = True
            Label28.Visible = True
            Label29.Visible = True
            Label30.Visible = True
            Label31.Visible = True
            Label32.Visible = True

           
             Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
         
ElseIf Label22.Caption = "" Then
        Picture19.Visible = True
End If
End If
End Sub

Private Sub Picture16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Picture2.Visible = False Then
Picture13.Visible = True
Picture14.Visible = False
Else
Picture4.Visible = True
Picture17.Visible = False
End If
End Sub

Private Sub Picture17_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Picture17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
  
If Button = 1 Then
  'Picture17.Visible = False
 ' Picture18.Visible = True
  'Picture2.Visible = False
  'Picture13.Visible = True
 ' Picture3.Left = 1295
  'Picture3.Visible = True
 ' Picture16.Left = 1295
 ' Picture15.Visible = False
 
   Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
  
  MsgBox "水平有限，暂不可用", , "草哥提示"
End If
End Sub

Private Sub Picture17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Picture2.Visible = False Then
Picture13.Visible = True
Picture14.Visible = False
End If

If Picture15.Visible = False Then
 Picture3.Visible = True
 Picture16.Visible = False
Else
 Picture3.Visible = False
 Picture16.Visible = False
End If
End Sub

Private Sub Picture18_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Picture18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Picture13.Visible = True
Picture14.Visible = False
Picture3.Visible = True
Picture16.Visible = False
End Sub

Private Sub Picture19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture13.Visible = True
Picture14.Visible = False
Picture4.Visible = True
Picture17.Visible = False

MousePointer = 1
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Picture2_Click()
  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1

Picture3.Visible = True
Picture16.Visible = False
Picture4.Visible = True
Picture17.Visible = False


End Sub

Private Sub Picture20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If 0 < X And X < 360 And 15 < Y And Y < 255 Then
   Picture22.Visible = True
   Picture24.Visible = False
   Picture26.Visible = False
   Picture28.Visible = False
   

ElseIf 360 < X And X < 720 And 15 < Y And Y < 255 Then
   Picture22.Visible = False
   Picture24.Visible = True
   Picture26.Visible = False
   Picture28.Visible = False
   

ElseIf 720 < X And X < 1080 And 15 < Y And Y < 255 Then
   Picture22.Visible = False
   Picture24.Visible = False
   Picture26.Visible = True
   Picture28.Visible = False
   

ElseIf 1080 < X And X < 1560 And 15 < Y And Y < 255 Then
   Picture22.Visible = False
   Picture24.Visible = False
   Picture26.Visible = False
   Picture28.Visible = True
   

End If
End Sub

Private Sub Picture21_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
   urlstring = "http://mp3.baidu.com/m?f=ms&tn=baidump3&ct=134217728&lf=&rn=&word=" & Text1.Text & "&lm=0"
   Form2.Visible = True
   Form2.WebBrowser1.Navigate (urlstring)
   
  Text1.Enabled = False
  If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
  End If

End If
End Sub

Private Sub Picture21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label20.ForeColor = &H4040&
Label21.ForeColor = &H808000

Picture22.Visible = False
Picture24.Visible = False
Picture26.Visible = False
Picture28.Visible = False
End Sub

Private Sub Picture22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
        Picture22.Visible = False
        
End If
End Sub

Private Sub Picture22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu Form10.youshangjiao, 0, 2785, 270
End Sub

Private Sub Picture23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0

Label53.BorderStyle = 0
Label54.BorderStyle = 0
End Sub

Private Sub Picture24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  Picture24.Visible = False
End If
End Sub

Private Sub Picture24_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Form1.WindowState = 1

End Sub

Private Sub Picture26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  Picture26.Visible = False
End If
End Sub

Private Sub Picture26_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form11.Show
Form1.Hide
End Sub

Private Sub Picture28_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  Picture28.Visible = False
End If
End Sub

Private Sub Picture28_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
unload Form11
unload Form2
unload Form3
unload Form4
unload Form5

unload Form7
unload Form8
unload Form9
unload Form10
unload Form1
End
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label49.BorderStyle = 0
Label50.BorderStyle = 0
If Image2.Visible = False Then
   Label51.BorderStyle = 0
End If
Label52.BorderStyle = 0
End Sub

Private Sub Picture9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Picture9.Visible = False
   Picture12.Visible = True

End If
End Sub

Private Sub Picture9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture12.Visible = False
Picture9.Visible = True

If Button = 1 Then

CommonDialog1.CancelError = True
On Error GoTo errline
CommonDialog1.ShowOpen
'CommonDialog1.Filter = "音乐文件 （*.mp3;*.wma）|*.mp3;*.wma"
WindowsMediaPlayer1.URL = CommonDialog1.FileName
Label16.Caption = "正在播放：" & CommonDialog1.FileTitle

If Label1.Caption = "" Then
   Label1.Caption = " " & CommonDialog1.FileTitle
   a = CommonDialog1.FileName
   simble = 1
           
ElseIf Label1.Caption <> "" And Label2.Caption = "" Then
   Label2.Caption = " " & CommonDialog1.FileTitle
   b = CommonDialog1.FileName
   simble = 2
           
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption = "" Then
   Label3.Caption = " " & CommonDialog1.FileTitle
   c = CommonDialog1.FileName
   simble = 3
    
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption = "" Then
   Label4.Caption = " " & CommonDialog1.FileTitle
   d = CommonDialog1.FileName
   simble = 4
   
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption = "" Then
   Label5.Caption = " " & CommonDialog1.FileTitle
   e = CommonDialog1.FileName
   simble = 5
   
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption = "" Then
   Label6.Caption = " " & CommonDialog1.FileTitle
   f = CommonDialog1.FileName
   simble = 6
   
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption = "" Then
   Label7.Caption = " " & CommonDialog1.FileTitle
   g = CommonDialog1.FileName
   simble = 7
   
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption = "" Then
   Label8.Caption = " " & CommonDialog1.FileTitle
   h = CommonDialog1.FileName
   simble = 8
   
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption <> "" And Label9.Caption = "" Then
   Label9.Caption = " " & CommonDialog1.FileTitle
   i = CommonDialog1.FileName
   simble = 9
   
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption <> "" And Label9.Caption <> "" And Label10.Caption = "" Then
   Label10.Caption = " " & CommonDialog1.FileTitle
   j = CommonDialog1.FileName
   simble = 10
   
ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption <> "" And Label9.Caption <> "" And Label10.Caption <> "" Then
           Label10.Caption = " " & CommonDialog1.FileTitle
         j = CommonDialog1.FileName
         simble = 10
          
End If
 Picture9.Visible = False
 Picture6.Visible = True
 Timer1.Enabled = True
 Timer5.Enabled = True
 
 
   Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
   
   
   Label23.Visible = False
   Label24.Visible = False
   Label25.Visible = False
   Label26.Visible = False
   Label27.Visible = False
   Label28.Visible = False
   Label29.Visible = False
   Label30.Visible = False
   Label31.Visible = False
   Label32.Visible = False
   
   Picture13.Visible = False
   Picture2.Visible = True
   Picture15.Visible = False
   Picture3.Visible = True
Exit Sub

errline:
Picture9.Visible = False
 Picture6.Visible = True
 
   Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
Exit Sub

End If
End Sub

Private Sub play_Click()
If Label1.BackColor = &HFF8080 Then
   WindowsMediaPlayer1.URL = a
   Label16.Caption = "正在播放：" & Label1.Caption
   simble = 1
   Timer1.Enabled = True
Timer2.Enabled = True
ElseIf Label2.BackColor = &HFF8080 Then
   WindowsMediaPlayer1.URL = b
   Label16.Caption = "正在播放：" & Label2.Caption
   simble = 2
   Timer1.Enabled = True
Timer2.Enabled = True
ElseIf Label3.BackColor = &HFF8080 Then
   WindowsMediaPlayer1.URL = c
   Label16.Caption = "正在播放：" & Label3.Caption
   simble = 3
   Timer1.Enabled = True
Timer2.Enabled = True
ElseIf Label4.BackColor = &HFF8080 Then
   WindowsMediaPlayer1.URL = d
   Label16.Caption = "正在播放：" & Label4.Caption
   simble = 4
   Timer1.Enabled = True
Timer2.Enabled = True
ElseIf Label5.BackColor = &HFF8080 Then
   WindowsMediaPlayer1.URL = e
   Label16.Caption = "正在播放：" & Label5.Caption
   simble = 5
   Timer1.Enabled = True
Timer2.Enabled = True
ElseIf Label6.BackColor = &HFF8080 Then
   WindowsMediaPlayer1.URL = f
   Label16.Caption = "正在播放：" & Label6.Caption
   simble = 6
   Timer1.Enabled = True
Timer2.Enabled = True
ElseIf Label7.BackColor = &HFF8080 Then
   WindowsMediaPlayer1.URL = g
   Label16.Caption = "正在播放：" & Label7.Caption
   simble = 7
   Timer1.Enabled = True
Timer2.Enabled = True
ElseIf Label8.BackColor = &HFF8080 Then
   WindowsMediaPlayer1.URL = h
   Label16.Caption = "正在播放：" & Label8.Caption
   simble = 8
   Timer1.Enabled = True
Timer2.Enabled = True
ElseIf Label9.BackColor = &HFF8080 Then
   WindowsMediaPlayer1.URL = i
   Label16.Caption = "正在播放：" & Label9.Caption
   simble = 9
   Timer1.Enabled = True
Timer2.Enabled = True
ElseIf Label10.BackColor = &HFF8080 Then
   WindowsMediaPlayer1.URL = j
   Label16.Caption = "正在播放：" & Label10.Caption
   simble = 10
   Timer1.Enabled = True
Timer2.Enabled = True
End If
'我的收藏
If Label23.BackColor = vbRed Then
   WindowsMediaPlayer1.URL = k
   Label16.Caption = "正在播放：" & Label23.Caption
   simble = 11
   Timer1.Enabled = True
ElseIf Label24.BackColor = vbRed Then
   WindowsMediaPlayer1.URL = l
   Label16.Caption = "正在播放：" & Label24.Caption
   simble = 12
   Timer1.Enabled = True
ElseIf Label25.BackColor = vbRed Then
   WindowsMediaPlayer1.URL = ma
   Label16.Caption = "正在播放：" & Label25.Caption
   simble = 13
   Timer1.Enabled = True
ElseIf Label26.BackColor = vbRed Then
   WindowsMediaPlayer1.URL = n
   Label16.Caption = "正在播放：" & Label26.Caption
   simble = 14
   Timer1.Enabled = True
ElseIf Label27.BackColor = vbRed Then
   WindowsMediaPlayer1.URL = l
   Label16.Caption = "正在播放：" & Label27.Caption
   simble = 15
   Timer1.Enabled = True
ElseIf Label28.BackColor = vbRed Then
   WindowsMediaPlayer1.URL = l
   Label16.Caption = "正在播放：" & Label28.Caption
   simble = 16
   Timer1.Enabled = True
ElseIf Label29.BackColor = vbRed Then
   WindowsMediaPlayer1.URL = l
   Label16.Caption = "正在播放：" & Label29.Caption
   simble = 17
   Timer1.Enabled = True
ElseIf Label30.BackColor = vbRed Then
   WindowsMediaPlayer1.URL = l
   Label16.Caption = "正在播放：" & Label30.Caption
   simble = 18
   Timer1.Enabled = True
ElseIf Label31.BackColor = vbRed Then
   WindowsMediaPlayer1.URL = l
   Label16.Caption = "正在播放：" & Label31.Caption
   simble = 19
   Timer1.Enabled = True
ElseIf Label32.BackColor = vbRed Then
   WindowsMediaPlayer1.URL = l
   Label16.Caption = "正在播放：" & Label32.Caption
   simble = 20
   Timer1.Enabled = True
End If
End Sub


Private Sub rename_Click()

            If Form1.Label1.BackColor = &HFF8080 Then
                Form1.Label35.Caption = a
                Form1.Label37.Caption = 1
            ElseIf Form1.Label2.BackColor = &HFF8080 Then
                Form1.Label35.Caption = b
                Form1.Label37.Caption = 2
            ElseIf Form1.Label3.BackColor = &HFF8080 Then
               Form1.Label35.Caption = c
               Form1.Label37.Caption = 3
            ElseIf Form1.Label4.BackColor = &HFF8080 Then
               Form1.Label35.Caption = d
               Form1.Label37.Caption = 4
            ElseIf Form1.Label5.BackColor = &HFF8080 Then
                Form1.Label35.Caption = e
                Form1.Label37.Caption = 5
            ElseIf Form1.Label6.BackColor = &HFF8080 Then
                Form1.Label35.Caption = f
                Form1.Label37.Caption = 6
            ElseIf Form1.Label7.BackColor = &HFF8080 Then
                Form1.Label35.Caption = g
                Form1.Label37.Caption = 7
            ElseIf Form1.Label8.BackColor = &HFF8080 Then
                Form1.Label35.Caption = h
                Form1.Label37.Caption = 8
            ElseIf Form1.Label9.BackColor = &HFF8080 Then
                Form1.Label35.Caption = i
                Form1.Label37.Caption = 9
            ElseIf Form1.Label10.BackColor = &HFF8080 Then
                Form1.Label35.Caption = j
                Form1.Label37.Caption = 10
           '我的收藏————————————————————————————————————————————
            ElseIf Form1.Label23.BackColor = vbRed Then
               Form1.Label35.Caption = k
               Form1.Label37.Caption = 11
            ElseIf Form1.Label24.BackColor = vbRed Then
               Form1.Label35.Caption = l
               Form1.Label37.Caption = 12
            ElseIf Form1.Label25.BackColor = vbRed Then
               Form1.Label35.Caption = ma
               Form1.Label37.Caption = 13
            ElseIf Form1.Label26.BackColor = vbRed Then
               Form1.Label35.Caption = n
               Form1.Label37.Caption = 14
            ElseIf Form1.Label27.BackColor = vbRed Then
               Form1.Label35.Caption = o
               Form1.Label37.Caption = 15
            ElseIf Form1.Label28.BackColor = vbRed Then
               Form1.Label35.Caption = p
               Form1.Label37.Caption = 16
            ElseIf Form1.Label29.BackColor = vbRed Then
               Form1.Label35.Caption = q
               Form1.Label37.Caption = 17
            ElseIf Form1.Label30.BackColor = vbRed Then
              Form1.Label35.Caption = r
              Form1.Label37.Caption = 18
            ElseIf Form1.Label31.BackColor = vbRed Then
               Form1.Label35.Caption = s
               Form1.Label37.Caption = 19
            ElseIf Form1.Label32.BackColor = vbRed Then
               Form1.Label35.Caption = t
               Form1.Label37.Caption = 20
            End If

Form5.Show 1
End Sub

Private Sub save_Click()
If Label22.Caption <> "" Then
   Open "c:\百度音乐播放器\" & Label22.Caption & "播放列表.txt" For Output As #2
   Print #2, Label1.Caption
   Print #2, a
   Print #2, Label2.Caption
   Print #2, b
   Print #2, Label3.Caption
   Print #2, c
   Print #2, Label4.Caption
   Print #2, d
   Print #2, Label5.Caption
   Print #2, e
   Print #2, Label6.Caption
   Print #2, f
   Print #2, Label7.Caption
   Print #2, g
   Print #2, Label8.Caption
   Print #2, h
   Print #2, Label9.Caption
   Print #2, i
   Print #2, Label10.Caption
   Print #2, j

   Print #2, pic

   Close #2
Else
  MsgBox "未登录，不可保存"
End If

   
End Sub

Private Sub shuiji_Click()
   Form3.Option3.Value = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = True
   Form10.danqu.Checked = False
End Sub

Private Sub shunxu_Click()
   Form3.Option1.Value = True

   
   Form10.shunxu.Checked = True
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
End Sub

Private Sub shunxuxunhuan_Click()
   Form3.Option2.Value = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = True
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
End Sub

Private Sub system_Click()
Form3.Show 1
End Sub

Private Sub Text1_Click()
If Text1.Text = "搜索 歌曲、歌手、专辑" Then
   Text1.Text = ""
   Text1.ForeColor = black
End If
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 3
Picture21.Visible = False

Picture22.Visible = False
Picture24.Visible = False
Picture26.Visible = False
Picture28.Visible = False
End Sub

Private Sub Text2_Click()
If Text2.Text = "搜索 歌曲、歌手、专辑" Then
   Text2.Text = ""
   Text2.ForeColor = black
End If
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 3
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label12.Caption = WindowsMediaPlayer1.Controls.currentPositionString
Label13.Caption = WindowsMediaPlayer1.currentMedia.durationString



If Label12.Caption = "" Then
   Label12.Caption = "00:00"
   Label12.Visible = False
   Label38.Visible = True
ElseIf Label12.Caption <> "00:00" Then
   Label12.Visible = True
   Label38.Visible = False
End If

If Picture2.Visible = True And WindowsMediaPlayer1.playState = 1 Then
    If Label1.Caption = "" And Label2.Caption = "" And Label3.Caption = "" And Label4.Caption = "" And Label5.Caption = "" And Label6.Caption = "" And Label7.Caption = "" And Label8.Caption = "" And Label9.Caption = "" And Label10.Caption = "" Then
       Label16.Caption = "百度音乐_音乐你的生活"
       Label16.Alignment = 2
       WindowsMediaPlayer1.Controls.stop
       WindowsMediaPlayer1.Enabled = False
       Label13.Caption = "00:00"
       Timer3.Enabled = False
    Else
       WindowsMediaPlayer1.Enabled = True
       Timer3.Enabled = True
    End If
ElseIf Picture15.Visible = True And Picture19.Visible = False And WindowsMediaPlayer1.playState = 1 Then
    If Label23.Caption = "" And Label24.Caption = "" And Label25.Caption = "" And Label26.Caption = "" And Label27.Caption = "" And Label28.Caption = "" And Label29.Caption = "" And Label30.Caption = "" And Label31.Caption = "" And Label32.Caption = "" Then
       Label16.Caption = "百度音乐_音乐你的生活"
       Label16.Alignment = 2
       WindowsMediaPlayer1.Controls.stop
        WindowsMediaPlayer1.Enabled = False
        Label13.Caption = "00:00"
        Timer3.Enabled = False
    Else
       WindowsMediaPlayer1.Enabled = True
       Timer3.Enabled = True
    End If
ElseIf Label1.Caption = "" And Label2.Caption = "" And Label3.Caption = "" And Label4.Caption = "" And Label5.Caption = "" And Label6.Caption = "" And Label7.Caption = "" And Label8.Caption = "" And Label9.Caption = "" And Label10.Caption = "" And Label23.Caption = "" And Label24.Caption = "" And Label25.Caption = "" And Label26.Caption = "" And Label27.Caption = "" And Label28.Caption = "" And Label29.Caption = "" And Label30.Caption = "" And Label31.Caption = "" And Label32.Caption = "" Then
       Label16.Caption = "百度音乐_音乐你的生活"
       Label16.Alignment = 2
       WindowsMediaPlayer1.Controls.stop
       WindowsMediaPlayer1.Enabled = False
       Label13.Caption = "00:00"
End If

If Picture2.Visible = True Then
   If Label1.Caption <> "" Then
      WindowsMediaPlayer1.Enabled = True
   End If
ElseIf Picture15.Visible = True And Picture19.Visible = False Then
     If Label23.Caption <> "" Then
      WindowsMediaPlayer1.Enabled = True
   End If
End If

  
   
End Sub

Private Sub Timer2_Timer()
Dim playtime, alltime, playhour, playmunite, allhour, allmunite, playmatch, allmatch As String
allhour = Hour(Label13.Caption)
allmunite = Minute(Label13.Caption)
alltime = allhour * 60 + allmunite
playhour = Hour(Label12.Caption)
playmunite = Minute(Label12.Caption)
playtime = playhour * 60 + playmunite



playmatch = Val(playtime) + 1
allmatch = Val(alltime)
Label18.Caption = playmatch
Label19.Caption = allmatch

'顺序播放——————————————————————————————————————————————————————————-——
If Label18.Caption = Label19.Caption Then
        If Form3.Option1.Value = True Then
               
                    If WindowsMediaPlayer1.URL = a And simble = 1 Then
                             If Label2.Caption <> "" Then
                                  WindowsMediaPlayer1.URL = b
                                    Label16.Caption = "正在播放：" & Label2.Caption
                                      simble = 2
                                End If
                   ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                               If Label3.Caption <> "" Then
                                        WindowsMediaPlayer1.URL = c
                                           Label16.Caption = "正在播放：" & Label3.Caption
                                          simble = 3
                               End If
                  ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                If Label4.Caption <> "" Then
                                        WindowsMediaPlayer1.URL = d
                                         Label16.Caption = "正在播放：" & Label4.Caption
                                         simble = 4
                                 End If
                   ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                   If Label5.Caption <> "" Then
                                          WindowsMediaPlayer1.URL = e
                                            Label16.Caption = "正在播放：" & Label5.Caption
                                            simble = 5
                                   End If
                   ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                    If Label6.Caption <> "" Then
                                            WindowsMediaPlayer1.URL = f
                                            Label16.Caption = "正在播放：" & Label6.Caption
                                            simble = 6
                                     End If
                   ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
                                    If Label7.Caption <> "" Then
                                            WindowsMediaPlayer1.URL = g
                                             Label16.Caption = "正在播放：" & Label7.Caption
                                           simble = 7
                                     End If
                    ElseIf WindowsMediaPlayer1.URL = g And simble = 7 Then
                                    If Label8.Caption <> "" Then
                                            WindowsMediaPlayer1.URL = h
                                            Label16.Caption = "正在播放：" & Label8.Caption
                                            simble = 8
                                    End If
                   ElseIf WindowsMediaPlayer1.URL = h And simble = 8 Then
                                    If Label9.Caption <> "" Then
                                              WindowsMediaPlayer1.URL = i
                                             Label16.Caption = "正在播放：" & Label9.Caption
                                            simble = 9
                                     End If
                   ElseIf WindowsMediaPlayer1.URL = i And simble = 9 Then
                                   If Label10.Caption <> "" Then
                                              WindowsMediaPlayer1.URL = j
                                              Label16.Caption = "正在播放：" & Label10.Caption
                                              simble = 10
                                     End If
'我的收藏——————————————————————————————————————————顺序播放—————————我的收藏———————

                  ElseIf WindowsMediaPlayer1.URL = k And simble = 11 Then
                                 If Label24.Caption <> "" Then
                                         WindowsMediaPlayer1.URL = l
                                            Label16.Caption = "正在播放：" & Label24.Caption
                                            simble = 12
                                 End If
                  ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                    If Label25.Caption <> "" Then
                                            WindowsMediaPlayer1.URL = ma
                                            Label16.Caption = "正在播放：" & Label25.Caption
                                            simble = 13
                                    End If
                  ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                     If Label26.Caption <> "" Then
                                            WindowsMediaPlayer1.URL = n
                                            Label16.Caption = "正在播放：" & Label26.Caption
                                            simble = 14
                                     End If
                    ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                   If Label27.Caption <> "" Then
                                                WindowsMediaPlayer1.URL = o
                                                Label16.Caption = "正在播放：" & Label27.Caption
                                                simble = 15
                                    End If
                    ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                       If Label28.Caption <> "" Then
                                                    WindowsMediaPlayer1.URL = p
                                                    Label16.Caption = "正在播放：" & Label28.Caption
                                                    simble = 16
                                     End If
                    ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
                                          If Label29.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = q
                                                        Label16.Caption = "正在播放：" & Label29.Caption
                                                        simble = 17
                                        End If
                        ElseIf WindowsMediaPlayer1.URL = q And simble = 17 Then
                                         If Label30.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = r
                                                        Label16.Caption = "正在播放：" & Label30.Caption
                                                        simble = 18
                                            End If
                        ElseIf WindowsMediaPlayer1.URL = r And simble = 18 Then
                                             If Label31.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = s
                                                        Label16.Caption = "正在播放：" & Label31.Caption
                                                        simble = 19
                                                End If
                        ElseIf WindowsMediaPlayer1.URL = s And simble = 19 Then
                                                 If Label32.Caption <> "" Then
                                                            WindowsMediaPlayer1.URL = t
                                                            Label16.Caption = "正在播放：" & Label32.Caption
                                                            simble = 20
                                                End If
                        End If
           End If
  End If
'顺序循环播放——————————————————————————————————————————————————
If Label18.Caption = Label19.Caption Then
           If Form3.Option2.Value = True Then
               Timer5.Enabled = True
                        If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                    If Label2.Caption <> "" Then
                                                                    WindowsMediaPlayer1.URL = b
                                                                    Label16.Caption = "正在播放：" & Label2.Caption
                                                                    simble = 2
        
                                                     Else
                                                                        WindowsMediaPlayer1.URL = a
                                                                    Label16.Caption = "正在播放：" & Label1.Caption
                                                                    simble = 1

                                                    End If
                      ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                 If Label3.Caption <> "" Then
                                                                       WindowsMediaPlayer1.URL = c
                                                                       Label16.Caption = "正在播放：" & Label3.Caption
                                                                       simble = 3
                                                                    
                                                  Else
                                                                            WindowsMediaPlayer1.URL = a
                                                                       Label16.Caption = "正在播放：" & Label1.Caption
                                                                       simble = 1
                                                                      
                                                 End If
                       ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                  If Label4.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = d
                                                        Label16.Caption = "正在播放：" & Label4.Caption
                                                        simble = 4
     
                                                  Else
                                                             WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                 End If
                       ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                If Label5.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = e
                                                        Label16.Caption = "正在播放：" & Label5.Caption
                                                        simble = 5
   
                                                 Else
                                                             WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                 End If
                    ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                                 If Label6.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = f
                                                        Label16.Caption = "正在播放：" & Label6.Caption
                                                        simble = 6
     
                                                Else
                                                             WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                End If
                 ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
                                                 If Label7.Caption <> "" Then
                                                        WindowsMediaPlayer1.URL = g
                                                        Label16.Caption = "正在播放：" & Label7.Caption
                                                        simble = 7
      
                                               Else
                                                            WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
     
                                                End If
                         ElseIf WindowsMediaPlayer1.URL = g And simble = 7 Then
                                                 If Label8.Caption <> "" Then
                                                            WindowsMediaPlayer1.URL = h
                                                            Label16.Caption = "正在播放：" & Label8.Caption
                                                            simble = 8
   
                                                    Else
                                                                 WindowsMediaPlayer1.URL = a
                                                            Label16.Caption = "正在播放：" & Label1.Caption
                                                            simble = 1
     
                                                 End If
                         ElseIf WindowsMediaPlayer1.URL = h And simble = 8 Then
                                                 If Label9.Caption <> "" Then
                                                                WindowsMediaPlayer1.URL = i
                                                                Label16.Caption = "正在播放：" & Label9.Caption
                                                                simble = 9
    
                                                    Else
                                                                     WindowsMediaPlayer1.URL = a
                                                                Label16.Caption = "正在播放：" & Label1.Caption
                                                                simble = 1
 
          
                                                     End If
                           ElseIf WindowsMediaPlayer1.URL = i And simble = 9 Then
                                                     If Label10.Caption <> "" Then
                                                                WindowsMediaPlayer1.URL = j
                                                                Label16.Caption = "正在播放：" & Label10.Caption
                                                                simble = 10
   
                                                    Else
                                                                         WindowsMediaPlayer1.URL = a
                                                                    Label16.Caption = "正在播放：" & Label1.Caption
                                                                    simble = 1
                                                                
                                                    End If
                         ElseIf WindowsMediaPlayer1.URL = j And simble = 10 Then
                                                                    WindowsMediaPlayer1.URL = a
                                                                    Label16.Caption = "正在播放：" & Label1.Caption
                                                                    simble = 1
  
                        

'我的收藏——————————————————————————————————————————顺序循环播放—————————我的收藏———————

                         ElseIf WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                  If Label24.Caption <> "" Then
                                                                    WindowsMediaPlayer1.URL = l
                                                                    Label16.Caption = "正在播放：" & Label24.Caption
                                                                    simble = 12
                    
       
                                                 Else
                                                                    WindowsMediaPlayer1.URL = k
                                                                    Label16.Caption = "正在播放：" & Label23.Caption
                                                                    simble = 11
        
                                                   End If

    
                        ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                    If Label25.Caption <> "" Then
                                                                    WindowsMediaPlayer1.URL = ma
                                                                    Label16.Caption = "正在播放：" & Label25.Caption
                                                                    simble = 13
        
       
                                                     Else
                                                                    WindowsMediaPlayer1.URL = k
                                                                    Label16.Caption = "正在播放：" & Label23.Caption
                                                                    simble = 11
       
                                                    End If

      
                                ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                     If Label26.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = n
                                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                                        simble = 14
        
  
                                                        Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        
   
                                                        End If

      
                            ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                    If Label27.Caption <> "" Then
                                                                            WindowsMediaPlayer1.URL = o
                                                                        Label16.Caption = "正在播放：" & Label27.Caption
                                                                        simble = 15
        
  
                                                    Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        

                                                    End If

      
                             ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                                         If Label28.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = p
                                                                        Label16.Caption = "正在播放：" & Label28.Caption
                                                                        simble = 16

                                                         Else
                                                                            WindowsMediaPlayer1.URL = k
                                                                            Label16.Caption = "正在播放：" & Label23.Caption
                                                                            simble = 11
        

                                                        End If

      
                             ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
                                                          If Label29.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = q
                                                                        Label16.Caption = "正在播放：" & Label29.Caption
                                                                        simble = 17
        
 
                                                        Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        

                                                        End If

      
                       ElseIf WindowsMediaPlayer1.URL = q And simble = 17 Then
                                                         If Label30.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = r
                                                                        Label16.Caption = "正在播放：" & Label30.Caption
                                                                        simble = 18
        
    
                                                        Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11

                                                        End If

      
                           ElseIf WindowsMediaPlayer1.URL = r And simble = 18 Then
                                                    If Label31.Caption <> "" Then
                                                                            WindowsMediaPlayer1.URL = s
                                                                            Label16.Caption = "正在播放：" & Label31.Caption
                                                                            simble = 19
        
      
                                                    Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        

                                                    End If

      
                             ElseIf WindowsMediaPlayer1.URL = s And simble = 19 Then
                                                    If Label32.Caption <> "" Then
                                                                        WindowsMediaPlayer1.URL = t
                                                                        Label16.Caption = "正在播放：" & Label32.Caption
                                                                        simble = 20
        
   
                                                    Else
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                                        simble = 11
        
    
                                                    End If



                        ElseIf WindowsMediaPlayer1.URL = t And simble = 20 Then
      
                                                                        WindowsMediaPlayer1.URL = k
                                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                                        simble = 11
        
 
                        End If
              End If
  End If
'随机播放————————————————————————————————————————————————————————————
 If Label18.Caption = Label19.Caption Then
    If Form3.Option3.Value = True Then
               If WindowsMediaPlayer1.URL = a Or WindowsMediaPlayer1.URL = b Or WindowsMediaPlayer1.URL = c Or WindowsMediaPlayer1.URL = d Or WindowsMediaPlayer1.URL = e Or WindowsMediaPlayer1.URL = f Or WindowsMediaPlayer1.URL = g Or WindowsMediaPlayer1.URL = h Or WindowsMediaPlayer1.URL = i Or WindowsMediaPlayer1.URL = j Then
                        If Label1.Caption <> "" And Label2.Caption = "" Then
                                     If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                        WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
                                     End If
                        ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption = "" Then
                                     If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                        WindowsMediaPlayer1.URL = b
                                                        Label16.Caption = "正在播放：" & Label2.Caption
                                                        simble = 2
                                    ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                        WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
                                    End If
                         ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption = "" Then
                                    If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                        WindowsMediaPlayer1.URL = c
                                                        Label16.Caption = "正在播放：" & Label3.Caption
                                                        simble = 3
                                     ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                        WindowsMediaPlayer1.URL = b
                                                        Label16.Caption = "正在播放：" & Label2.Caption
                                                        simble = 2
                                     ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                        WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
                                    End If
                      ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption = "" Then
                                    If WindowsMediaPlayer1.URL = a And simble = "1" Then
                                                            WindowsMediaPlayer1.URL = d
                                                            Label16.Caption = "正在播放：" & Label4.Caption
                                                            simble = 4
                                    ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                        WindowsMediaPlayer1.URL = c
                                                        Label16.Caption = "正在播放：" & Label3.Caption
                                                        simble = 3
                                    ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                        WindowsMediaPlayer1.URL = b
                                                        Label16.Caption = "正在播放：" & Label2.Caption
                                                        simble = 2
                                    ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                        WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
                                    End If
          
                     ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption = "" Then
                                 If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                        WindowsMediaPlayer1.URL = d
                                                        Label16.Caption = "正在播放：" & Label4.Caption
                                                        simble = 4
                                ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                        WindowsMediaPlayer1.URL = c
                                                        Label16.Caption = "正在播放：" & Label3.Caption
                                                        simble = 3
                                 ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                        WindowsMediaPlayer1.URL = e
                                                        Label16.Caption = "正在播放：" & Label5.Caption
                                                        simble = 5
                                 ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                                        WindowsMediaPlayer1.URL = b
                                                        Label16.Caption = "正在播放：" & Label2.Caption
                                                        simble = 2
                             ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                        WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
                                 End If
             ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption = "" Then
                              If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                        WindowsMediaPlayer1.URL = d
                                                        Label16.Caption = "正在播放：" & Label4.Caption
                                                        simble = 4
                                 ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                        WindowsMediaPlayer1.URL = c
                                                        Label16.Caption = "正在播放：" & Label3.Caption
                                                        simble = 3
                                ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                        WindowsMediaPlayer1.URL = e
                                                        Label16.Caption = "正在播放：" & Label5.Caption
                                                        simble = 5
                                ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                                        WindowsMediaPlayer1.URL = b
                                                        Label16.Caption = "正在播放：" & Label2.Caption
                                                        simble = 2
                                ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                        WindowsMediaPlayer1.URL = f
                                                        Label16.Caption = "正在播放：" & Label6.Caption
                                                        simble = 6
                                ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
                                                        WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
                                End If
             ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption = "" Then
                                 If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                        WindowsMediaPlayer1.URL = d
                                                        Label16.Caption = "正在播放：" & Label4.Caption
                                                        simble = 4
                                 ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                        WindowsMediaPlayer1.URL = g
                                                        Label16.Caption = "正在播放：" & Label7.Caption
                                                        simble = 7
                                ElseIf WindowsMediaPlayer1.URL = g And simble = 7 Then
                                                        WindowsMediaPlayer1.URL = c
                                                        Label16.Caption = "正在播放：" & Label3.Caption
                                                        simble = 3
                                 ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                        WindowsMediaPlayer1.URL = e
                                                        Label16.Caption = "正在播放：" & Label5.Caption
                                                        simble = 5
                                ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                                        WindowsMediaPlayer1.URL = b
                                                        Label16.Caption = "正在播放：" & Label2.Caption
                                                        simble = 2
                                 ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                        WindowsMediaPlayer1.URL = f
                                                        Label16.Caption = "正在播放：" & Label6.Caption
                                                        simble = 6
                                ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
                                                        WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
                               End If
                 ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption <> "" And Label9.Caption = "" Then
                                If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                        WindowsMediaPlayer1.URL = d
                                                        Label16.Caption = "正在播放：" & Label4.Caption
                                                        simble = 4
                                ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                        WindowsMediaPlayer1.URL = g
                                                        Label16.Caption = "正在播放：" & Label7.Caption
                                                        simble = 7
                                 ElseIf WindowsMediaPlayer1.URL = g And simble = 7 Then
                                                        WindowsMediaPlayer1.URL = c
                                                        Label16.Caption = "正在播放：" & Label3.Caption
                                                        simble = 3
                                 ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                        WindowsMediaPlayer1.URL = e
                                                        Label16.Caption = "正在播放：" & Label5.Caption
                                                        simble = 5
                                 ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                                        WindowsMediaPlayer1.URL = b
                                                        Label16.Caption = "正在播放：" & Label2.Caption
                                                        simble = 2
                                ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                        WindowsMediaPlayer1.URL = f
                                                        Label16.Caption = "正在播放：" & Label6.Caption
                                                        simble = 6
                                 ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
                                                        WindowsMediaPlayer1.URL = h
                                                        Label16.Caption = "正在播放：" & Label8.Caption
                                                        simble = 8
                                ElseIf WindowsMediaPlayer1.URL = h And simble = 8 Then
                                                        WindowsMediaPlayer1.URL = a
                                                        Label16.Caption = "正在播放：" & Label1.Caption
                                                        simble = 1
                                 End If
                     ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption <> "" And Label9.Caption <> "" And Label10.Caption = "" Then
                                 If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                        WindowsMediaPlayer1.URL = d
                                                        Label16.Caption = "正在播放：" & Label4.Caption
                                                        simble = 4
                                 ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                        WindowsMediaPlayer1.URL = g
                                                        Label16.Caption = "正在播放：" & Label7.Caption
                                                        simble = 7
                                ElseIf WindowsMediaPlayer1.URL = g And simble = 7 Then
                                                        WindowsMediaPlayer1.URL = c
                                                        Label16.Caption = "正在播放：" & Label3.Caption
                                                        simble = 3
                                  ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                        WindowsMediaPlayer1.URL = i
                                                        Label16.Caption = "正在播放：" & Label9.Caption
                                                        simble = 9
                                ElseIf WindowsMediaPlayer1.URL = i And simble = 9 Then
                                                        WindowsMediaPlayer1.URL = e
                                                        Label16.Caption = "正在播放：" & Label5.Caption
                                                        simble = 5
                                ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                                        WindowsMediaPlayer1.URL = b
                                                        Label16.Caption = "正在播放：" & Label2.Caption
                                                        simble = 2
                                 ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                        WindowsMediaPlayer1.URL = f
                                                        Label16.Caption = "正在播放：" & Label6.Caption
                                                        simble = 6
                                 ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
                                                        WindowsMediaPlayer1.URL = h
                                                        Label16.Caption = "正在播放：" & Label8.Caption
                                                        simble = 8
                             ElseIf WindowsMediaPlayer1.URL = h And simble = 8 Then
                                                            WindowsMediaPlayer1.URL = a
                                                            Label16.Caption = "正在播放：" & Label1.Caption
                                                            simble = 1
                                End If
            ElseIf Label1.Caption <> "" And Label2.Caption <> "" And Label3.Caption <> "" And Label4.Caption <> "" And Label5.Caption <> "" And Label6.Caption <> "" And Label7.Caption <> "" And Label8.Caption <> "" And Label9.Caption <> "" And Label10.Caption <> "" Then
                             If WindowsMediaPlayer1.URL = a And simble = 1 Then
                                                            WindowsMediaPlayer1.URL = d
                                                            Label16.Caption = "正在播放：" & Label4.Caption
                                                            simble = 4
                             ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
                                                            WindowsMediaPlayer1.URL = g
                                                            Label16.Caption = "正在播放：" & Label7.Caption
                                                            simble = 7
                               ElseIf WindowsMediaPlayer1.URL = g And simble = 7 Then
                                                            WindowsMediaPlayer1.URL = c
                                                            Label16.Caption = "正在播放：" & Label3.Caption
                                                            simble = 3
                           ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
                                                                WindowsMediaPlayer1.URL = i
                                                                Label16.Caption = "正在播放：" & Label9.Caption
                                                                simble = 9
                              ElseIf WindowsMediaPlayer1.URL = i And simble = 9 Then
                                                            WindowsMediaPlayer1.URL = e
                                                            Label16.Caption = "正在播放：" & Label5.Caption
                                                            simble = 5
                             ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
                                                            WindowsMediaPlayer1.URL = b
                                                            Label16.Caption = "正在播放：" & Label2.Caption
                                                            simble = 2
                            ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
                                                            WindowsMediaPlayer1.URL = j
                                                            Label16.Caption = "正在播放：" & Label10.Caption
                                                            simble = 10
                            ElseIf WindowsMediaPlayer1.URL = j And simble = 10 Then
                                                            WindowsMediaPlayer1.URL = f
                                                            Label16.Caption = "正在播放：" & Label6.Caption
                                                            simble = 6
                             ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
                                                                WindowsMediaPlayer1.URL = h
                                                                Label16.Caption = "正在播放：" & Label8.Caption
                                                                simble = 8
                            ElseIf WindowsMediaPlayer1.URL = h And simble = 8 Then
                                                            WindowsMediaPlayer1.URL = a
                                                            Label16.Caption = "正在播放：" & Label1.Caption
                                                            simble = 1
                              End If
                End If
        End If


'随机播放——我的收藏——————————————————————————————————————————————————————

    If WindowsMediaPlayer1.URL = k Or WindowsMediaPlayer1.URL = l Or WindowsMediaPlayer1.URL = ma Or WindowsMediaPlayer1.URL = n Or WindowsMediaPlayer1.URL = o Or WindowsMediaPlayer1.URL = p Or WindowsMediaPlayer1.URL = q Or WindowsMediaPlayer1.URL = r Or WindowsMediaPlayer1.URL = s Or WindowsMediaPlayer1.URL = t Then
                         If Label23.Caption <> "" And Label24.Caption = "" Then
                                     If WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                        WindowsMediaPlayer1.URL = k
                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                        simble = 11
                                     End If
                        ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption = "" Then
                            If WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                        WindowsMediaPlayer1.URL = l
                                                        Label16.Caption = "正在播放：" & Label24.Caption
                                                        simble = 12
                             ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                        WindowsMediaPlayer1.URL = k
                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                        simble = 11
                             End If
                         ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption = "" Then
                            If WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                        WindowsMediaPlayer1.URL = ma
                                                        Label16.Caption = "正在播放：" & Label25.Caption
                                                        simble = 13
                             ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                        WindowsMediaPlayer1.URL = l
                                                        Label16.Caption = "正在播放：" & Label24.Caption
                                                        simble = 12
                             ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                        WindowsMediaPlayer1.URL = k
                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                        simble = 11
                             End If
                      ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption = "" Then
                                    If WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                        WindowsMediaPlayer1.URL = n
                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                        simble = 14
                              ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                        WindowsMediaPlayer1.URL = ma
                                                        Label16.Caption = "正在播放：" & Label25.Caption
                                                        simble = 13
                             ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                        WindowsMediaPlayer1.URL = l
                                                        Label16.Caption = "正在播放：" & Label24.Caption
                                                        simble = 12
                             ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                        WindowsMediaPlayer1.URL = k
                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                        simble = 11
                             End If
                     ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption = "" Then
                                 If WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                        WindowsMediaPlayer1.URL = n
                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                        simble = 14
                              ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                        WindowsMediaPlayer1.URL = ma
                                                        Label16.Caption = "正在播放：" & Label25.Caption
                                                        simble = 13
                             ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                        WindowsMediaPlayer1.URL = o
                                                        Label16.Caption = "正在播放：" & Label27.Caption
                                                        simble = 15
                             ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                                        WindowsMediaPlayer1.URL = l
                                                        Label16.Caption = "正在播放：" & Label24.Caption
                                                        simble = 12
                             ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                        WindowsMediaPlayer1.URL = k
                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                        simble = 11
                             End If
             ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption = "" Then
                              If WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                        WindowsMediaPlayer1.URL = n
                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                        simble = 14
                              ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                        WindowsMediaPlayer1.URL = ma
                                                        Label16.Caption = "正在播放：" & Label25.Caption
                                                        simble = 13
                             ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                        WindowsMediaPlayer1.URL = o
                                                        Label16.Caption = "正在播放：" & Label27.Caption
                                                        simble = 15
                             ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                                        WindowsMediaPlayer1.URL = l
                                                        Label16.Caption = "正在播放：" & Label24.Caption
                                                        simble = 12
                              ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                        WindowsMediaPlayer1.URL = p
                                                        Label16.Caption = "正在播放：" & Label28.Caption
                                                        simble = 16
                             ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
                                                        WindowsMediaPlayer1.URL = k
                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                        simble = 11
                             End If
             ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption = "" Then
                                  If WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                        WindowsMediaPlayer1.URL = n
                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                        simble = 14
                             ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                        WindowsMediaPlayer1.URL = q
                                                        Label16.Caption = "正在播放：" & Label29.Caption
                                                        simble = 17
                              ElseIf WindowsMediaPlayer1.URL = q And simble = 17 Then
                                                        WindowsMediaPlayer1.URL = ma
                                                        Label16.Caption = "正在播放：" & Label25.Caption
                                                        simble = 13
                             ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                        WindowsMediaPlayer1.URL = o
                                                        Label16.Caption = "正在播放：" & Label27.Caption
                                                        simble = 15
                             ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                                        WindowsMediaPlayer1.URL = l
                                                        Label16.Caption = "正在播放：" & Label24.Caption
                                                        simble = 12
                              ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                        WindowsMediaPlayer1.URL = p
                                                        Label16.Caption = "正在播放：" & Label28.Caption
                                                        simble = 16
                             ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
                                                        WindowsMediaPlayer1.URL = k
                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                        simble = 11
                             End If
                 ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption = "" Then
                              If WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                        WindowsMediaPlayer1.URL = n
                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                        simble = 14
                             ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                        WindowsMediaPlayer1.URL = q
                                                        Label16.Caption = "正在播放：" & Label29.Caption
                                                        simble = 17
                              ElseIf WindowsMediaPlayer1.URL = q And simble = 17 Then
                                                        WindowsMediaPlayer1.URL = ma
                                                        Label16.Caption = "正在播放：" & Label25.Caption
                                                        simble = 13
                             ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                        WindowsMediaPlayer1.URL = o
                                                        Label16.Caption = "正在播放：" & Label27.Caption
                                                        simble = 15
                             ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                                        WindowsMediaPlayer1.URL = l
                                                        Label16.Caption = "正在播放：" & Label24.Caption
                                                        simble = 12
                              ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                        WindowsMediaPlayer1.URL = p
                                                        Label16.Caption = "正在播放：" & Label28.Caption
                                                        simble = 16
                              ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
                                                        WindowsMediaPlayer1.URL = r
                                                        Label16.Caption = "正在播放：" & Label30.Caption
                                                        simble = 18
                             ElseIf WindowsMediaPlayer1.URL = r And simble = 18 Then
                                                        WindowsMediaPlayer1.URL = k
                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                        simble = 11
                             End If
                     ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption = "" Then
                                If WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                        WindowsMediaPlayer1.URL = n
                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                        simble = 14
                         ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                        WindowsMediaPlayer1.URL = q
                                                        Label16.Caption = "正在播放：" & Label29.Caption
                                                        simble = 17
                        ElseIf WindowsMediaPlayer1.URL = q And simble = 17 Then
                                                        WindowsMediaPlayer1.URL = ma
                                                        Label16.Caption = "正在播放：" & Label25.Caption
                                                        simble = 13
                          ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                        WindowsMediaPlayer1.URL = s
                                                        Label16.Caption = "正在播放：" & Label31.Caption
                                                        simble = 19
                        ElseIf WindowsMediaPlayer1.URL = s And simble = 19 Then
                                                        WindowsMediaPlayer1.URL = o
                                                        Label16.Caption = "正在播放：" & Label27.Caption
                                                        simble = 15
                        ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                                        WindowsMediaPlayer1.URL = l
                                                        Label16.Caption = "正在播放：" & Label24.Caption
                                                        simble = 12
                    ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                        WindowsMediaPlayer1.URL = p
                                                        Label16.Caption = "正在播放：" & Label28.Caption
                                                        simble = 16
                     ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
                                                        WindowsMediaPlayer1.URL = r
                                                        Label16.Caption = "正在播放：" & Label30.Caption
                                                        simble = 18
                     ElseIf WindowsMediaPlayer1.URL = r And simble = 18 Then
                                                        WindowsMediaPlayer1.URL = k
                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                        simble = 11
                    End If
             ElseIf Label23.Caption <> "" And Label24.Caption <> "" And Label25.Caption <> "" And Label26.Caption <> "" And Label27.Caption <> "" And Label28.Caption <> "" And Label29.Caption <> "" And Label30.Caption <> "" And Label31.Caption <> "" And Label32.Caption <> "" Then
                         If WindowsMediaPlayer1.URL = k And simble = 11 Then
                                                        WindowsMediaPlayer1.URL = n
                                                        Label16.Caption = "正在播放：" & Label26.Caption
                                                        simble = 14
                         ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
                                                        WindowsMediaPlayer1.URL = q
                                                        Label16.Caption = "正在播放：" & Label29.Caption
                                                        simble = 17
                        ElseIf WindowsMediaPlayer1.URL = q And simble = 17 Then
                                                        WindowsMediaPlayer1.URL = ma
                                                        Label16.Caption = "正在播放：" & Label25.Caption
                                                        simble = 13
                          ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
                                                        WindowsMediaPlayer1.URL = s
                                                        Label16.Caption = "正在播放：" & Label31.Caption
                                                        simble = 19
                        ElseIf WindowsMediaPlayer1.URL = s And simble = 19 Then
                                                        WindowsMediaPlayer1.URL = o
                                                        Label16.Caption = "正在播放：" & Label27.Caption
                                                        simble = 15
                        ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
                                                        WindowsMediaPlayer1.URL = l
                                                        Label16.Caption = "正在播放：" & Label24.Caption
                                                        simble = 12
                     ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
                                                        WindowsMediaPlayer1.URL = t
                                                        Label16.Caption = "正在播放：" & Label32.Caption
                                                        simble = 20
                    ElseIf WindowsMediaPlayer1.URL = t And simble = 20 Then
                                                        WindowsMediaPlayer1.URL = p
                                                        Label16.Caption = "正在播放：" & Label28.Caption
                                                        simble = 16
                     ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
                                                        WindowsMediaPlayer1.URL = r
                                                        Label16.Caption = "正在播放：" & Label30.Caption
                                                        simble = 18
                     ElseIf WindowsMediaPlayer1.URL = r And simble = 18 Then
                                                        WindowsMediaPlayer1.URL = k
                                                        Label16.Caption = "正在播放：" & Label23.Caption
                                                        simble = 11
                    End If
             End If
     End If
  End If
End If
'单曲循环————————————————————————————————————————————————————————————
If Label18.Caption = Label19.Caption Then
  If Form3.Option4.Value = True Then

          If WindowsMediaPlayer1.URL = a And simble = 1 Then
             WindowsMediaPlayer1.URL = a
             Label16.Caption = "正在播放：" & Label1.Caption
             simble = 1
       
          ElseIf WindowsMediaPlayer1.URL = b And simble = 2 Then
             WindowsMediaPlayer1.URL = b
             Label16.Caption = "正在播放：" & Label2.Caption
             simble = 2
      
          ElseIf WindowsMediaPlayer1.URL = c And simble = 3 Then
             WindowsMediaPlayer1.URL = c
             Label16.Caption = "正在播放：" & Label3.Caption
             simble = 3
         
          ElseIf WindowsMediaPlayer1.URL = d And simble = 4 Then
             WindowsMediaPlayer1.URL = d
             Label16.Caption = "正在播放：" & Label4.Caption
             simble = 4

          ElseIf WindowsMediaPlayer1.URL = e And simble = 5 Then
             WindowsMediaPlayer1.URL = e
             Label16.Caption = "正在播放：" & Label5.Caption
             simble = 5
  
          ElseIf WindowsMediaPlayer1.URL = f And simble = 6 Then
             WindowsMediaPlayer1.URL = f
             Label16.Caption = "正在播放：" & Label6.Caption
             simble = 6
    
          ElseIf WindowsMediaPlayer1.URL = g And simble = 7 Then
             WindowsMediaPlayer1.URL = g
             Label16.Caption = "正在播放：" & Label7.Caption
             simble = 7
  
          ElseIf WindowsMediaPlayer1.URL = h And simble = 8 Then
             WindowsMediaPlayer1.URL = h
             Label16.Caption = "正在播放：" & Label8.Caption
             simble = 8
   
          ElseIf WindowsMediaPlayer1.URL = i And simble = 9 Then
             WindowsMediaPlayer1.URL = i
             Label16.Caption = "正在播放：" & Label9.Caption
             simble = 9
     
          ElseIf WindowsMediaPlayer1.URL = j And simble = 10 Then
             WindowsMediaPlayer1.URL = j
             Label16.Caption = "正在播放：" & Label10.Caption
             simble = 10
             
'单曲循环——我的收藏————————————————————————————————————————————————————————

          ElseIf WindowsMediaPlayer1.URL = k And simble = 11 Then
             WindowsMediaPlayer1.URL = k
             Label16.Caption = "正在播放：" & Label23.Caption
             simble = 11
      
          ElseIf WindowsMediaPlayer1.URL = l And simble = 12 Then
             WindowsMediaPlayer1.URL = l
             Label16.Caption = "正在播放：" & Label24.Caption
             simble = 12
        
          ElseIf WindowsMediaPlayer1.URL = ma And simble = 13 Then
             WindowsMediaPlayer1.URL = ma
             Label16.Caption = "正在播放：" & Label25.Caption
             simble = 13
                
          ElseIf WindowsMediaPlayer1.URL = n And simble = 14 Then
             WindowsMediaPlayer1.URL = n
             Label16.Caption = "正在播放：" & Label26.Caption
             simble = 14
    
          ElseIf WindowsMediaPlayer1.URL = o And simble = 15 Then
             WindowsMediaPlayer1.URL = o
             Label16.Caption = "正在播放：" & Label27.Caption
             simble = 5

          ElseIf WindowsMediaPlayer1.URL = p And simble = 16 Then
             WindowsMediaPlayer1.URL = p
             Label16.Caption = "正在播放：" & Label28.Caption
             simble = 16
    
          ElseIf WindowsMediaPlayer1.URL = q And simble = 17 Then
             WindowsMediaPlayer1.URL = q
             Label16.Caption = "正在播放：" & Label29.Caption
             simble = 17
  
          ElseIf WindowsMediaPlayer1.URL = r And simble = 18 Then
             WindowsMediaPlayer1.URL = r
             Label16.Caption = "正在播放：" & Label30.Caption
             simble = 18

          ElseIf WindowsMediaPlayer1.URL = s And simble = 19 Then
             WindowsMediaPlayer1.URL = s
             Label16.Caption = "正在播放：" & Label31.Caption
             simble = 19

          ElseIf WindowsMediaPlayer1.URL = t And simble = 20 Then
             WindowsMediaPlayer1.URL = t
             Label16.Caption = "正在播放：" & Label32.Caption
             simble = 20
             
          End If
     End If
     
   Timer5.Enabled = True
 End If




          
          
 



If Label12.Caption = "00:00" And Label13.Caption = "00:00" And WindowsMediaPlayer1.playState = 10 Then
   Timer3.Enabled = True
End If
If WindowsMediaPlayer1.playState = 3 Then
   Timer3.Enabled = False
ElseIf Label18.Caption = Label19.Caption Then
   Timer3.Enabled = True
End If




 If Label22.Caption = "" Then
      Form10.load.Enabled = True
      Form10.unload.Enabled = False
   ElseIf Label22.Caption <> "" Then
      Form10.load.Enabled = False
      Form10.unload.Enabled = True
   End If

End Sub

Private Sub Timer3_Timer()
On Error Resume Next
If WindowsMediaPlayer1.URL <> "" Then
 If WindowsMediaPlayer1.playState <> wmppsPlaying And WindowsMediaPlayer1.playState <> wmppsPaused And WindowsMediaPlayer1.playState <> wmppsStopped Then
      If simble = 1 Then
           Label1.Caption = ""
      ElseIf simble = 2 Then
        
          Label2.Caption = ""
      ElseIf simble = 3 Then
       
          Label3.Caption = ""
      ElseIf simble = 4 Then
        
           Label4.Caption = ""
      ElseIf simble = 5 Then
           Label5.Caption = ""
      ElseIf simble = 6 Then
          Label6.Caption = ""
      ElseIf simble = 7 Then
           Label7.Caption = ""
      ElseIf simble = 8 Then
           Label8.Caption = ""
      ElseIf simble = 9 Then
           Label9.Caption = ""
      ElseIf simble = 10 Then
           Label10.Caption = ""
      End If
      
      
 
'我的收藏——————————————————————————————————————————————————

   If simble = 11 Then
          Label23.Caption = ""
   ElseIf simble = 12 Then
           Label24.Caption = ""
   ElseIf simble = 13 Then
      
          Label25.Caption = ""
   ElseIf simble = 14 Then
     
          Label26.Caption = ""
  ElseIf simble = 15 Then
    
         Label27.Caption = ""
  ElseIf simble = 16 Then
      
         Label28.Caption = ""
  ElseIf simble = 17 Then
      
          Label29.Caption = ""
    ElseIf simble = 18 Then
      
         Label30.Caption = ""
  ElseIf simble = 19 Then
      
          Label31.Caption = ""
   ElseIf simble = 20 Then
     
           Label32.Caption = ""
    End If
     


   
If Form1.Label1.Caption = "" Then
   m = a
   Form1.Label1.Caption = Label2.Caption
   a = b
   Label2.Caption = Label3.Caption
   b = c
   Label3.Caption = Label4.Caption
   c = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "1" Then
     If Label1.Caption <> "" Then
                WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
     Else
        WindowsMediaPlayer1.URL = ""
        Label16.Caption = "百度音乐_音乐你的生活"
        Label16.Alignment = 2
           Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label2.Caption = "" Then
   m = b
   Label2.Caption = Label3.Caption
   b = c
   Label3.Caption = Label4.Caption
   c = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "2" Then
       If Label2.Caption <> "" Then
                WindowsMediaPlayer1.URL = b
                Label16.Caption = "正在播放：" & Label2.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
                Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label3.Caption = "" Then
   m = c
   Label3.Caption = Label4.Caption
   c = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "3" Then
       If Label3.Caption <> "" Then
                WindowsMediaPlayer1.URL = c
                Label16.Caption = "正在播放：" & Label3.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
            Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label4.Caption = "" Then
   m = d
   Label4.Caption = Label5.Caption
   d = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "4" Then
       If Label4.Caption <> "" Then
                WindowsMediaPlayer1.URL = d
                Label16.Caption = "正在播放：" & Label4.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
                Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label5.Caption = "" Then
   m = e
   Label5.Caption = Label6.Caption
   e = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "5" Then
       If Label5.Caption <> "" Then
                WindowsMediaPlayer1.URL = e
                Label16.Caption = "正在播放：" & Label5.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
                Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label6.Caption = "" Then
   m = f
   Label6.Caption = Label7.Caption
   f = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "6" Then
       If Label6.Caption <> "" Then
                WindowsMediaPlayer1.URL = f
                Label16.Caption = "正在播放：" & Label6.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
                Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label7.Caption = "" Then
   m = g
   Label7.Caption = Label8.Caption
   g = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "7" Then
       If Label7.Caption <> "" Then
                WindowsMediaPlayer1.URL = g
                Label16.Caption = "正在播放：" & Label7.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
                Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label8.Caption = "" Then
   m = h
   Label8.Caption = Label9.Caption
   h = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "8" Then
       If Label8.Caption <> "" Then
                WindowsMediaPlayer1.URL = h
                Label16.Caption = "正在播放：" & Label8.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
                Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label9.Caption = "" Then
   m = i
   Label9.Caption = Label10.Caption
   i = j
   Label10.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "9" Then
       If Label9.Caption <> "" Then
                WindowsMediaPlayer1.URL = i
                Label16.Caption = "正在播放：" & Label9.Caption
       Else
                 WindowsMediaPlayer1.URL = a
                Label16.Caption = "正在播放：" & Label1.Caption
                Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
       End If
   End If
End If
'我的收藏————————————————————————————————————————————————————
If Form1.Label23.Caption = "" Then
   m = k
   Label23.Caption = Label24.Caption
   k = l
   Label24.Caption = Label25.Caption
   l = ma
   Label25.Caption = Label26.Caption
   ma = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "11" Then
       If Label23.Caption <> "" Then
                WindowsMediaPlayer1.URL = k
                Label16.Caption = "正在播放：" & Label23.Caption
       Else
                 WindowsMediaPlayer1.URL = ""
                 Label16.Caption = "百度音乐_音乐你的生活"
                 Label16.Alignment = 2
                   Label23.BackColor = &H8000000F
                   Label25.BackColor = &H8000000F
                   Label27.BackColor = &H8000000F
                   Label29.BackColor = &H8000000F
                   Label31.BackColor = &H8000000F
        
                   Label24.BackColor = &HFFC0FF
                   Label26.BackColor = &HFFC0FF
                   Label28.BackColor = &HFFC0FF
                   Label30.BackColor = &HFFC0FF
                   Label32.BackColor = &HFFC0FF
                   
                   Label23.FontSize = 9
                   Label23.ForeColor = RGB(0, 0, 0)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
       End If
   End If
ElseIf Form1.Label24.Caption = "" Then
   m = l
   Label24.Caption = Label25.Caption
   l = ma
   Label25.Caption = Label26.Caption
   ma = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "12" Then
     If Label24.Caption <> "" Then
            WindowsMediaPlayer1.URL = l
            Label16.Caption = "正在播放：" & Label24.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                               Label23.BackColor = &H8000000F
                   Label25.BackColor = &H8000000F
                   Label27.BackColor = &H8000000F
                   Label29.BackColor = &H8000000F
                   Label31.BackColor = &H8000000F
        
                   Label24.BackColor = &HFFC0FF
                   Label26.BackColor = &HFFC0FF
                   Label28.BackColor = &HFFC0FF
                   Label30.BackColor = &HFFC0FF
                   Label32.BackColor = &HFFC0FF
                   
                   Label23.FontSize = 9
                   Label23.ForeColor = RGB(0, 0, 0)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label25.Caption = "" Then
   m = ma
   Label25.Caption = Label26.Caption
   ma = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "13" Then
     If Label25.Caption <> "" Then
            WindowsMediaPlayer1.URL = ma
            Label16.Caption = "正在播放：" & Label25.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                               Label23.BackColor = &H8000000F
                   Label25.BackColor = &H8000000F
                   Label27.BackColor = &H8000000F
                   Label29.BackColor = &H8000000F
                   Label31.BackColor = &H8000000F
        
                   Label24.BackColor = &HFFC0FF
                   Label26.BackColor = &HFFC0FF
                   Label28.BackColor = &HFFC0FF
                   Label30.BackColor = &HFFC0FF
                   Label32.BackColor = &HFFC0FF
                   
                   Label23.FontSize = 9
                   Label23.ForeColor = RGB(0, 0, 0)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label26.Caption = "" Then
   m = n
   Label26.Caption = Label27.Caption
   n = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "14" Then
     If Label26.Caption <> "" Then
            WindowsMediaPlayer1.URL = n
            Label16.Caption = "正在播放：" & Label26.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                               Label23.BackColor = &H8000000F
                   Label25.BackColor = &H8000000F
                   Label27.BackColor = &H8000000F
                   Label29.BackColor = &H8000000F
                   Label31.BackColor = &H8000000F
        
                   Label24.BackColor = &HFFC0FF
                   Label26.BackColor = &HFFC0FF
                   Label28.BackColor = &HFFC0FF
                   Label30.BackColor = &HFFC0FF
                   Label32.BackColor = &HFFC0FF
                   
                   Label23.FontSize = 9
                   Label23.ForeColor = RGB(0, 0, 0)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label27.Caption = "" Then
   m = o
   Label27.Caption = Label28.Caption
   o = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "15" Then
     If Label27.Caption <> "" Then
            WindowsMediaPlayer1.URL = o
            Label16.Caption = "正在播放：" & Label27.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                               Label23.BackColor = &H8000000F
                   Label25.BackColor = &H8000000F
                   Label27.BackColor = &H8000000F
                   Label29.BackColor = &H8000000F
                   Label31.BackColor = &H8000000F
        
                   Label24.BackColor = &HFFC0FF
                   Label26.BackColor = &HFFC0FF
                   Label28.BackColor = &HFFC0FF
                   Label30.BackColor = &HFFC0FF
                   Label32.BackColor = &HFFC0FF
                   
                   Label23.FontSize = 9
                   Label23.ForeColor = RGB(0, 0, 0)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label28.Caption = "" Then
   m = p
   Label28.Caption = Label29.Caption
   p = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "16" Then
     If Label28.Caption <> "" Then
            WindowsMediaPlayer1.URL = p
            Label16.Caption = "正在播放：" & Label28.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                               Label23.BackColor = &H8000000F
                   Label25.BackColor = &H8000000F
                   Label27.BackColor = &H8000000F
                   Label29.BackColor = &H8000000F
                   Label31.BackColor = &H8000000F
        
                   Label24.BackColor = &HFFC0FF
                   Label26.BackColor = &HFFC0FF
                   Label28.BackColor = &HFFC0FF
                   Label30.BackColor = &HFFC0FF
                   Label32.BackColor = &HFFC0FF
                   
                   Label23.FontSize = 9
                   Label23.ForeColor = RGB(0, 0, 0)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label29.Caption = "" Then
   m = q
   Label29.Caption = Label30.Caption
   q = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "17" Then
     If Label29.Caption <> "" Then
            WindowsMediaPlayer1.URL = q
            Label16.Caption = "正在播放：" & Label29.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                               Label23.BackColor = &H8000000F
                   Label25.BackColor = &H8000000F
                   Label27.BackColor = &H8000000F
                   Label29.BackColor = &H8000000F
                   Label31.BackColor = &H8000000F
        
                   Label24.BackColor = &HFFC0FF
                   Label26.BackColor = &HFFC0FF
                   Label28.BackColor = &HFFC0FF
                   Label30.BackColor = &HFFC0FF
                   Label32.BackColor = &HFFC0FF
                   
                   Label23.FontSize = 9
                   Label23.ForeColor = RGB(0, 0, 0)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label30.Caption = "" Then
   m = r
   Label30.Caption = Label31.Caption
   r = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "18" Then
     If Label30.Caption <> "" Then
            WindowsMediaPlayer1.URL = r
            Label16.Caption = "正在播放：" & Label30.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                               Label23.BackColor = &H8000000F
                   Label25.BackColor = &H8000000F
                   Label27.BackColor = &H8000000F
                   Label29.BackColor = &H8000000F
                   Label31.BackColor = &H8000000F
        
                   Label24.BackColor = &HFFC0FF
                   Label26.BackColor = &HFFC0FF
                   Label28.BackColor = &HFFC0FF
                   Label30.BackColor = &HFFC0FF
                   Label32.BackColor = &HFFC0FF
                   
                   Label23.FontSize = 9
                   Label23.ForeColor = RGB(0, 0, 0)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
ElseIf Form1.Label31.Caption = "" Then
   m = s
   Label31.Caption = Label32.Caption
   s = t
   Label32.Caption = ""
   If WindowsMediaPlayer1.URL = m And simble = "19" Then
     If Label31.Caption <> "" Then
            WindowsMediaPlayer1.URL = s
            Label16.Caption = "正在播放：" & Label31.Caption
     Else
            WindowsMediaPlayer1.URL = k
            Label16.Caption = "正在播放：" & Label23.Caption
                               Label23.BackColor = &H8000000F
                   Label25.BackColor = &H8000000F
                   Label27.BackColor = &H8000000F
                   Label29.BackColor = &H8000000F
                   Label31.BackColor = &H8000000F
        
                   Label24.BackColor = &HFFC0FF
                   Label26.BackColor = &HFFC0FF
                   Label28.BackColor = &HFFC0FF
                   Label30.BackColor = &HFFC0FF
                   Label32.BackColor = &HFFC0FF
                   
                   Label23.FontSize = 9
                   Label23.ForeColor = RGB(0, 0, 0)
                   Label24.FontSize = 9
                   Label24.ForeColor = RGB(0, 0, 0)
                   Label25.FontSize = 9
                   Label25.ForeColor = RGB(0, 0, 0)
                   Label26.FontSize = 9
                   Label26.ForeColor = RGB(0, 0, 0)
                   Label27.FontSize = 9
                   Label27.ForeColor = RGB(0, 0, 0)
                   Label28.FontSize = 9
                   Label28.ForeColor = RGB(0, 0, 0)
                   Label29.FontSize = 9
                   Label29.ForeColor = RGB(0, 0, 0)
                   Label30.FontSize = 9
                   Label30.ForeColor = RGB(0, 0, 0)
                   Label31.FontSize = 9
                   Label31.ForeColor = RGB(0, 0, 0)
                   Label32.FontSize = 9
                   Label32.ForeColor = RGB(0, 0, 0)
     End If
   End If
End If

If Label1.Caption = "" Then
   Label1.BackColor = &H8000000F
   Label1.FontSize = 9
   Label1.ForeColor = RGB(0, 0, 0)
ElseIf Label2.Caption = "" Then
  Label2.BackColor = &HFFC0FF
  Label2.FontSize = 9
  Label2.ForeColor = RGB(0, 0, 0)
ElseIf Label3.Caption = "" Then
  Label3.BackColor = &H8000000F
  Label3.FontSize = 9
  Label3.ForeColor = RGB(0, 0, 0)
ElseIf Label4.Caption = "" Then
  Label4.BackColor = &HFFC0FF
  Label4.FontSize = 9
  Label4.ForeColor = RGB(0, 0, 0)
ElseIf Label5.Caption = "" Then
  Label5.BackColor = &H8000000F
  Label5.FontSize = 9
  Label5.ForeColor = RGB(0, 0, 0)
ElseIf Label6.Caption = "" Then
  Label6.BackColor = &HFFC0FF
  Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
ElseIf Label7.Caption = "" Then
  Label7.BackColor = &H8000000F
  Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
ElseIf Label8.Caption = "" Then
  Label8.BackColor = &HFFC0FF
  Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
ElseIf Label9.Caption = "" Then
  Label9.BackColor = &H8000000F
  Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
ElseIf Label10.Caption = "" Then
  Label10.BackColor = &HFFC0FF
  Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
End If
'my count________________________________________________________
If Label23.Caption = "" Then
  Label23.BackColor = &H8000000F
  Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
ElseIf Label24.Caption = "" Then
Label24.BackColor = &HFFC0FF
Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
ElseIf Label25.Caption = "" Then
Label25.BackColor = &H8000000F
 Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
ElseIf Label26.Caption = "" Then
 Label26.BackColor = &HFFC0FF
 Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
ElseIf Label27.Caption = "" Then
Label27.BackColor = &H8000000F
Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
ElseIf Label28.Caption = "" Then
Label28.BackColor = &HFFC0FF
Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
ElseIf Label29.Caption = "" Then
Label29.BackColor = &H8000000F
Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
ElseIf Label30.Caption = "" Then
Label30.BackColor = &HFFC0FF
Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
ElseIf Label31.Caption = "" Then
Label31.BackColor = &H8000000F
Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
ElseIf Label32.Caption = "" Then
Label32.BackColor = &HFFC0FF
 Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)

End If

Timer5.Enabled = True
Timer3.Enabled = False
End If
End If
End Sub

Private Sub Timer4_Timer()

 
   
If Label22.Caption <> "" Then
Dim a1, a2, a3, a4, a5, a6, a7, a8, a9, a10 As String
On Error Resume Next
If Dir("c:\百度音乐播放器\" & Label22.Caption & "的收藏.m3u") <> "" And Dir("c:\百度音乐播放器\" & Label22.Caption & "的收藏.txt") = "" Then
   Name "c:\百度音乐播放器\" & Label22.Caption & "的收藏.m3u" As "c:\百度音乐播放器\" & Label22.Caption & "的收藏.txt"
End If

If Dir("c:\百度音乐播放器\" & Label22.Caption & "的收藏.txt") <> "" Then
   Open "c:\百度音乐播放器\" & Label22.Caption & "的收藏.txt" For Input As #2
    Line Input #2, a1
      Label23.Caption = a1
    Line Input #2, k
    Line Input #2, a2
      Label24.Caption = a2
    Line Input #2, l
    Line Input #2, a3
      Label25.Caption = a3
    Line Input #2, ma
    Line Input #2, a4
      Label26.Caption = a4
    Line Input #2, n
    Line Input #2, a5
      Label27.Caption = a5
    Line Input #2, o
    Line Input #2, a6
      Label28.Caption = a6
    Line Input #2, p
    Line Input #2, a7
      Label29.Caption = a7
    Line Input #2, q
    Line Input #2, a8
      Label30.Caption = a8
    Line Input #2, r
    Line Input #2, a9
      Label31.Caption = a9
    Line Input #2, s
    Line Input #2, a10
      Label32.Caption = a10
    Line Input #2, t
    Line Input #2, pic
    Image1.Picture = LoadPicture(pic)
    Form3.Image1.Picture = LoadPicture(pic)
   Close #2
 End If
 

 Form3.Label4.Caption = Form1.Label20.Caption
  
 
  Label23.Visible = True
  Label24.Visible = True
  Label25.Visible = True
  Label26.Visible = True
  Label27.Visible = True
  Label28.Visible = True
  Label29.Visible = True
  Label30.Visible = True
  Label31.Visible = True
  Label32.Visible = True
  
  Picture2.Visible = False
  Picture13.Visible = True
  Picture15.Visible = True
  Picture3.Visible = False
  Picture16.Visible = False
  
            Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
           
           
           With Form3
                .Label2.Visible = False
                .Label13.Visible = False
                .Option1.Visible = False
               .Option2.Visible = False
               .Option3.Visible = False
               .Option4.Visible = False
               .Check4.Visible = False
               
               .Label1.Visible = False
             .Label3.Visible = False
       .Label4.Visible = True
    

        .Label2.Visible = False
       .Label13.Visible = False
        .Option1.Visible = False
        .Option2.Visible = False
        .Option3.Visible = False
        .Option4.Visible = False
        .Check4.Visible = False
        
        
        .Label5.Visible = True
        .Label6.Visible = True
        .Check1.Visible = True
        .Check2.Visible = True
        .Image1.Visible = True
        .Picture15.Visible = True
        .Label12.Visible = True
       
        .Label8.Visible = True
       .Label9.Visible = True
        .Label10.Visible = True
        .Label14.Visible = True
        .Text1.Visible = True
        .Label9.Caption = Form7.Combo1.Text
        .Text1.Text = Form1.Label22.Caption
       
        .Picture5.Top = 840
        .Picture6.Top = 840
        .Picture7.Top = 840
           End With
           Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
If simble = "1" Then
           Label1.BackColor = &HFF8080
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 13
           Label1.ForeColor = RGB(255, 255, 255)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "2" Then
 Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFF8080
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 13
           Label2.ForeColor = RGB(255, 255, 255)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "3" Then
Label1.BackColor = &H8000000F
           Label3.BackColor = &HFF8080
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
                           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 13
           Label3.ForeColor = RGB(255, 255, 255)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "4" Then
 Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFF8080
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 13
           Label4.ForeColor = RGB(255, 255, 255)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "5" Then
Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &HFF8080
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 13
           Label5.ForeColor = RGB(255, 255, 255)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "6" Then
 Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFF8080
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 13
           Label6.ForeColor = RGB(255, 255, 255)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "7" Then
Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &HFF8080
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 13
           Label7.ForeColor = RGB(255, 255, 255)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "8" Then
Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFF8080
           Label10.BackColor = &HFFC0FF
                        
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 13
           Label8.ForeColor = RGB(255, 255, 255)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "9" Then
Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &HFF8080

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFFC0FF
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 13
           Label9.ForeColor = RGB(255, 255, 255)
           Label10.FontSize = 9
           Label10.ForeColor = RGB(0, 0, 0)
ElseIf simble = "10" Then
 Label1.BackColor = &H8000000F
           Label3.BackColor = &H8000000F
           Label5.BackColor = &H8000000F
           Label7.BackColor = &H8000000F
           Label9.BackColor = &H8000000F

           Label2.BackColor = &HFFC0FF
           Label4.BackColor = &HFFC0FF
           Label6.BackColor = &HFFC0FF
           Label8.BackColor = &HFFC0FF
           Label10.BackColor = &HFF8080
           
           Label1.FontSize = 9
           Label1.ForeColor = RGB(0, 0, 0)
           Label2.FontSize = 9
           Label2.ForeColor = RGB(0, 0, 0)
           Label3.FontSize = 9
           Label3.ForeColor = RGB(0, 0, 0)
           Label4.FontSize = 9
           Label4.ForeColor = RGB(0, 0, 0)
           Label5.FontSize = 9
           Label5.ForeColor = RGB(0, 0, 0)
           Label6.FontSize = 9
           Label6.ForeColor = RGB(0, 0, 0)
           Label7.FontSize = 9
           Label7.ForeColor = RGB(0, 0, 0)
           Label8.FontSize = 9
           Label8.ForeColor = RGB(0, 0, 0)
           Label9.FontSize = 9
           Label9.ForeColor = RGB(0, 0, 0)
           Label10.FontSize = 13
           Label10.ForeColor = RGB(255, 255, 255)
ElseIf simble = "11" Then
Label23.BackColor = vbRed
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 13
           Label23.ForeColor = RGB(255, 255, 255)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "12" Then
 Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = vbRed
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 13
           Label24.ForeColor = RGB(255, 255, 255)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "13" Then
 Label23.BackColor = &H8000000F
           Label25.BackColor = vbRed
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 13
           Label25.ForeColor = RGB(255, 255, 255)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "14" Then
Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = vbRed
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 13
           Label26.ForeColor = RGB(255, 255, 255)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "15" Then
Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = vbRed
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 13
           Label27.ForeColor = RGB(255, 255, 255)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "16" Then
 Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = vbRed
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 13
           Label28.ForeColor = RGB(255, 255, 255)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "17" Then
Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = vbRed
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 13
           Label29.ForeColor = RGB(255, 255, 255)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "18" Then
Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = vbRed
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 13
           Label30.ForeColor = RGB(255, 255, 255)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "19" Then
 Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = vbRed

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = &HFFC0FF
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 13
           Label31.ForeColor = RGB(255, 255, 255)
           Label32.FontSize = 9
           Label32.ForeColor = RGB(0, 0, 0)
ElseIf simble = "20" Then
Label23.BackColor = &H8000000F
           Label25.BackColor = &H8000000F
           Label27.BackColor = &H8000000F
           Label29.BackColor = &H8000000F
           Label31.BackColor = &H8000000F

           Label24.BackColor = &HFFC0FF
           Label26.BackColor = &HFFC0FF
           Label28.BackColor = &HFFC0FF
           Label30.BackColor = &HFFC0FF
           Label32.BackColor = vbRed
           
           Label23.FontSize = 9
           Label23.ForeColor = RGB(0, 0, 0)
           Label24.FontSize = 9
           Label24.ForeColor = RGB(0, 0, 0)
           Label25.FontSize = 9
           Label25.ForeColor = RGB(0, 0, 0)
           Label26.FontSize = 9
           Label26.ForeColor = RGB(0, 0, 0)
           Label27.FontSize = 9
           Label27.ForeColor = RGB(0, 0, 0)
           Label28.FontSize = 9
           Label28.ForeColor = RGB(0, 0, 0)
           Label29.FontSize = 9
           Label29.ForeColor = RGB(0, 0, 0)
           Label30.FontSize = 9
           Label30.ForeColor = RGB(0, 0, 0)
           Label31.FontSize = 9
           Label31.ForeColor = RGB(0, 0, 0)
           Label32.FontSize = 13
           Label32.ForeColor = RGB(255, 255, 255)
End If
  
Timer5.Enabled = False
End Sub



Private Sub Timer6_Timer()
If Label34.Caption = "1" Then
   a = Label33.Caption
   simble = 1
   Label16.Caption = "正在播放：" & Label1.Caption
   
ElseIf Label34.Caption = "2" Then
   b = Label33.Caption
   simble = 2
   Label16.Caption = "正在播放：" & Label2.Caption
ElseIf Label34.Caption = "3" Then
   c = Label33.Caption
   simble = 3
   Label16.Caption = "正在播放：" & Label3.Caption
ElseIf Label34.Caption = "4" Then
   d = Label33.Caption
   simble = 4
   Label16.Caption = "正在播放：" & Label4.Caption
ElseIf Label34.Caption = "5" Then
   e = Label33.Caption
   simble = 5
   Label16.Caption = "正在播放：" & Label5.Caption
ElseIf Label34.Caption = "6" Then
   f = Label33.Caption
   simble = 6
   Label16.Caption = "正在播放：" & Label6.Caption
ElseIf Label34.Caption = "7" Then
   g = Label33.Caption
   simble = 7
   Label16.Caption = "正在播放：" & Label7.Caption
ElseIf Label34.Caption = "8" Then
   h = Label33.Caption
   simble = 8
   Label16.Caption = "正在播放：" & Label8.Caption
ElseIf Label34.Caption = "9" Then
   i = Label33.Caption
   simble = 9
   Label16.Caption = "正在播放：" & Label9.Caption
ElseIf Label34.Caption = "10" Then
   j = Label33.Caption
   simble = 10
   Label16.Caption = "正在播放：" & Label10.Caption
End If
   Timer5.Enabled = True
   Timer6.Enabled = False
End Sub

Private Sub Timer7_Timer()
If Form2.Visible = True Then
    If Form1.WindowState = 1 Then
       Form2.Hide
    ElseIf Form1.WindowState = 0 Then
       Form2.Show
    End If
End If

If Picture2.Visible = True Then
        If Label2.Caption = "" Then
           Image3.Visible = False
           Image4.Visible = False
        ElseIf Label2.Caption <> "" Then
           Image3.Visible = True
           Image4.Visible = True
        End If
ElseIf Picture15.Visible = True Then
        If Label24.Caption = "" Then
           Image3.Visible = False
           Image4.Visible = False
        ElseIf Label24.Caption <> "" Then
           Image3.Visible = True
           Image4.Visible = True
        End If
End If
End Sub

Private Sub Timer8_Timer()

End Sub

Private Sub WindowsMediaPlayer1_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)

  Text1.Enabled = False
If Text1.Text = "" Then
   Text1.Text = "搜索 歌曲、歌手、专辑"
   Text1.ForeColor = &H80000011
End If
  
  
End Sub

Private Sub WindowsMediaPlayer1_MouseMove(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
 Picture10.Visible = False
 Picture7.Visible = True
 Picture11.Visible = False
 Picture8.Visible = True
 Picture9.Visible = False
 Picture6.Visible = True
   
 If Picture2.Visible = False Then
 Picture13.Visible = True
 Picture14.Visible = False
 End If
 If Picture15.Visible = False Then
 Picture3.Visible = True
 Picture16.Visible = False
 End If
 If Picture18.Visible = False Then
 Picture4.Visible = True
 Picture17.Visible = False
 End If
End Sub

