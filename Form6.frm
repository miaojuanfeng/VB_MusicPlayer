VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "百度修改"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5535
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame9 
      Caption         =   "密保修改成功"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   4695
      Left            =   0
      TabIndex        =   78
      Top             =   1200
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command12 
         Caption         =   "完成"
         Height          =   255
         Left            =   4200
         TabIndex        =   24
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "请牢记您的新密保"
         Height          =   180
         Left            =   360
         TabIndex        =   89
         Top             =   3720
         Width           =   1440
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5160
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label54 
         Caption         =   "Label54"
         Height          =   255
         Left            =   1320
         TabIndex        =   88
         Top             =   3000
         Width           =   3495
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "答案："
         Height          =   180
         Left            =   480
         TabIndex        =   87
         Top             =   3000
         Width           =   540
      End
      Begin VB.Label Label52 
         Caption         =   "Label52"
         Height          =   255
         Left            =   1320
         TabIndex        =   86
         Top             =   2640
         Width           =   3615
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "密保问题："
         Height          =   180
         Left            =   240
         TabIndex        =   85
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label Label50 
         Caption         =   "Label50"
         Height          =   255
         Left            =   1320
         TabIndex        =   84
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "答案："
         Height          =   180
         Left            =   480
         TabIndex        =   83
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label Label48 
         Caption         =   "Label48"
         Height          =   255
         Left            =   1320
         TabIndex        =   82
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "密保问题1："
         Height          =   180
         Left            =   240
         TabIndex        =   81
         Top             =   1680
         Width           =   990
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "您的账户信息为："
         Height          =   180
         Left            =   600
         TabIndex        =   80
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "恭喜您，密保修改成功"
         Height          =   180
         Left            =   720
         TabIndex        =   79
         Top             =   600
         Width           =   1800
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "修改密保"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   4695
      Left            =   0
      TabIndex        =   71
      Top             =   1200
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   840
         Top             =   3600
      End
      Begin VB.CommandButton Command11 
         Caption         =   "下一步"
         Height          =   255
         Left            =   4320
         TabIndex        =   23
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox Text13 
         Height          =   270
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   22
         Top             =   3000
         Width           =   3255
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   21
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   20
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   19
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         Caption         =   "答案："
         Height          =   180
         Left            =   480
         TabIndex        =   77
         Top             =   3000
         Width           =   540
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "密保问题2："
         Height          =   180
         Left            =   240
         TabIndex        =   76
         Top             =   2520
         Width           =   990
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "答案："
         Height          =   180
         Left            =   480
         TabIndex        =   75
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "密保问题1："
         Height          =   180
         Left            =   240
         TabIndex        =   74
         Top             =   1560
         Width           =   990
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5160
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "请输入新密保，并且牢记新密保"
         Height          =   180
         Left            =   360
         TabIndex        =   73
         Top             =   960
         Width           =   2520
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "最后一步：输入新密保"
         Height          =   180
         Left            =   360
         TabIndex        =   72
         Top             =   480
         Width           =   1800
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "修改密保"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   4695
      Left            =   0
      TabIndex        =   60
      Top             =   1200
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command10 
         Caption         =   "下一步"
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton Command9 
         Caption         =   "上一步"
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   1080
         MaxLength       =   16
         TabIndex        =   16
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   1080
         MaxLength       =   16
         TabIndex        =   15
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "答案："
         Height          =   180
         Left            =   360
         TabIndex        =   70
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label Label37 
         Caption         =   "Label37"
         Height          =   255
         Left            =   1080
         TabIndex        =   69
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "问题2："
         Height          =   180
         Left            =   240
         TabIndex        =   68
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "答案："
         Height          =   180
         Left            =   360
         TabIndex        =   67
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label34 
         Caption         =   "Label34"
         Height          =   255
         Left            =   1080
         TabIndex        =   66
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "问题1："
         Height          =   180
         Left            =   360
         TabIndex        =   65
         Top             =   1320
         Width           =   630
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5160
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "要修改密保，请先回答密保问题"
         Height          =   180
         Left            =   600
         TabIndex        =   61
         Top             =   600
         Width           =   2520
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "修改密保"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   4695
      Left            =   0
      TabIndex        =   57
      Top             =   1200
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1800
         Top             =   3480
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1080
         Top             =   3480
      End
      Begin VB.CommandButton Command8 
         Caption         =   "上一步"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "登录"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1080
         MaxLength       =   16
         TabIndex        =   11
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "忘记密码？"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4440
         TabIndex        =   64
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "密码："
         Height          =   180
         Left            =   480
         TabIndex        =   63
         Top             =   2520
         Width           =   540
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "账号："
         Height          =   180
         Left            =   480
         TabIndex        =   62
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "请先登录："
         Height          =   180
         Left            =   360
         TabIndex        =   59
         Top             =   1440
         Width           =   900
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5160
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "此向导将带领您完成密保修改"
         Height          =   180
         Left            =   600
         TabIndex        =   58
         Top             =   480
         Width           =   2340
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "密码修改成功"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   4695
      Left            =   0
      TabIndex        =   49
      Top             =   1200
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command6 
         Caption         =   "完成"
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "请牢记您的新密码并使用新密码登陆"
         Height          =   180
         Left            =   360
         TabIndex        =   56
         Top             =   3000
         Width           =   2880
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5160
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label25 
         Caption         =   "label25"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   55
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "密码："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   600
         TabIndex        =   54
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label Label23 
         Caption         =   "label23"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1440
         TabIndex        =   53
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label22 
         Caption         =   "账号："
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         TabIndex        =   52
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "您的账户信息为："
         Height          =   180
         Left            =   480
         TabIndex        =   51
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "恭喜您，密码修改成功"
         Height          =   180
         Left            =   600
         TabIndex        =   50
         Top             =   600
         Width           =   1800
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "修改密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   4695
      Left            =   0
      TabIndex        =   43
      Top             =   1200
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   960
         Top             =   3720
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   360
         Top             =   3720
      End
      Begin VB.CommandButton Command7 
         Caption         =   "下一步"
         Height          =   255
         Left            =   4200
         TabIndex        =   9
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "16个字符内"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   4440
         TabIndex        =   48
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "再次输入新密码："
         Height          =   180
         Left            =   240
         TabIndex        =   47
         Top             =   2280
         Width           =   1440
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "新密码："
         Height          =   180
         Left            =   480
         TabIndex        =   46
         Top             =   1440
         Width           =   720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5160
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "请输入新密码，并且牢记新密码"
         Height          =   180
         Left            =   600
         TabIndex        =   45
         Top             =   840
         Width           =   2520
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "最后一步：输入新密码"
         Height          =   180
         Left            =   600
         TabIndex        =   44
         Top             =   480
         Width           =   1800
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "修改密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   4695
      Left            =   0
      TabIndex        =   40
      Top             =   1200
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   360
         Top             =   2400
      End
      Begin VB.CommandButton Command3 
         Caption         =   "下一步"
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "上一步"
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   840
         MaxLength       =   16
         TabIndex        =   0
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "请输入要修改密码的账号："
         Height          =   180
         Left            =   360
         TabIndex        =   42
         Top             =   1200
         Width           =   2160
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "此向导将带领您完成密码修改"
         Height          =   180
         Left            =   720
         TabIndex        =   41
         Top             =   480
         Width           =   2340
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "修改密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   4695
      Left            =   0
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command5 
         Caption         =   "下一步"
         Height          =   255
         Left            =   4200
         TabIndex        =   6
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "上一步"
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1080
         MaxLength       =   16
         TabIndex        =   4
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   16
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "忘记密码？"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4320
         MouseIcon       =   "Form6.frx":0000
         TabIndex        =   39
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "换一个问题"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   4320
         MouseIcon       =   "Form6.frx":030A
         TabIndex        =   38
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "答案："
         Height          =   180
         Left            =   480
         TabIndex        =   37
         Top             =   2640
         Width           =   540
      End
      Begin VB.Label Label9 
         Height          =   255
         Left            =   1080
         TabIndex        =   36
         Top             =   2160
         Width           =   3135
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "密保问题："
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "旧密码："
         Height          =   180
         Left            =   360
         TabIndex        =   34
         Top             =   1560
         Width           =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5160
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "要更改密码，请先输入旧密码及回答一个密保问题"
         Height          =   180
         Left            =   720
         TabIndex        =   33
         Top             =   480
         Width           =   3960
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "修改向导"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   4695
      Left            =   0
      TabIndex        =   27
      Top             =   1200
      Width           =   5535
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "2、修改密保问题"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   600
         MouseIcon       =   "Form6.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   2760
         Width           =   2595
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1、修改账户密码"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   600
         MouseIcon       =   "Form6.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   2040
         Width           =   2595
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "请选择要修改的账户信息："
         Height          =   180
         Left            =   240
         TabIndex        =   31
         Top             =   1440
         Width           =   2160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   120
         X2              =   5160
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "密码或密保问题"
         Height          =   180
         Left            =   600
         TabIndex        =   30
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "此向导将带领您更改账户秘密信息，将使用账户的"
         Height          =   180
         Left            =   600
         TabIndex        =   29
         Top             =   480
         Width           =   3960
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      Picture         =   "Form6.frx":0C28
      ScaleHeight     =   1095
      ScaleWidth      =   5535
      TabIndex        =   28
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zhanghao, mima, nicheng, w1, w2, d1, d2 As String
Dim zhanghao2, mima2, nicheng2, w3, d3, w4, d4

Private Sub Command1_Click()
If Dir("c:\百度音乐播放器\用户列表.dll") <> "" And Dir("c:\百度音乐播放器\用户列表.txt") = "" Then
   Name "c:\百度音乐播放器\用户列表.dll" As "c:\百度音乐播放器\用户列表.txt"
End If

If Dir("c:\百度音乐播放器\用户列表.txt") <> "" Then
   Open "c:\百度音乐播放器\用户列表.txt" For Input As #100
   Timer4.Enabled = True
Else
  MsgBox "数据丢失，请重新注册", , "草哥提示"
End If

End Sub

Private Sub Command10_Click()
If Text8.Text <> "" And Text9.Text <> "" Then
      If Text8.Text = d1 Then
             If Text9.Text = d2 Then
                   Frame8.Visible = True
                   Frame7.Visible = False
             Else
                   Text9.Text = ""
                   MsgBox "密保问题回答错误", , "草哥提示"
             End If
       Else
           Text8.Text = ""
           MsgBox "密保问题回答错误", , "草哥提示"
       End If
Else
    MsgBox "请回答密保问题", , "草哥提示"
End If
End Sub

Private Sub Command11_Click()
If Text10.Text <> "" And Text11.Text <> "" And Text12.Text <> "" And Text13.Text <> "" Then
    Close #200
    Open "c:\百度音乐播放器\密保.txt" For Input As #200
    Open "c:\百度音乐播放器\密保2.txt" For Output As #300
    Timer6.Enabled = True
Else
   MsgBox "请输入密保问题和答案", , "草哥提示"
End If
End Sub

Private Sub Command12_Click()
unload Me
End Sub

Private Sub Command2_Click()
Close #78, #68
Frame1.Visible = True
Frame2.Visible = False
End Sub

Private Sub Command3_Click()
If Dir("c:\百度音乐播放器\用户列表.txt") <> "" And Dir("c:\百度音乐播放器\密保.txt") <> "" Then
    
        Open "c:\百度音乐播放器\用户列表.txt" For Input As #78
        
    
    
        Open "c:\百度音乐播放器\密保.txt" For Input As #68
        
    
  Timer1.Enabled = True
End If

End Sub

Private Sub Command4_Click()
Close #68, #78
Frame2.Visible = True
Frame3.Visible = False
End Sub

Private Sub Command5_Click()
If Text2.Text <> "" And Text3.Text <> "" Then
    If Text2.Text = mima Then
          If Label9.Caption = w1 Then
               If Text3.Text = d1 Then
                    Frame4.Visible = True
                    Frame3.Visible = False
               Else
                  MsgBox "密保答案错误", , "草哥提示"
               End If
           ElseIf Label9.Caption = w2 Then
               If Text3.Text = d2 Then
                    Frame4.Visible = True
                    Frame3.Visible = False
               Else
                  MsgBox "密保答案错误", , "草哥提示"
               End If
           End If
    Else
       MsgBox "密码错误", , "草哥提示"
    End If
ElseIf Text2.Text = "" Then
      MsgBox "请输入旧密码", , "草哥提示"
ElseIf Text3.Text = "" Then
        MsgBox "请输入密保答案", , "草哥提示"
End If
End Sub

Private Sub Command6_Click()
unload Me
End Sub

Private Sub Command7_Click()
If Text4.Text = Text5.Text Then
   Close #68, #78
   Open "c:\百度音乐播放器\用户列表.txt" For Input As #78
   Open "c:\百度音乐播放器\密保.txt" For Input As #68
   Open "c:\百度音乐播放器\用户列表2.txt" For Output As #79
   Open "c:\百度音乐播放器\密保2.txt" For Output As #69
    Timer2.Enabled = True
Else
    MsgBox "两次密码输入不一致", , "草哥提示"
End If
End Sub

Private Sub Command9_Click()
Close #100
Close #200
Frame6.Visible = True
Frame7.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close
If Dir("c:\百度音乐播放器\密保.txt") <> "" And Dir("c:\百度音乐播放器\密保.dll") = "" Then
   Name "c:\百度音乐播放器\密保.txt" As "c:\百度音乐播放器\密保.dll"
End If
If Dir("c:\百度音乐播放器\用户列表.txt") <> "" And Dir("c:\百度音乐播放器\用户列表.dll") = "" Then
   Name "c:\百度音乐播放器\用户列表.txt" As "c:\百度音乐播放器\用户列表.dll"
End If
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFF0000
Label5.ForeColor = &HFF0000
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = &HFF0000
Label12.ForeColor = &HFF0000
End Sub

Private Sub Frame6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label32.ForeColor = &HFF0000
End Sub

Private Sub Label11_Click()
If Label9.Caption = w1 Then
   Label9.Caption = w2
ElseIf Label9.Caption = w2 Then
   Label9.Caption = w1
End If
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = vbRed
End Sub

Private Sub Label12_Click()
Form9.Show 1
unload Me
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = vbRed
End Sub

Private Sub Label32_Click()
Form9.Show 1
unload Me
End Sub

Private Sub Label32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label32.ForeColor = vbRed
End Sub

Private Sub Label4_Click()
Frame2.Visible = True
Frame1.Visible = False

If Dir("c:\百度音乐播放器\用户列表.dll") <> "" And Dir("c:\百度音乐播放器\用户列表.txt") = "" Then
    Name "c:\百度音乐播放器\用户列表.dll" As "c:\百度音乐播放器\用户列表.txt"
End If
If Dir("c:\百度音乐播放器\密保.dll") <> "" And Dir("c:\百度音乐播放器\密保.txt") = "" Then
    Name "c:\百度音乐播放器\密保.dll" As "c:\百度音乐播放器\密保.txt"
End If


End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
End Sub

Private Sub Label5_Click()
Frame6.Visible = True
Frame1.Visible = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbRed
End Sub

Private Sub Timer1_Timer()
If Not EOF(78) Then
            Line Input #78, zhanghao
            Line Input #78, mima
            Line Input #78, nicheng
             If Text1.Text = zhanghao Then
                  If Not EOF(68) Then
                        Line Input #68, zhanghao
                        Line Input #68, mima
                        Line Input #68, w1
                        Line Input #68, d1
                        Line Input #68, w2
                        Line Input #68, d2
                        
                        Label9.Caption = w1
                        
                        Frame3.Visible = True
                        Frame2.Visible = False
                        Timer1.Enabled = False
                  Else
                     MsgBox "数据丢失,请重新注册", , "草哥提示"
                     Timer1.Enabled = False
                  End If
             ElseIf Text1.Text <> zhanghao Then
                 If Not EOF(78) Then
                        Line Input #78, zhanghao
                        Line Input #78, mima
                        Line Input #78, nicheng
                 Else
                    Close #68, #78
                    Timer1.Enabled = False
                    MsgBox "该账号未被注册", , "草哥提示"
                 End If
             End If
Else
   Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If Not EOF(78) Then
   Line Input #78, zhanghao2
   Line Input #78, mima2
   Line Input #78, nicheng2
      If zhanghao2 = zhanghao Then
          Print #79, zhanghao2
          Print #79, Text4.Text
          Print #79, nicheng2
      ElseIf zhanghao2 <> zhanghao Then
          Print #79, zhanghao2
          Print #79, mima2
          Print #79, nicheng2
      End If
Else
   Timer2.Enabled = False
   Timer3.Enabled = True
End If
   
End Sub

Private Sub Timer3_Timer()
If Not EOF(68) Then
   Line Input #68, zhanghao2
   Line Input #68, mima2
   Line Input #68, w3
   Line Input #68, d3
   Line Input #68, w4
   Line Input #68, d4
           If zhanghao2 = zhanghao Then
               Print #69, zhanghao2
               Print #69, Text4.Text
               Print #69, w3
               Print #69, d3
               Print #69, w4
               Print #69, d4
           ElseIf zhanghao2 <> zhanghao Then
               Print #69, zhanghao2
               Print #69, mima2
               Print #69, w3
               Print #69, d3
               Print #69, w4
               Print #69, d4
           End If
Else
    Label23.Caption = zhanghao
    Label25.Caption = Text4.Text
    Timer2.Enabled = False
    Timer3.Enabled = False
    Frame5.Visible = True
    Frame4.Visible = False
    Close #78, #79, #68, #69
    If Dir("c:\百度音乐播放器\用户列表.txt") <> "" And Dir("c:\百度音乐播放器\用户列表2.txt") <> "" Then
        Kill "c:\百度音乐播放器\用户列表.txt"
        Name "c:\百度音乐播放器\用户列表2.txt" As "c:\百度音乐播放器\用户列表.txt"
    End If
    If Dir("c:\百度音乐播放器\密保.txt") <> "" And Dir("c:\百度音乐播放器\密保2.txt") <> "" Then
            Kill "c:\百度音乐播放器\密保.txt"
            Name "c:\百度音乐播放器\密保2.txt" As "c:\百度音乐播放器\密保.txt"
    End If
End If
End Sub

Private Sub Timer4_Timer()
If Not EOF(100) Then
   Line Input #100, zhanghao
   Line Input #100, mima
   Line Input #100, nicheng
          If Text6.Text = zhanghao Then
                 If Text7.Text = mima Then
                        If Dir("c:\百度音乐播放器\密保.dll") <> "" And Dir("c:\百度音乐播放器\密保.txt") = "" Then
                           Name "c:\百度音乐播放器\密保.dll" As "c:\百度音乐播放器\密保.txt"
                        End If
                        If Dir("c:\百度音乐播放器\密保.txt") <> "" Then
                            Open "c:\百度音乐播放器\密保.txt" For Input As #200
                        End If
                      Timer5.Enabled = True
                      Close #100
                      Timer4.Enabled = False
                 Else
                    Close #100
                    Timer4.Enabled = False
                    MsgBox "密码错误", , "草哥提示"
                 End If
          ElseIf Text6.Text <> zhanghao Then
              If Not EOF(100) Then
                Line Input #100, zhanghao
                Line Input #100, mima
                Line Input #100, nicheng
              End If
          End If
Else
  Close #100
  Timer4.Enabled = False
  MsgBox "该账号未注册", , "草哥提示"
End If
End Sub

Private Sub Timer5_Timer()
If Not EOF(200) Then
    Line Input #200, zhanghao
    Line Input #200, mima
    Line Input #200, w1
    Line Input #200, d1
    Line Input #200, w2
    Line Input #200, d2
        If zhanghao = Text6.Text Then
            Label34.Caption = w1
            Label37.Caption = w2
            Frame7.Visible = True
            Frame6.Visible = False
            Timer5.Enabled = False
        Else
          If Not EOF(200) Then
            Line Input #200, zhanghao
            Line Input #200, mima
            Line Input #200, w1
            Line Input #200, d1
            Line Input #200, w2
            Line Input #200, d2
          End If
         End If
Else
   Close #200
   Timer5.Enabled = False
   MsgBox "该账号未注册", , "草哥提示"
End If
End Sub

Private Sub Timer6_Timer()
If Not EOF(200) Then
     Line Input #200, zhanghao
     Line Input #200, mima
     Line Input #200, w1
     Line Input #200, d1
     Line Input #200, w2
     Line Input #200, d2
            If zhanghao = Text6.Text Then
                  Print #300, zhanghao
                  Print #300, mima
                  Print #300, Text10.Text
                  Print #300, Text11.Text
                  Print #300, Text12.Text
                  Print #300, Text13.Text
             ElseIf zhanghao <> Text6.Text Then
                  If Not EOF(200) Then
                        Line Input #200, zhanghao
                        Line Input #200, mima
                        Line Input #200, w1
                        Line Input #200, d1
                        Line Input #200, w2
                        Line Input #200, d2
                  End If
             End If
Else
   Label48.Caption = Text10.Text
   Label50.Caption = Text11.Text
   Label52.Caption = Text12.Text
   Label54.Caption = Text13.Text
   Frame9.Visible = True
   Frame8.Visible = False
   Timer6.Enabled = False
   Close #100, #200, #300
   If Dir("c:\百度音乐播放器\密保.txt") <> "" And Dir("c:\百度音乐播放器\密保2.txt") <> "" Then
            Kill "c:\百度音乐播放器\密保.txt"
            Name "c:\百度音乐播放器\密保2.txt" As "c:\百度音乐播放器\密保.txt"
   End If
End If
End Sub
