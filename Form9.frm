VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "百度密码找回"
   ClientHeight    =   5940
   ClientLeft      =   6585
   ClientTop       =   2805
   ClientWidth     =   5535
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command3 
      Caption         =   "完成"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   38
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   5400
   End
   Begin VB.Frame Frame3 
      Height          =   3975
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Width           =   5295
      Begin VB.Label Label29 
         Height          =   255
         Left            =   1320
         TabIndex        =   37
         Top             =   3480
         Width           =   3735
      End
      Begin VB.Label Label28 
         Height          =   255
         Left            =   1320
         TabIndex        =   36
         Top             =   3120
         Width           =   3855
      End
      Begin VB.Label Label27 
         Height          =   255
         Left            =   1320
         TabIndex        =   35
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label Label26 
         Height          =   255
         Left            =   1320
         TabIndex        =   34
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Label Label25 
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1320
         TabIndex        =   33
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label24 
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "答案："
         Height          =   180
         Left            =   720
         TabIndex        =   31
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "密保问题二："
         Height          =   180
         Left            =   240
         TabIndex        =   30
         Top             =   3120
         Width           =   1080
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "答案："
         Height          =   180
         Left            =   720
         TabIndex        =   29
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "密保问题一："
         Height          =   180
         Left            =   240
         TabIndex        =   28
         Top             =   2400
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "密码："
         Height          =   180
         Left            =   720
         TabIndex        =   27
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "账号："
         Height          =   180
         Left            =   720
         TabIndex        =   26
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "您的账户信息为："
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   1440
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "--------------------------------------------------------"
         Enabled         =   0   'False
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   5040
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "密保问题回答正确"
         Height          =   180
         Left            =   960
         TabIndex        =   23
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "恭喜您"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一步"
      Height          =   375
      Left            =   3600
      TabIndex        =   20
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   5295
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   840
         MousePointer    =   3  'I-Beam
         TabIndex        =   17
         Top             =   3120
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   840
         MousePointer    =   3  'I-Beam
         TabIndex        =   16
         Top             =   2040
         Width           =   3495
      End
      Begin VB.Label Label13 
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label Label12 
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "答案："
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Top             =   3120
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "答案："
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "问题二："
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "问题一："
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "回答密保问题，回答正确后，便可获知密码"
         Height          =   180
         Left            =   840
         TabIndex        =   11
         Top             =   840
         Width           =   3420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "密保问题"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一步"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   600
         MaxLength       =   16
         MousePointer    =   3  'I-Beam
         TabIndex        =   6
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "--------------------------------------------------------"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   5055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "请输入要找回密码的账号："
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   2160
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "便可轻松找回密码"
         Height          =   180
         Left            =   600
         TabIndex        =   4
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "没关系，只要您记得注册时的密保问题和答案，"
         Height          =   180
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   3780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "忘记密码？"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      Picture         =   "Form9.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   5535
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c1, c2, c3, c4, c5, c6 As String

Private Sub Command1_Click()
If Text1.Text <> "" Then
    If Dir("c:\百度音乐播放器\密保.txt") <> "" Then
      Open "c:\百度音乐播放器\密保.txt" For Input As #1
       If Not EOF(1) Then
          Line Input #1, c1
          Line Input #1, c2
          Line Input #1, c3
          Line Input #1, c4
          Line Input #1, c5
          Line Input #1, c6
          
          Timer1.Enabled = True
          MousePointer = 11
       ElseIf EOF(1) Then
          MsgBox "数据为空", , "草哥提示"
       End If
       
    Else
     MsgBox "请重新注册", , "草哥提示"
    End If
ElseIf Text1.Text = "" Then
   MsgBox "请填写账号", , "草哥提示"
End If
End Sub

Private Sub Command2_Click()
If Text2.Text <> "" And Text3.Text <> "" Then
   If Text2.Text = c4 And Text3.Text = c6 Then
      Frame3.Visible = True
      Frame2.Visible = False
      Command2.Visible = False
      Command3.Visible = True
      

  
      
      Label24.Caption = c1
      Label25.Caption = c2
      Label26.Caption = c3
      Label27.Caption = c4
      Label28.Caption = c5
      Label29.Caption = c6
   Else
      MsgBox "回答错误", , "草哥提示"
   End If
Else
   MsgBox "请输入答案", , "草哥提示"
End If
End Sub

Private Sub Command3_Click()
unload Me
End Sub

Private Sub Form_Load()
 If Dir("c:\百度音乐播放器\密保.dll") <> "" And Dir("c:\百度音乐播放器\密保.txt") = "" Then
    Name "c:\百度音乐播放器\密保.dll" As "c:\百度音乐播放器\密保.txt"
 Else
    MsgBox "数据丢失，请重新注册", , "草哥提示"
 End If
 
 Frame1.Visible = True
 Frame2.Visible = False
 Frame3.Visible = False
 Command1.Visible = True
 Command2.Visible = False
 Command3.Visible = False
 Timer1.Enabled = False
 

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Dir("c:\百度音乐播放器\密保.dll") = "" And Dir("c:\百度音乐播放器\密保.txt") <> "" Then
    Name "c:\百度音乐播放器\密保.txt" As "c:\百度音乐播放器\密保.dll"
End If
End Sub

Private Sub Timer1_Timer()

  If Text1.Text <> c1 Then
    If Not EOF(1) Then
          Line Input #1, c1
          Line Input #1, c2
          Line Input #1, c3
          Line Input #1, c4
          Line Input #1, c5
          Line Input #1, c6
    ElseIf EOF(1) Then
       Close #1
       Timer1.Enabled = False
       MousePointer = 1
       MsgBox "该账号未注册", , "草哥提示"
   End If
 ElseIf Text1.Text = c1 Then
    Close #1
    Label12.Caption = c3
    Label13.Caption = c5

    Frame1.Visible = False
    Frame2.Visible = True
    Command1.Visible = False
    Command2.Visible = True
    

    
    MousePointer = 1
    Timer1.Enabled = False
    
 End If

End Sub
