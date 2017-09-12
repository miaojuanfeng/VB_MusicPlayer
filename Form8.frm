VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "百度账号注册"
   ClientHeight    =   5925
   ClientLeft      =   7170
   ClientTop       =   3045
   ClientWidth     =   5505
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   960
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   5520
   End
   Begin VB.CommandButton Command3 
      Caption         =   "完成"
      Height          =   375
      Left            =   3720
      TabIndex        =   37
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "注册"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "注册完成"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Width           =   5295
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "2.查看“百度音乐播放器”功能介绍"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   480
         MouseIcon       =   "Form8.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   1920
         Width           =   4095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "1.立即登录百度音乐，尽享更多更高品质"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   480
         MouseIcon       =   "Form8.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   1320
         Width           =   4605
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "现在您可以："
         Height          =   180
         Left            =   480
         TabIndex        =   29
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "恭喜您！ 成功注册百度音乐账号"
         Height          =   180
         Left            =   480
         TabIndex        =   28
         Top             =   480
         Width           =   2610
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "注册账号"
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
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5295
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   1320
         MaxLength       =   16
         MousePointer    =   3  'I-Beam
         TabIndex        =   9
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1320
         MaxLength       =   16
         MousePointer    =   3  'I-Beam
         TabIndex        =   8
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1320
         MaxLength       =   16
         MousePointer    =   3  'I-Beam
         TabIndex        =   7
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   1320
         MaxLength       =   16
         MousePointer    =   3  'I-Beam
         TabIndex        =   6
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   1320
         MaxLength       =   6
         MousePointer    =   3  'I-Beam
         TabIndex        =   5
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   16
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   16
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1320
         MaxLength       =   16
         MousePointer    =   3  'I-Beam
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label22 
         Caption         =   "设置密码保护问题和答案，用于密码找回。"
         Height          =   615
         Left            =   3600
         TabIndex        =   36
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "设置小于等于6字符的昵称"
         Height          =   375
         Left            =   3600
         TabIndex        =   35
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   "确认密码一致"
         Height          =   255
         Left            =   3600
         TabIndex        =   34
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "设置小于等于16个字符的密码"
         Height          =   375
         Left            =   3600
         TabIndex        =   33
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "昵称："
         Height          =   255
         Left            =   600
         TabIndex        =   32
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "检查账号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3480
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "————————————————————————————"
         ForeColor       =   &H80000011&
         Height          =   135
         Left            =   120
         TabIndex        =   26
         Top             =   2280
         Width           =   5055
      End
      Begin VB.Label Label11 
         Caption         =   "确认密码："
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "回答："
         Height          =   180
         Left            =   600
         TabIndex        =   24
         Top             =   3600
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "回答："
         Height          =   180
         Left            =   600
         TabIndex        =   23
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label Label8 
         Caption         =   "密保问题2："
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "密保问题1："
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "密码："
         Height          =   180
         Left            =   480
         TabIndex        =   20
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "账号："
         Height          =   180
         Left            =   480
         TabIndex        =   19
         Top             =   480
         Width           =   660
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "注册账号"
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "同意以上条例"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "账号注册："
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
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   5295
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "登录百度后，将享受更多功能与聆听高品质音乐"
         Height          =   180
         Left            =   1200
         TabIndex        =   38
         Top             =   1320
         Width           =   3780
      End
      Begin VB.Label Label4 
         Caption         =   $"Form8.frx":02A4
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   690
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   4935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "注意："
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Top             =   2640
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "  此向导将带领你完成百度音乐的账号注册      "
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   3960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登陆百度，更流畅更高品质"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1320
         TabIndex        =   13
         Top             =   360
         Width           =   3780
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      Picture         =   "Form8.frx":032B
      ScaleHeight     =   1095
      ScaleWidth      =   5535
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b1, b2, b3 As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Command1.Enabled = True
Else
   Command1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Frame1.Visible = False
Frame2.Visible = True
Check1.Visible = False

Command1.Visible = False
Command2.Visible = True


      
End Sub

Private Sub Command2_Click()
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" And Text6.Text <> "" And Text7.Text <> "" And Text8.Text <> "" Then
  If Text2.Text = Text3.Text Then
    If Text5.Text <> Text7.Text Then
    
    
    
    
             MousePointer = 11
        If Dir("c:\百度音乐播放器\用户列表.txt") <> "" Then
             Open "c:\百度音乐播放器\用户列表.txt" For Input As #5
                 If Not EOF(5) Then
                      Line Input #5, b1
                         Line Input #5, b2
                          Line Input #5, b3
                  End If
                Timer2.Enabled = True
        End If










   ElseIf Text5.Text = Text7.Text Then
        Close #5
        MsgBox "密保问题不能一样", , "注册失败"
   End If
 ElseIf Text2.Text <> Text3.Text Then
      Close #5
      MsgBox "密码输入不一致", , "注册失败"
      Text2.Text = ""
      Text3.Text = ""
 End If
 
ElseIf Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Then
  Close #5
  MsgBox "请填写完整资料", , "注册失败"
End If

End Sub

Private Sub Command3_Click()
unload Me
End Sub

Private Sub Form_Load()

Timer1.Enabled = False
Timer2.Enabled = False

Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False


Command1.Enabled = False
Command1.Visible = True
Command2.Visible = False
Command3.Visible = False

Check1.Value = 0

If Dir("c:\百度音乐播放器", vbDirectory) = "" Then
    MkDir ("c:\百度音乐播放器")
End If

If Dir("c:\百度音乐播放器\用户列表.dll") <> "" And Dir("c:\百度音乐播放器\用户列表.txt") = "" Then
   Name "c:\百度音乐播放器\用户列表.dll" As "c:\百度音乐播放器\用户列表.txt"
ElseIf Dir("c:\百度音乐播放器\用户列表.dll") = "" And Dir("c:\百度音乐播放器\用户列表.txt") = "" Then
   Open "c:\百度音乐播放器\用户列表.txt" For Output As #6
   Close #6
End If

                If Dir("c:\百度音乐播放器\密保.dll") <> "" And Dir("c:\百度音乐播放器\密保.txt") = "" Then
                     Name "c:\百度音乐播放器\密保.dll" As "c:\百度音乐播放器\密保.txt"
                ElseIf Dir("c:\百度音乐播放器\密保.txt") = "" And Dir("c:\百度音乐播放器\密保.dll") = "" Then
                     Open "c:\百度音乐播放器\密保.txt" For Output As #7
                     Close #7
                End If
     

End Sub

Private Sub Form_Unload(Cancel As Integer)
               If Dir("c:\百度音乐播放器\用户列表.dll") = "" And Dir("c:\百度音乐播放器\用户列表.txt") <> "" Then
                    Name "c:\百度音乐播放器\用户列表.txt" As "c:\百度音乐播放器\用户列表.dll"
                 End If
                If Dir("c:\百度音乐播放器\密保.dll") = "" And Dir("c:\百度音乐播放器\密保.txt") <> "" Then
                  Name "c:\百度音乐播放器\密保.txt" As "c:\百度音乐播放器\密保.dll"
                 End If
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.ForeColor = &HFF0000
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.ForeColor = &HFF0000
Label17.ForeColor = &HFF0000
End Sub

Private Sub Label13_Click()
If Text1.Text <> "" Then
 If Dir("c:\百度音乐播放器\用户列表.txt") <> "" Then
          Open "c:\百度音乐播放器\用户列表.txt" For Input As #5
        If Not EOF(5) Then
            Line Input #5, b1
            Line Input #5, b2
            Line Input #5, b3
        End If
     MousePointer = 11
     Timer1.Enabled = True
  Else
     MsgBox "用户文件不存在！", , "检查错误"
  End If
Else
   MsgBox "请输入账号", , "检查错误"
End If

 

End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.ForeColor = vbRed
End Sub

Private Sub Label16_Click()

unload Me
Form7.Show 1

End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.ForeColor = vbRed
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.ForeColor = vbRed
End Sub

Private Sub Timer1_Timer()


    If Text1.Text <> b1 Then
       
           If Not EOF(5) Then
            Line Input #5, b1
            Line Input #5, b2
            Line Input #5, b3
           
           ElseIf EOF(5) Then
              Close #5
            MsgBox "该账号未被注册", , "可以使用"
            MousePointer = 1
            Timer1.Enabled = False
           End If
    ElseIf Text1.Text = b1 Then
        Close #5
        MsgBox "该账号已被注册！", , "重新注册"
        Text1.Text = ""
        MousePointer = 1
        Timer1.Enabled = False
    End If
End Sub

Private Sub Timer2_Timer()

  
       If Text1.Text <> b1 Then
         If Not EOF(5) Then
                  Line Input #5, b1
                  Line Input #5, b2
                  Line Input #5, b3
                         
         ElseIf EOF(5) Then
                 Close #5
                 
                  Open "c:\百度音乐播放器\用户列表.txt" For Append As #1
                 Print #1, Text1.Text
                Print #1, Text2.Text
               Print #1, Text4.Text
              Close #1
      
             Open "c:\百度音乐播放器\密保.txt" For Append As #1
                Print #1, Text1.Text
                Print #1, Text2.Text
                Print #1, Text5.Text
                Print #1, Text6.Text
                Print #1, Text7.Text
                Print #1, Text8.Text
               Close #1

                    Frame1.Visible = False
                    Frame2.Visible = False
                    Frame3.Visible = True
                   Command1.Visible = False
                   Command2.Visible = False
                    Command3.Visible = True
                    MousePointer = 1
                    
                    Timer2.Enabled = False
           End If
            
           
           
        
     ElseIf Text1.Text = b1 Then
           Close #5
           MsgBox "该账号已被注册！", , "重新注册"
           Text1.Text = ""
           Text2.Text = ""
           Text3.Text = ""
           Timer2.Enabled = False
           MousePointer = 1
     End If
    
        
    

 
 





   
End Sub
