VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   " 百度°音乐设置"
   ClientHeight    =   6720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8880
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   6720
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   3480
      MaxLength       =   6
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000E&
      Caption         =   "登录时记住密码"
      Height          =   255
      Left            =   3600
      TabIndex        =   25
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H8000000E&
      Caption         =   "自动登录"
      Height          =   255
      Left            =   3600
      TabIndex        =   24
      Top             =   3480
      Width           =   1935
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   8280
      Picture         =   "Form3.frx":B8CD
      ScaleHeight     =   285
      ScaleWidth      =   525
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7745
      Picture         =   "Form3.frx":BD90
      ScaleHeight     =   315
      ScaleWidth      =   1050
      TabIndex        =   21
      Top             =   6340
      Width           =   1050
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7745
      Picture         =   "Form3.frx":C296
      ScaleHeight     =   315
      ScaleWidth      =   1050
      TabIndex        =   20
      Top             =   6340
      Width           =   1050
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7745
      Picture         =   "Form3.frx":C7D0
      ScaleHeight     =   315
      ScaleWidth      =   1050
      TabIndex        =   19
      Top             =   6340
      Width           =   1050
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6600
      Picture         =   "Form3.frx":CCDE
      ScaleHeight     =   315
      ScaleWidth      =   1050
      TabIndex        =   18
      Top             =   6340
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6620
      Picture         =   "Form3.frx":D1F4
      ScaleHeight     =   315
      ScaleWidth      =   1050
      TabIndex        =   17
      Top             =   6340
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5480
      Picture         =   "Form3.frx":D73C
      ScaleHeight     =   315
      ScaleWidth      =   1050
      TabIndex        =   16
      Top             =   6340
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5480
      Picture         =   "Form3.frx":DC4B
      ScaleHeight     =   315
      ScaleWidth      =   1050
      TabIndex        =   15
      Top             =   6340
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "启动时显示功能介绍"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   3120
      Width           =   2175
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   80
      Picture         =   "Form3.frx":E195
      ScaleHeight     =   435
      ScaleWidth      =   2280
      TabIndex        =   12
      Top             =   5790
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.PictureBox Picture6 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   80
      Picture         =   "Form3.frx":E9B2
      ScaleHeight     =   435
      ScaleWidth      =   2280
      TabIndex        =   11
      Top             =   5790
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   435
      Left            =   80
      Picture         =   "Form3.frx":F242
      ScaleHeight     =   435
      ScaleWidth      =   2280
      TabIndex        =   10
      Top             =   5790
      Width           =   2280
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   8280
      Picture         =   "Form3.frx":FAB1
      ScaleHeight     =   270
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   80
      Picture         =   "Form3.frx":FF0E
      ScaleHeight     =   435
      ScaleWidth      =   2280
      TabIndex        =   8
      Top             =   420
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   80
      Picture         =   "Form3.frx":1073D
      ScaleHeight     =   435
      ScaleWidth      =   2280
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "单曲循环"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "随机播放"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "顺序循环播放"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "顺序播放"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.PictureBox Picture15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   7440
      Picture         =   "Form3.frx":10FFB
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   29
      Top             =   840
      Width           =   855
      Begin VB.Image Image1 
         Height          =   855
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "修改密码、密保"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   3480
      MouseIcon       =   "Form3.frx":114C6
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   2160
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "昵称："
      Height          =   180
      Left            =   2880
      TabIndex        =   34
      Top             =   1680
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      Height          =   180
      Left            =   3480
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "账号："
      Height          =   180
      Left            =   2880
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      Height          =   255
      Left            =   960
      TabIndex        =   31
      Top             =   6360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请登录后进行设置"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4560
      TabIndex        =   30
      Top             =   2760
      Width           =   1440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "账号设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   285
      Left            =   2640
      TabIndex        =   28
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "登录设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   285
      Left            =   2640
      TabIndex        =   27
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "更改头像"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   7485
      MouseIcon       =   "Form3.frx":117D0
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "启动设置————————————————"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Top             =   2640
      Width           =   6000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "播放模式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "登录"
      ForeColor       =   &H8000000C&
      Height          =   180
      Left            =   960
      MouseIcon       =   "Form3.frx":11922
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a1, a2, a3, a4 As String
Dim e1, e2, e3, e4 As String
Dim mx, my As String
Dim b1, b2, b3 As String
Dim zhanghao, mima, q1, q2 As String
Dim zhanghao2, mima2, s1, s2, d1, d2 As String
Dim w1, w2, w3, w4 As String

Private Sub Command2_Click()

End Sub

Private Sub Command4_Click()
unload Me
End Sub

Private Sub Check1_Click()
Picture12.Visible = True
End Sub

Private Sub Check2_Click()
Picture12.Visible = True
End Sub

Private Sub Check4_Click()
Picture12.Visible = True
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command5_Click()

    
   
End Sub

Private Sub Form_Load()

If Picture5.Top = 840 And Form1.Label22.Caption = "" Then
   Label1.Visible = True
ElseIf Picture5.Top = 5790 Then
   Label1.Visible = False
End If
If Form1.Label22.Caption <> "" Then
   Label4.Visible = True
   Label4.Caption = "Hello_" & Form1.Label22.Caption
End If

Label3.Visible = False
Label4.Visible = False

Label2.Visible = True
Label13.Visible = True
Option1.Visible = True
Option2.Visible = True
Option3.Visible = True
Option4.Visible = True
Check4.Visible = True

Label5.Visible = False
Label6.Visible = False
Check1.Visible = False
Check2.Visible = False
Image1.Visible = False
Picture15.Visible = False
Label12.Visible = False





Picture9.Visible = False


Label12.ForeColor = &HFF0000


If Dir("c:\百度音乐播放器\系统设置.ini") <> "" And Dir("c:\百度音乐播放器\系统设置.txt") = "" Then
   Name "c:\百度音乐播放器\系统设置.ini" As "c:\百度音乐播放器\系统设置.txt"
ElseIf Dir("c:\百度音乐播放器\系统设置.ini") = "" And Dir("c:\百度音乐播放器\系统设置.txt") = "" Then
   Open "c:\百度音乐播放器\系统设置.txt" For Output As #15
   Close #15
End If



If Dir("c:\百度音乐播放器\系统设置.txt") <> "" Then
   Open "c:\百度音乐播放器\系统设置.txt" For Input As #7
      If Not EOF(7) Then
        Line Input #7, e1
        Line Input #7, e2
        Line Input #7, e3
        Line Input #7, e4
      End If
      If e1 <> "" And e2 <> "" And e3 <> "" And e4 <> "" Then
        Form3.Option1.Value = e1
        Form3.Option2.Value = e2
        Form3.Option3.Value = e3
        Form3.Option4.Value = e4

      End If
   Close #7
ElseIf Dir("c:\百度音乐播放器\系统设置.txt") = "" Then
     Form3.Option1.Value = True
     Form3.Option2.Value = False
     Form3.Option3.Value = False
     Form3.Option4.Value = False
     
     Form10.shunxu.Checked = True
     Form10.shunxuxunhuan.Checked = False
     Form10.shuiji.Checked = False
     Form10.danqu.Checked = False
     

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
     
      
 End If



If Form3.Option1.Value = True Then

   
   Form10.shunxu.Checked = True
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
ElseIf Form3.Option2.Value = True Then
  
   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = True
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
ElseIf Form3.Option3.Value = True Then
   
   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = True
   Form10.danqu.Checked = False
ElseIf Form3.Option4.Value = True Then

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = True
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &H8000000C
Label14.ForeColor = &HFF0000
If 80 < X And X < 2355 And 420 < Y And Y < 855 Then
    Picture1.Visible = True
    Picture2.Visible = False

ElseIf 5460 < X And X < 6520 And 6330 < Y And Y < 6645 Then
    Picture8.Visible = True
    Picture9.Visible = False
ElseIf 6585 < X And X < 7665 And 6330 < Y And Y < 6645 Then
    Picture10.Visible = True
    Picture11.Visible = False
ElseIf 7785 < X And X < 8850 And 6330 < Y And Y < 6645 Then
   If Picture12.Visible = True Then
    Picture13.Visible = True
    Picture14.Visible = False
   End If
Else
    Picture1.Visible = False
    Picture2.Visible = False
    Picture8.Visible = False
    Picture9.Visible = False
    Picture10.Visible = False
    Picture11.Visible = False
    Picture13.Visible = False
    Picture14.Visible = False
End If

If Picture6.Top = 5790 Then
        If 80 < X And X < 2355 And 5775 < Y And Y < 6210 Then
            Picture6.Visible = True
            Picture7.Visible = False
        Else
            Picture6.Visible = False
            Picture7.Visible = False
        End If
ElseIf Picture6.Top = 840 Then
       If 80 < X And X < 2355 And 840 < Y And Y < 1275 Then
            Picture6.Visible = True
            Picture7.Visible = False
        Else
            Picture6.Visible = False
            Picture7.Visible = False
        End If
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
Close

Open "c:\百度音乐播放器\系统设置.txt" For Output As #12
 Print #12, Form3.Option1.Value
 Print #12, Form3.Option2.Value
 Print #12, Form3.Option3.Value
 Print #12, Form3.Option4.Value
Close #12


If Dir("c:\百度音乐播放器\系统设置.txt") <> "" And Dir("c:\百度音乐播放器\系统设置.ini") = "" Then
   Name "c:\百度音乐播放器\系统设置.txt" As "c:\百度音乐播放器\系统设置.ini"
ElseIf Dir("c:\百度音乐播放器\系统设置.txt") = "" And Dir("c:\百度音乐播放器\系统设置.ini") = "" Then
   Open "c:\百度音乐播放器\系统设置.txt" For Output As #18
   Close #18
   Name "c:\百度音乐播放器\系统设置.txt" As "c:\百度音乐播放器\系统设置.ini"
End If

If Dir("c:\百度音乐播放器\用户列表.txt") <> "" And Dir("c:\百度音乐播放器\用户列表2.txt") <> "" Then
   Kill "c:\百度音乐播放器\用户列表.txt"
   Name "c:\百度音乐播放器\用户列表2.txt" As "c:\百度音乐播放器\用户列表.txt"
End If
If Dir("c:\百度音乐播放器\密保.txt") <> "" And Dir("c:\百度音乐播放器\密保2.txt") <> "" Then
   Kill "c:\百度音乐播放器\密保.txt"
   Name "c:\百度音乐播放器\密保2.txt" As "c:\百度音乐播放器\密保.txt"
End If

If Dir("c:\百度音乐播放器\密保.dll") = "" And Dir("c:\百度音乐播放器\密保.txt") <> "" Then
    Name "c:\百度音乐播放器\密保.txt" As "c:\百度音乐播放器\密保.dll"
End If

If Dir("c:\百度音乐播放器\auto.txt") <> "" Then
    Open "c:\百度音乐播放器\auto.txt" For Output As #4
               Print #4, Form7.Combo1.Text
              If Check1.Value = 1 Then
               Print #4, Form7.Text1.Text
              Else
               Print #4,
              End If
               Print #4, Check1.Value
               Print #4, Check2.Value
            Close #4
 End If

Form1.Label20.ForeColor = &H4040&
Form1.Label21.ForeColor = &H808000
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = &HFF0000
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


If Button = 1 Then
   mx = X
   my = Y
End If
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Form3.Left = Form3.Left + X - mx
   Form3.Top = Form3.Top + Y - my
End If

If 8265 < X And X < 8760 And 15 < Y And Y < 285 Then
    Picture3.Visible = True
    Picture4.Visible = False
Else
    Picture3.Visible = False
    Picture4.Visible = False
End If

End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   If Form1.Label55.Caption = "1" Then
      Form1.Label55.Caption = "2"
   ElseIf Form1.Label55.Caption = "2" Then
      Form1.Label55.Caption = "1"
   End If
End If
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = vbRed
End Sub

Private Sub Label14_Click()
Form6.Show 1
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.ForeColor = vbRed

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Form7.Show 1
End If
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbRed
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &H8000000C
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Picture1.Visible = False
   Picture2.Visible = True
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture5.Top = 5790
Picture6.Top = 5790
Picture7.Top = 5790

Label2.Visible = True
Label13.Visible = True
Option1.Visible = True
Option2.Visible = True
Option3.Visible = True
Option4.Visible = True
Check4.Visible = True

Label5.Visible = False
Label6.Visible = False
Check1.Visible = False
Check2.Visible = False
Image1.Visible = False
Picture15.Visible = False
Label12.Visible = False


Picture1.Visible = True
Picture2.Visible = False

Label1.Visible = False
Label3.Visible = False
Label4.Visible = False

Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label14.Visible = False
Text1.Visible = False

End Sub

Private Sub Picture10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Picture10.Visible = False
   Picture11.Visible = True
End If
End Sub

Private Sub Picture10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Close
unload Me
End Sub

Private Sub Picture13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Picture13.Visible = False
   Picture14.Visible = True
End If
End Sub

Private Sub Picture13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Form3.Option1.Value = True Then
   Form1.Timer2.Enabled = True

   
   Form10.shunxu.Checked = True
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
ElseIf Form3.Option2.Value = True Then
   Form1.Timer2.Enabled = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = True
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
ElseIf Form3.Option3.Value = True Then
   Form1.Timer2.Enabled = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = True
   Form10.danqu.Checked = False
ElseIf Form3.Option4.Value = True Then
   Form1.Timer2.Enabled = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = True
End If

Open "c:\百度音乐播放器\系统设置.txt" For Output As #13
 Print #13, Form3.Option1.Value
 Print #13, Form3.Option2.Value
 Print #13, Form3.Option3.Value
 Print #13, Form3.Option4.Value
Close #13



End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Picture4.Visible = True
    Picture3.Visible = False
End If
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 
   unload Form3
End If
End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Picture6.Visible = False
   Picture7.Visible = True
End If
End Sub

Private Sub Picture6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Picture6.Top = 5790 Then
   Picture5.Top = 840
   Picture6.Top = 840
   Picture7.Top = 840
End If

    If Form1.Label22.Caption = "" Then
       Label3.Visible = True
       Label4.Visible = False
       Label1.Visible = True
       
       Label2.Visible = False
        Label13.Visible = False
        Option1.Visible = False
        Option2.Visible = False
        Option3.Visible = False
        Option4.Visible = False
        Check4.Visible = False
        
        Label8.Visible = False
        Label9.Visible = False
        Label10.Visible = False
        Text1.Visible = False
    Else
       Label1.Visible = False
       Label3.Visible = False
       Label4.Visible = True
    

        Label2.Visible = False
        Label13.Visible = False
        Option1.Visible = False
        Option2.Visible = False
        Option3.Visible = False
        Option4.Visible = False
        Check4.Visible = False
        
        
        Label5.Visible = True
        Label6.Visible = True
        Check1.Visible = True
        Check2.Visible = True
        Image1.Visible = True
        Picture15.Visible = True
        Label12.Visible = True
        
        Label8.Visible = True
        Label9.Visible = True
        Label10.Visible = True
        Label14.Visible = True
        Text1.Visible = True
        Label9.Caption = Form7.Combo1.Text
        Text1.Text = Form1.Label22.Caption
     End If
Picture6.Visible = True
Picture7.Visible = False
End Sub

Private Sub Picture8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Picture8.Visible = False
   Picture9.Visible = True
End If
End Sub

Private Sub Picture8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Form3.Option1.Value = True Then
   Form1.Timer2.Enabled = True

   
   Form10.shunxu.Checked = True
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = True
   Form10.shunxuxunhuan2.Checked = False
   Form10.shuiji2.Checked = False
   Form10.danqu2.Checked = False
ElseIf Form3.Option2.Value = True Then
   Form1.Timer2.Enabled = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = True
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = False
   Form10.shunxuxunhuan2.Checked = True
   Form10.shuiji2.Checked = False
   Form10.danqu2.Checked = False
ElseIf Form3.Option3.Value = True Then
   Form1.Timer2.Enabled = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = True
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = False
   Form10.shunxuxunhuan2.Checked = False
   Form10.shuiji2.Checked = True
   Form10.danqu2.Checked = False
ElseIf Form3.Option4.Value = True Then
   Form1.Timer2.Enabled = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = True
   
    Form10.shunxu2.Checked = False
   Form10.shunxuxunhuan2.Checked = False
   Form10.shuiji2.Checked = False
   Form10.danqu2.Checked = True
End If

Open "c:\百度音乐播放器\系统设置.txt" For Output As #13
 Print #13, Form3.Option1.Value
 Print #13, Form3.Option2.Value
 Print #13, Form3.Option3.Value
 Print #13, Form3.Option4.Value
Close #13



unload Me
End Sub

