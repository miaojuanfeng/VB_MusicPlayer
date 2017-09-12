VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   0  'None
   Caption         =   "Mini模式"
   ClientHeight    =   420
   ClientLeft      =   6540
   ClientTop       =   5385
   ClientWidth     =   5460
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   420
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture16 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      ScaleHeight     =   255
      ScaleWidth      =   1545
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   1545
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         Height          =   180
         Left            =   1230
         TabIndex        =   20
         Top             =   45
         Width           =   300
      End
      Begin VB.Label Label2 
         Height          =   165
         Left            =   150
         TabIndex        =   18
         Top             =   45
         Width           =   90
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000001&
         Height          =   35
         Left            =   160
         TabIndex        =   19
         Top             =   100
         Width           =   15
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Height          =   60
         Left            =   150
         TabIndex        =   17
         Top             =   90
         Width           =   1000
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3720
      Top             =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   4120
      Picture         =   "Form11.frx":4097
      ScaleHeight     =   420
      ScaleWidth      =   1335
      TabIndex        =   2
      Top             =   0
      Width           =   1335
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   970
         Picture         =   "Form11.frx":6CFF
         ScaleHeight     =   300
         ScaleWidth      =   285
         TabIndex        =   11
         Top             =   60
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   690
         Picture         =   "Form11.frx":70F5
         ScaleHeight     =   300
         ScaleWidth      =   285
         TabIndex        =   10
         Top             =   60
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   410
         Picture         =   "Form11.frx":7463
         ScaleHeight     =   300
         ScaleWidth      =   285
         TabIndex        =   9
         Top             =   60
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   120
         Picture         =   "Form11.frx":77CB
         ScaleHeight     =   300
         ScaleWidth      =   285
         TabIndex        =   8
         Top             =   60
         Visible         =   0   'False
         Width           =   285
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      Picture         =   "Form11.frx":7B42
      ScaleHeight     =   420
      ScaleWidth      =   1950
      TabIndex        =   1
      Top             =   0
      Width           =   1950
      Begin VB.PictureBox Picture15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1370
         Picture         =   "Form11.frx":B0FF
         ScaleHeight     =   300
         ScaleWidth      =   255
         TabIndex        =   15
         ToolTipText     =   "取消静音"
         Top             =   60
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1360
         Picture         =   "Form11.frx":B49A
         ScaleHeight     =   300
         ScaleWidth      =   270
         TabIndex        =   14
         Top             =   60
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   430
         Picture         =   "Form11.frx":B7EA
         ScaleHeight     =   300
         ScaleWidth      =   315
         TabIndex        =   13
         ToolTipText     =   "暂停"
         Top             =   60
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   450
         Picture         =   "Form11.frx":BB50
         ScaleHeight     =   300
         ScaleWidth      =   255
         TabIndex        =   12
         Top             =   60
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1640
         Picture         =   "Form11.frx":BE80
         ScaleHeight     =   300
         ScaleWidth      =   180
         TabIndex        =   7
         ToolTipText     =   "音量"
         Top             =   60
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1370
         Picture         =   "Form11.frx":C1B9
         ScaleHeight     =   300
         ScaleWidth      =   255
         TabIndex        =   6
         ToolTipText     =   "静音"
         Top             =   60
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1050
         Picture         =   "Form11.frx":C51C
         ScaleHeight     =   300
         ScaleWidth      =   315
         TabIndex        =   5
         ToolTipText     =   "下一首"
         Top             =   60
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   740
         Picture         =   "Form11.frx":C88E
         ScaleHeight     =   300
         ScaleWidth      =   315
         TabIndex        =   4
         ToolTipText     =   "上一首"
         Top             =   60
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   430
         Picture         =   "Form11.frx":CBFF
         ScaleHeight     =   300
         ScaleWidth      =   315
         TabIndex        =   3
         ToolTipText     =   "播放"
         Top             =   60
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   405
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   180
      Left            =   1960
      TabIndex        =   0
      Top             =   105
      Width           =   105
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ab, shu As Integer
Dim mouse_x, mouse_y As String
Dim acs As Integer
Dim mo_x, mo_y As String



Private Sub Form_Load()
Label1.Caption = "百度音乐_音乐你的生活"
Label1.Left = 1960
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    mouse_x = X
    mouse_y = Y
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 1 Then
            If 150 <= Label2.Left And Label2.Left <= 1100 Then
               Label2.Left = Label2.Left + X - mouse_x
            ElseIf Label2.Left < 150 Then
              
               Label2.Left = 150
               
            ElseIf Label2.Left > 1100 Then
               Label2.Left = 1100
               
            End If
            
            
            If Label4.Width > 1030 Then
               Label4.Width = 1030
            ElseIf Label4.Width < 15 Then
               Label4.Width = 15
            ElseIf 15 <= Label4.Width And Label4.Width <= 1030 Then
               Label4.Width = Label2.Left - 100
            End If
                        
                        If Label2.Left = 150 Then
                           Label5.Caption = "0"
                           Form1.WindowsMediaPlayer1.settings.volume = 0
                        ElseIf 150 < Label2.Left And Label2.Left <= 250 Then
                           Label5.Caption = "10"
                           Form1.WindowsMediaPlayer1.settings.volume = 10
                        ElseIf 250 < Label2.Left And Label2.Left <= 350 Then
                           Label5.Caption = "20"
                           Form1.WindowsMediaPlayer1.settings.volume = 20
                        ElseIf 350 < Label2.Left And Label2.Left <= 450 Then
                           Label5.Caption = "30"
                           Form1.WindowsMediaPlayer1.settings.volume = 30
                        ElseIf 450 < Label2.Left And Label2.Left <= 550 Then
                           Label5.Caption = "40"
                           Form1.WindowsMediaPlayer1.settings.volume = 40
                        ElseIf 550 < Label2.Left And Label2.Left <= 650 Then
                           Label5.Caption = "50"
                           Form1.WindowsMediaPlayer1.settings.volume = 50
                        ElseIf 650 < Label2.Left And Label2.Left <= 750 Then
                           Label5.Caption = "60"
                           Form1.WindowsMediaPlayer1.settings.volume = 60
                        ElseIf 750 < Label2.Left And Label2.Left <= 850 Then
                           Label5.Caption = "70"
                           Form1.WindowsMediaPlayer1.settings.volume = 70
                        ElseIf 850 < Label2.Left And Label2.Left <= 950 Then
                           Label5.Caption = "80"
                           Form1.WindowsMediaPlayer1.settings.volume = 80
                        ElseIf 950 < Label2.Left And Label2.Left <= 1050 Then
                           Label5.Caption = "90"
                           Form1.WindowsMediaPlayer1.settings.volume = 90
                        ElseIf Label2.Left = 1100 Then
                           Label5.Caption = "100"
                           Form1.WindowsMediaPlayer1.settings.volume = 100
                        End If
     End If
     
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   mo_x = X
   mo_y = Y
End If
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Form11.Left = Form11.Left + X - mo_x
   Form11.Top = Form11.Top + Y - mo_y
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If 420 < X And X < 735 And 75 < Y And Y < 360 Then
       If Picture12.Visible = False Then
        Picture3.Visible = True
       Else
        Picture13.Visible = True
       End If
        Picture4.Visible = False
        Picture5.Visible = False
        Picture6.Visible = False
        Picture7.Visible = False
        
        Picture8.Visible = False
        Picture9.Visible = False
        Picture10.Visible = False
        Picture11.Visible = False
        
        Picture15.Visible = False
    ElseIf 738 < X And X < 1050 And 75 < Y And Y < 360 Then
        Picture3.Visible = False
        Picture4.Visible = True
        Picture5.Visible = False
        Picture6.Visible = False
        Picture7.Visible = False
        
        Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = False
     
     Picture13.Visible = False
     Picture15.Visible = False
    ElseIf 1053 < X And X < 1350 And 75 < Y And Y < 360 Then
        Picture3.Visible = False
        Picture4.Visible = False
        Picture5.Visible = True
        Picture6.Visible = False
        Picture7.Visible = False
        
        Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = False
     
     Picture13.Visible = False
     Picture15.Visible = False
    ElseIf 1353 < X And X < 1605 And 75 < Y And Y < 360 Then
        Picture3.Visible = False
        Picture4.Visible = False
        Picture5.Visible = False
       If Form1.WindowsMediaPlayer1.settings.mute = False Then
        Picture6.Visible = True
       Else
        Picture15.Visible = True
       End If
        Picture7.Visible = False
        
        Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = False
     
     Picture13.Visible = False
    ElseIf 1608 < X And X < 1815 And 75 < Y And Y < 360 Then
        Picture3.Visible = False
        Picture4.Visible = False
        Picture5.Visible = False
        Picture6.Visible = False
        Picture7.Visible = True
        
        Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = False
     
     Picture13.Visible = False
     Picture15.Visible = False
    Else
        Picture3.Visible = False
        Picture4.Visible = False
        Picture5.Visible = False
        Picture6.Visible = False
        Picture7.Visible = False
        Picture13.Visible = False
        Picture15.Visible = False
    End If

End Sub

Private Sub Picture10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Picture10.Visible = False
End If
End Sub

Private Sub Picture10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Form1.Show
   Form11.Hide
End If
End Sub

Private Sub Picture11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Picture11.Visible = False
End If
End Sub

Private Sub Picture11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

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

Private Sub Picture12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture13.Visible = True
End Sub

Private Sub Picture13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Picture13.Visible = False
End If
End Sub

Private Sub Picture13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Picture4.Visible = False
        Picture5.Visible = False
        Picture6.Visible = False
        Picture7.Visible = False
        
        Picture8.Visible = False
        Picture9.Visible = False
        Picture10.Visible = False
        Picture11.Visible = False
        
        Picture15.Visible = False
End Sub

Private Sub Picture13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.WindowsMediaPlayer1.Controls.pause
Picture12.Visible = False
Picture13.Visible = False
End Sub

Private Sub Picture14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture15.Visible = True
End Sub

Private Sub Picture15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  Picture15.Visible = False
End If
End Sub

Private Sub Picture15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = False
        Picture4.Visible = False
        Picture5.Visible = False
        Picture7.Visible = False
        
        Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = False
     
     Picture13.Visible = False
End Sub

Private Sub Picture15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  Form1.WindowsMediaPlayer1.settings.mute = False
  Picture15.Visible = False
  Picture14.Visible = False
  Picture6.Visible = False
End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox X & " " & Y
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If 105 < X And X < 405 And 75 < Y And Y < 360 Then
     Picture8.Visible = True
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = False
  ElseIf 408 < X And X < 690 And 75 < Y And Y < 360 Then
     Picture8.Visible = False
     Picture9.Visible = True
     Picture10.Visible = False
     Picture11.Visible = False
  ElseIf 693 < X And X < 960 And 75 < Y And Y < 360 Then
     Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = True
     Picture11.Visible = False
  ElseIf 963 < X And X < 1260 And 75 < Y And Y < 360 Then
     Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = True
  Else
     Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = False
  End If
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Picture3.Visible = False
End If
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Picture4.Visible = False
        Picture5.Visible = False
        Picture6.Visible = False
        Picture7.Visible = False
        
        Picture8.Visible = False
        Picture9.Visible = False
        Picture10.Visible = False
        Picture11.Visible = False
        
        Picture15.Visible = False
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Form1.WindowsMediaPlayer1.Controls.play
   Picture12.Visible = True
End If
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
       Picture4.Visible = False
    End If

End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = False
        
        Picture5.Visible = False
        Picture6.Visible = False
        Picture7.Visible = False
        
        Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = False
     
     Picture13.Visible = False
     Picture15.Visible = False
End Sub

Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
       If Form1.Label56.Caption = "1" Then
          Form1.Label56.Caption = "2"
       ElseIf Form1.Label56.Caption = "2" Then
          Form1.Label56.Caption = "1"
       End If

   End If
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
       Picture5.Visible = False
    End If

End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = False
        Picture4.Visible = False
        
        Picture6.Visible = False
        Picture7.Visible = False
        
        Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = False
     
     Picture13.Visible = False
     Picture15.Visible = False
End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
       If Form1.Label57.Caption = "1" Then
          Form1.Label57.Caption = "2"
       ElseIf Form1.Label57.Caption = "2" Then
          Form1.Label57.Caption = "1"
       End If
    End If

End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Picture6.Visible = False
End If
End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = False
        Picture4.Visible = False
        Picture5.Visible = False
        Picture7.Visible = False
        
        Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = False
     
     Picture13.Visible = False
End Sub

Private Sub Picture6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Form1.WindowsMediaPlayer1.settings.mute = True
   Picture6.Visible = False
   Picture14.Visible = True
End If
End Sub

Private Sub Picture7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Picture3.Visible = False
        Picture4.Visible = False
        Picture5.Visible = False
        Picture6.Visible = False
        
        
        Picture8.Visible = False
     Picture9.Visible = False
     Picture10.Visible = False
     Picture11.Visible = False
     
     Picture13.Visible = False
     Picture15.Visible = False
End Sub

Private Sub Picture9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  Picture9.Visible = False
End If
End Sub

Private Sub Picture9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Form11.WindowState = 1
End If
End Sub

Private Sub Timer1_Timer()

    Label1.Caption = Form1.Label16.Caption
    If Label1.Caption = "百度音乐_音乐你的生活" Then
       Label1.Left = 1960
    End If
    
    
    
    
    
 If Form1.WindowsMediaPlayer1.playState = 3 Then
    If Label1.Width > 2200 And Label1.Caption <> "百度音乐_音乐你的生活" Then
          ab = 4120 - Label1.Width
       
        If Label1.Left >= 1960 Then
          shu = 0
        ElseIf Label1.Left <= ab Then
          shu = 1
        End If
        
       If shu = 0 Then
            Label1.Left = Label1.Left - 10
       ElseIf shu = 1 Then
            Label1.Left = Label1.Left + 10
       End If
    End If
 ElseIf Form1.WindowsMediaPlayer1.playState = 1 Or 2 Then
    Label1.Left = 1960
 End If


If Form1.WindowsMediaPlayer1.playState = 3 Then
   Picture12.Visible = True
   Picture3.Visible = False
ElseIf Form1.WindowsMediaPlayer1.playState = 1 Or 2 Then
   Picture12.Visible = False
End If

End Sub
