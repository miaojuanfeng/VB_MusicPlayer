VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "添加网络音乐"
   ClientHeight    =   2400
   ClientLeft      =   10260
   ClientTop       =   4860
   ClientWidth     =   5280
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox Check1 
      Caption         =   "自动命名"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "请输入标准的url地址"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3360
      TabIndex        =   10
      Top             =   240
      Width           =   1710
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "URL："
      Height          =   180
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "歌曲名："
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "歌手名："
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "网络音乐信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1530
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim asd As String
Dim f1, f2, f3, f4 As String
Dim w1, w2, w3, w4, w5, w6, w7, w8 As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Text2.Enabled = False
   Text3.Enabled = False
   Label2.Enabled = False
   Label3.Enabled = False
Else
   Text2.Enabled = True
   Text3.Enabled = True
   Label2.Enabled = True
   Label3.Enabled = True
End If
End Sub

Private Sub Command1_Click()

If Check1.Value = 0 Then
        If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then

            Form1.WindowsMediaPlayer1.URL = Text1.Text
            
            If Form1.Label1.Caption = "" Then
               Form1.Label1.Caption = " " & Text2.Text & " - " & Text3.Text
               
               asd = 1

            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption = "" Then
               Form1.Label2.Caption = " " & Text2.Text & " - " & Text3.Text
               
               asd = 2
                
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption = "" Then
               Form1.Label3.Caption = " " & Text2.Text & " - " & Text3.Text
               
               asd = 3
                
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption = "" Then
               Form1.Label4.Caption = " " & Text2.Text & " - " & Text3.Text
               
               asd = 4
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption = "" Then
               Form1.Label5.Caption = " " & Text2.Text & " - " & Text3.Text
               
               asd = 5
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption = "" Then
               Form1.Label6.Caption = " " & Text2.Text & " - " & Text3.Text
               
               asd = 6
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption <> "" And Form1.Label7.Caption = "" Then
               Form1.Label7.Caption = " " & Text2.Text & " - " & Text3.Text
               
               asd = 7
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption <> "" And Form1.Label7.Caption <> "" And Form1.Label8.Caption = "" Then
               Form1.Label8.Caption = " " & Text2.Text & " - " & Text3.Text
               
               asd = 8
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption <> "" And Form1.Label7.Caption <> "" And Form1.Label8.Caption <> "" And Form1.Label9.Caption = "" Then
               Form1.Label9.Caption = " " & Text2.Text & " - " & Text3.Text
               
               asd = 9
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption <> "" And Form1.Label7.Caption <> "" And Form1.Label8.Caption <> "" And Form1.Label9.Caption <> "" And Form1.Label10.Caption = "" Then
               Form1.Label10.Caption = " " & Text2.Text & " - " & Text3.Text
               
               asd = 10
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption <> "" And Form1.Label7.Caption <> "" And Form1.Label8.Caption <> "" And Form1.Label9.Caption <> "" And Form1.Label10.Caption <> "" Then
                Form1.Label10.Caption = " " & Text2.Text & " - " & Text3.Text
                
                asd = 10
            End If
            
            
            Form1.Label33.Caption = Text1.Text
            Form1.Label34.Caption = asd
            

            
                             Form1.Picture9.Visible = False
                             Form1.Picture6.Visible = True
                             Form1.Timer1.Enabled = True
                             Form1.Timer6.Enabled = True
                             
                               Form1.Text1.Enabled = False
                            If Form1.Text1.Text = "" Then
                               Form1.Text1.Text = "搜索 歌曲、歌手、专辑"
                               Form1.Text1.ForeColor = &H80000011
                            End If
                               
                               
                               Form1.Label23.Visible = False
                               Form1.Label24.Visible = False
                               Form1.Label25.Visible = False
                               Form1.Label26.Visible = False
                               Form1.Label27.Visible = False
                               Form1.Label28.Visible = False
                               Form1.Label29.Visible = False
                               Form1.Label30.Visible = False
                               Form1.Label31.Visible = False
                               Form1.Label32.Visible = False
                               
                               Form1.Picture13.Visible = False
                               Form1.Picture2.Visible = True
                               Form1.Picture15.Visible = False
                               Form1.Picture3.Visible = True
                               Form1.Picture19.Visible = False
           Unload Me
            
    Else
         MsgBox "请输入歌曲信息", , "草哥提示"
    End If
ElseIf Check1.Value = 1 Then
  If Text1.Text <> "" Then
    f1 = Text1.Text
    f2 = Len(f1)
    f3 = Mid(f1, f2, 1)
On Error GoTo errl
    Do While f3 <> "/"
       f2 = f2 - 1
       f3 = Mid(f1, f2, 1)
    Loop
errl:
  Select Case Err.Number
    Case 5
      Text1.Text = ""
      MsgBox "URL输入错误", , "草哥提示"
  End Select
    
    
      If f3 = "/" Then
        f2 = f2 + 1
        f4 = Mid(f1, f2)
        
            Form1.WindowsMediaPlayer1.URL = Text1.Text
        
            If Form1.Label1.Caption = "" Then
               Form1.Label1.Caption = " " & f4
               
               asd = 1

            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption = "" Then
               Form1.Label2.Caption = " " & f4
               
               asd = 2
                
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption = "" Then
               Form1.Label3.Caption = " " & f4
               
               asd = 3
                
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption = "" Then
               Form1.Label4.Caption = " " & f4
               
               asd = 4
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption = "" Then
               Form1.Label5.Caption = " " & f4
               
               asd = 5
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption = "" Then
               Form1.Label6.Caption = " " & f4
               
               asd = 6
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption <> "" And Form1.Label7.Caption = "" Then
               Form1.Label7.Caption = " " & f4
               
               asd = 7
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption <> "" And Form1.Label7.Caption <> "" And Form1.Label8.Caption = "" Then
               Form1.Label8.Caption = " " & f4
               
               asd = 8
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption <> "" And Form1.Label7.Caption <> "" And Form1.Label8.Caption <> "" And Form1.Label9.Caption = "" Then
               Form1.Label9.Caption = " " & f4
               
               asd = 9
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption <> "" And Form1.Label7.Caption <> "" And Form1.Label8.Caption <> "" And Form1.Label9.Caption <> "" And Form1.Label10.Caption = "" Then
               Form1.Label10.Caption = " " & f4
               
               asd = 10
               
            ElseIf Form1.Label1.Caption <> "" And Form1.Label2.Caption <> "" And Form1.Label3.Caption <> "" And Form1.Label4.Caption <> "" And Form1.Label5.Caption <> "" And Form1.Label6.Caption <> "" And Form1.Label7.Caption <> "" And Form1.Label8.Caption <> "" And Form1.Label9.Caption <> "" And Form1.Label10.Caption <> "" Then
                Form1.Label10.Caption = " " & f4
                
                asd = 10
            End If
            
            
             Form1.Label33.Caption = Text1.Text
            Form1.Label34.Caption = asd
            

            
                             Form1.Picture9.Visible = False
                             Form1.Picture6.Visible = True
                             Form1.Timer1.Enabled = True
                             Form1.Timer6.Enabled = True
                             
                               Form1.Text1.Enabled = False
                            If Form1.Text1.Text = "" Then
                               Form1.Text1.Text = "搜索 歌曲、歌手、专辑"
                               Form1.Text1.ForeColor = &H80000011
                            End If
                               
                               
                               Form1.Label23.Visible = False
                               Form1.Label24.Visible = False
                               Form1.Label25.Visible = False
                               Form1.Label26.Visible = False
                               Form1.Label27.Visible = False
                               Form1.Label28.Visible = False
                               Form1.Label29.Visible = False
                               Form1.Label30.Visible = False
                               Form1.Label31.Visible = False
                               Form1.Label32.Visible = False
                               
                               Form1.Picture13.Visible = False
                               Form1.Picture2.Visible = True
                               Form1.Picture15.Visible = False
                               Form1.Picture3.Visible = True
                               Form1.Picture19.Visible = False
      End If
      
      Unload Me
   Else
     MsgBox "请输入URL", , "草哥提示"
   End If
End If
  
  
 
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

