VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��½baidu����"
   ClientHeight    =   3975
   ClientLeft      =   7110
   ClientTop       =   4080
   ClientWidth     =   5535
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox Text1 
      Height          =   305
      IMEMode         =   3  'DISABLE
      Left            =   1930
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1720
      Width           =   2490
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form7.frx":038A
      Left            =   1930
      List            =   "Form7.frx":038C
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Top             =   1220
      Width           =   2520
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "Form7.frx":038E
      ScaleHeight     =   855
      ScaleWidth      =   3375
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3120
      Width           =   3375
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   960
         Top             =   480
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "< ��¼ >"
         Default         =   -1  'True
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3240
      Picture         =   "Form7.frx":354A
      ScaleHeight     =   855
      ScaleWidth      =   2295
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2295
      Begin VB.CommandButton Command2 
         Caption         =   "ȡ��"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "�Զ���¼"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��ס����"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   895
      Left            =   1080
      Picture         =   "Form7.frx":64C8
      ScaleHeight     =   900
      ScaleWidth      =   3390
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3395
      Begin VB.Label Label2 
         Caption         =   "  ���룺"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   590
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "  �˺ţ�"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   240
      Picture         =   "Form7.frx":AF03
      ScaleHeight     =   780
      ScaleWidth      =   780
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   780
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      Picture         =   "Form7.frx":DD31
      ScaleHeight     =   1095
      ScaleWidth      =   5535
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   -120
      Width           =   5535
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�һ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   4680
      MouseIcon       =   "Form7.frx":162A2
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ע���˺�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   4680
      MouseIcon       =   "Form7.frx":163F4
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "��¼�ٶȣ�����������Ʒ��"
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   2760
      Width           =   2295
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zhanghao, mima, nicheng  As String
Dim ad, autoz, autom, ch1, ch2 As String

Private Sub Command1_Click()

If Dir("c:\�ٶ����ֲ�����\�û��б�.txt") <> "" Then
   Open "c:\�ٶ����ֲ�����\�û��б�.txt" For Input As #1
   If Not EOF(1) Then
    Line Input #1, zhanghao
    Line Input #1, mima
    Line Input #1, nicheng
        If zhanghao = "" And mima = "" And nicheng = "" Then
          Line Input #1, zhanghao
          Line Input #1, mima
          Line Input #1, nicheng
        End If
   End If
  
   Timer2.Enabled = True
   
   Label3.Caption = "���ڵ�½......"
   Command1.Enabled = False
   
ElseIf Dir("c:\�ٶ����ֲ�����\�û��б�.txt") = "" Then
   MsgBox "������ע���¼", , "��¼ʧ��"
   Timer2.Enabled = False
End If

End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Form_Load()


Timer2.Enabled = False
'�Զ���¼
If Dir("c:\�ٶ����ֲ�����\auto.dll") <> "" And Dir("c:\�ٶ����ֲ�����\auto.txt") = "" Then
   Name "c:\�ٶ����ֲ�����\auto.dll" As "c:\�ٶ����ֲ�����\auto.txt"
End If
'�û��б�
If Dir("c:\�ٶ����ֲ�����\�û��б�.dll") <> "" And Dir("c:\�ٶ����ֲ�����\�û��б�.txt") = "" Then
   Name "c:\�ٶ����ֲ�����\�û��б�.dll" As "c:\�ٶ����ֲ�����\�û��б�.txt"
ElseIf Dir("c:\�ٶ����ֲ�����\�û��б�.dll") = "" And Dir("c:\�ٶ����ֲ�����\�û��б�.txt") = "" Then
   MsgBox "���ݿ���Ϣ��ʧ��������ע���¼", , "�ݸ���ʾ"
End If
'combo1�б�
If Dir("c:\�ٶ����ֲ�����\combo1.dll") <> "" And Dir("c:\�ٶ����ֲ�����\combo1.txt") = "" Then
  Name "c:\�ٶ����ֲ�����\combo1.dll" As "c:\�ٶ����ֲ�����\combo1.txt"
End If
'��������������������������������������������������������������������������������������������������������������������
  If Dir("c:\�ٶ����ֲ�����\combo1.txt") <> "" Then
    Open "c:\�ٶ����ֲ�����\combo1.txt" For Input As #3
       Do While Not EOF(3)
         Line Input #3, ad
         Combo1.AddItem ad
       Loop
    Close #3
  End If
                
 If Dir("c:\�ٶ����ֲ�����\auto.txt") <> "" Then
    Open "c:\�ٶ����ֲ�����\auto.txt" For Input As #4
    If Not EOF(4) Then
        Line Input #4, autoz
        Combo1.Text = autoz
        Line Input #4, autom
        Text1.Text = autom
        Line Input #4, ch1
        Check1.Value = ch1
        Form3.Check1.Value = ch1
        Line Input #4, ch2
        Check2.Value = ch2
        Form3.Check2.Value = ch2
      End If
    Close #4
 End If
 
 If Dir("c:\�ٶ����ֲ�����\auto.txt") <> "" Then
    If Check2.Value = 1 Then
      If Dir("c:\�ٶ����ֲ�����\�û��б�.txt") <> "" Then
          Open "c:\�ٶ����ֲ�����\�û��б�.txt" For Input As #1
            Line Input #1, zhanghao
            Line Input #1, mima
            Line Input #1, nicheng
        If zhanghao = "" And mima = "" And nicheng = "" Then
          Line Input #1, zhanghao
          Line Input #1, mima
          Line Input #1, nicheng
        End If
     Timer2.Enabled = True
   
     Label3.Caption = "���ڵ�¼......"
      Command1.Enabled = False
   
    ElseIf Dir("c:\�ٶ����ֲ�����\�û��б�.txt") = "" Then
     MsgBox "������ע���¼", , "��¼ʧ��"
     Timer2.Enabled = False
    End If
  End If
Else
    Timer2.Enabled = False
    Command1.Enabled = True
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MousePointer = 1
Label4.ForeColor = &HC00000
Label5.ForeColor = &HC00000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Close #1
If Dir("c:\�ٶ����ֲ�����\�û��б�.txt") <> "" And Dir("c:\�ٶ����ֲ�����\�û��б�.dll") = "" Then
   Name "c:\�ٶ����ֲ�����\�û��б�.txt" As "c:\�ٶ����ֲ�����\�û��б�.dll"
End If

If Dir("c:\�ٶ����ֲ�����\auto.txt") <> "" And Dir("c:\�ٶ����ֲ�����\auto.dll") = "" Then
    Name "c:\�ٶ����ֲ�����\auto.txt" As "c:\�ٶ����ֲ�����\auto.dll"
End If

If Dir("c:\�ٶ����ֲ�����\combo1.txt") <> "" And Dir("c:\�ٶ����ֲ�����\combo1.dll") = "" Then
    Name "c:\�ٶ����ֲ�����\combo1.txt" As "c:\�ٶ����ֲ�����\combo1.dll"
End If


Form1.Label20.ForeColor = &H4040&
Form1.Label21.ForeColor = &H808000
End Sub

Private Sub Label4_Click()
Close #1

Unload Form7
Form8.Show 1


End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
End Sub

Private Sub Label5_Click()
Close #1

Unload Me
Form9.Show 1
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbRed
End Sub

Private Sub Timer1_Timer()
If Form1.Label22.Caption = "" Then
                    Label3.Caption = "��¼�ٶȣ�����������Ʒ��"
                    Command1.Enabled = True
               Timer2.Enabled = False
            '�Զ���¼
            If Dir("c:\�ٶ����ֲ�����\auto.dll") <> "" And Dir("c:\�ٶ����ֲ�����\auto.txt") = "" Then
               Name "c:\�ٶ����ֲ�����\auto.dll" As "c:\�ٶ����ֲ�����\auto.txt"
            End If
            '�û��б�
            If Dir("c:\�ٶ����ֲ�����\�û��б�.dll") <> "" And Dir("c:\�ٶ����ֲ�����\�û��б�.txt") = "" Then
               Name "c:\�ٶ����ֲ�����\�û��б�.dll" As "c:\�ٶ����ֲ�����\�û��б�.txt"
            ElseIf Dir("c:\�ٶ����ֲ�����\�û��б�.dll") = "" And Dir("c:\�ٶ����ֲ�����\�û��б�.txt") = "" Then
               MsgBox "���ݿ���Ϣ��ʧ��������ע���¼", , "�ݸ���ʾ"
            End If
            'combo1�б�
            If Dir("c:\�ٶ����ֲ�����\combo1.dll") <> "" And Dir("c:\�ٶ����ֲ�����\combo1.txt") = "" Then
              Name "c:\�ٶ����ֲ�����\combo1.dll" As "c:\�ٶ����ֲ�����\combo1.txt"
            End If
            '��������������������������������������������������������������������������������������������������������������������
              If Dir("c:\�ٶ����ֲ�����\combo1.txt") <> "" Then
                Open "c:\�ٶ����ֲ�����\combo1.txt" For Input As #3
                   Do While Not EOF(3)
                     Line Input #3, ad
                     Combo1.AddItem ad
                   Loop
                Close #3
              End If
                            
             If Dir("c:\�ٶ����ֲ�����\auto.txt") <> "" Then
                Open "c:\�ٶ����ֲ�����\auto.txt" For Input As #4
                If Not EOF(4) Then
                    Line Input #4, autoz
                    Combo1.Text = autoz
                    Line Input #4, autom
                    Text1.Text = autom
                    Line Input #4, ch1
                    Check1.Value = ch1
                    Form3.Check1.Value = ch1
                    Line Input #4, ch2
                    Check2.Value = ch2
                    Form3.Check2.Value = ch2
                  End If
                Close #4
             End If
             
   
     Timer1.Enabled = False
Else
     Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()

If Dir("c:\�ٶ����ֲ�����\�û��б�.txt") <> "" Then
    
     If Combo1.Text <> zhanghao Then
           If Not EOF(1) Then
            Line Input #1, zhanghao
            Line Input #1, mima
            Line Input #1, nicheng
                
           ElseIf EOF(1) Then
              Close #1
            MsgBox "���˺�δ��ע��", , "��¼ʧ��"
            Timer2.Enabled = False
            Command1.Enabled = True
            Label3.Caption = "��¼�ٶȣ�����������Ʒ��"
           End If
             
      
   ElseIf Combo1.Text = zhanghao Then
      If Text1.Text = mima Then
         Close #1
       
         Combo1.AddItem Combo1
         Open "c:\�ٶ����ֲ�����\combo1.txt" For Append As #3
            Print #3, Combo1.Text
         Close #3
         Form1.Picture19.Visible = False
         Form1.Label20.Visible = True
         Form1.Label48.Visible = False
         Form1.Label20.Caption = "Hello_" & nicheng
         Form1.Label21.Visible = False
         Form1.Label22.Caption = nicheng
         Form1.Timer4.Enabled = True
          
         Form3.Label3.Visible = False
         Form3.Label4.Visible = True
         Form3.Check1.Value = Check1.Value
         Form3.Check2.Value = Check2.Value
         
         If Dir("c:\�ٶ����ֲ�����", vbDirectory) <> "" Then
            Open "c:\�ٶ����ֲ�����\auto.txt" For Output As #4
               Print #4, Combo1.Text
              If Check1.Value = 1 Then
               Print #4, Text1.Text
              Else
               Print #4,
              End If
               Print #4, Check1.Value
               Print #4, Check2.Value
            Close #4
           
         End If
         Me.Hide
      ElseIf Text1.Text <> mima Then
         Close #1
         MsgBox "�������", , "��¼ʧ��"
         Command1.Enabled = True
         Label3.Caption = "��¼�ٶȣ�����������Ʒ��"
      End If
      
      Timer2.Enabled = False
   End If
   
End If



End Sub
