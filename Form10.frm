VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4680
   LinkTopic       =   "Form10"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu right 
      Caption         =   "�Ҽ�"
      Visible         =   0   'False
      Begin VB.Menu play 
         Caption         =   "  ����(&P)"
      End
      Begin VB.Menu pause 
         Caption         =   "  ��ͣ"
      End
      Begin VB.Menu stop 
         Caption         =   "  ֹͣ"
      End
      Begin VB.Menu delete 
         Caption         =   "  ɾ��(&D)  "
      End
      Begin VB.Menu adafsfdgdgdfgdgdgf 
         Caption         =   "-"
      End
      Begin VB.Menu addbendi 
         Caption         =   "  ��ӱ�������(&b)"
      End
      Begin VB.Menu addwangluo 
         Caption         =   "  �����������(&w)"
      End
      Begin VB.Menu sdafsdfsfsgdgdg 
         Caption         =   "-"
      End
      Begin VB.Menu add 
         Caption         =   "  ��ӵ�  "
         Begin VB.Menu mycount 
            Caption         =   "  �ҵ��ղ� "
         End
      End
      Begin VB.Menu adsasfdsdfsgfgdf 
         Caption         =   "-"
      End
      Begin VB.Menu rename 
         Caption         =   "  ������(&M) "
      End
      Begin VB.Menu asdasdadasdas 
         Caption         =   "-"
      End
      Begin VB.Menu system 
         Caption         =   "  ϵͳ����(&s)"
      End
      Begin VB.Menu playmodel 
         Caption         =   "  ����ģʽ"
         Begin VB.Menu shunxu 
            Caption         =   "  ˳�򲥷�"
         End
         Begin VB.Menu shunxuxunhuan 
            Caption         =   "  ˳��ѭ������"
         End
         Begin VB.Menu shuiji 
            Caption         =   "  �������"
         End
         Begin VB.Menu danqu 
            Caption         =   "  ����ѭ��"
         End
      End
      Begin VB.Menu fsdfsddgdfgdfgdgdf 
         Caption         =   "-"
      End
      Begin VB.Menu save 
         Caption         =   "  ���沥���б�"
      End
      Begin VB.Menu clean 
         Caption         =   "  ��������б�"
      End
   End
   Begin VB.Menu addmusic 
      Caption         =   "���"
      Visible         =   0   'False
      Begin VB.Menu addbendimusic 
         Caption         =   "  ��ӱ�������(&b)"
      End
      Begin VB.Menu addwangluomusic 
         Caption         =   "  �����������(&w)"
      End
   End
   Begin VB.Menu delete2 
      Caption         =   "ɾ��"
      Visible         =   0   'False
      Begin VB.Menu deletemusic 
         Caption         =   "  ɾ����ѡ����(&d)"
      End
      Begin VB.Menu asdasdasdsdsds 
         Caption         =   "-"
      End
      Begin VB.Menu cleanmusic 
         Caption         =   "  ��յ�ǰ�б�"
      End
   End
   Begin VB.Menu model 
      Caption         =   "ģʽ"
      Visible         =   0   'False
      Begin VB.Menu shunxu2 
         Caption         =   "  ˳�򲥷�"
      End
      Begin VB.Menu shunxuxunhuan2 
         Caption         =   "  ˳��ѭ������"
      End
      Begin VB.Menu shuiji2 
         Caption         =   "  �������"
      End
      Begin VB.Menu danqu2 
         Caption         =   "  ����ѭ��"
      End
   End
   Begin VB.Menu youshangjiao 
      Caption         =   "���Ͻ�"
      Visible         =   0   'False
      Begin VB.Menu load 
         Caption         =   "  ��½"
      End
      Begin VB.Menu unload 
         Caption         =   "  ע����¼"
      End
      Begin VB.Menu aaaaaaaaaaaaaaaaaa 
         Caption         =   "-"
      End
      Begin VB.Menu tianjiabendi 
         Caption         =   "  ��ӱ�������"
      End
      Begin VB.Menu tianjiawangluo 
         Caption         =   "  �����������"
      End
      Begin VB.Menu saaaaaaaaaa 
         Caption         =   "-"
      End
      Begin VB.Menu set 
         Caption         =   "  ϵͳ����"
      End
      Begin VB.Menu look 
         Caption         =   "  �鿴���ܽ���"
      End
      Begin VB.Menu sadasddddddddddddddd 
         Caption         =   "-"
      End
      Begin VB.Menu help 
         Caption         =   "  ����"
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub addbendi_Click()
If Form1.Label41.Caption = "1" Then
   Form1.Label41.Caption = "2"
ElseIf Form1.Label41.Caption = "2" Then
   Form1.Label41.Caption = "1"
End If
End Sub

Private Sub addbendimusic_Click()
If Form1.Label41.Caption = "1" Then
   Form1.Label41.Caption = "2"
ElseIf Form1.Label41.Caption = "2" Then
   Form1.Label41.Caption = "1"
End If
End Sub

Private Sub addwangluo_Click()
Form4.Show 1
End Sub

Private Sub addwangluomusic_Click()
Form4.Show 1
End Sub

Private Sub clean_Click()
If Form1.Label45.Caption = "1" Then
   Form1.Label45.Caption = "2"
ElseIf Form1.Label45.Caption = "2" Then
   Form1.Label45.Caption = "1"
End If
End Sub

Private Sub cleanmusic_Click()
If Form1.Label45.Caption = "1" Then
   Form1.Label45.Caption = "2"
ElseIf Form1.Label45.Caption = "2" Then
   Form1.Label45.Caption = "1"
End If
End Sub

Private Sub danqu_Click()
   Form3.Option4.Value = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = True
   
      Form10.shunxu2.Checked = False
    Form10.shunxuxunhuan2.Checked = False
    Form10.shuiji2.Checked = False
    Form10.danqu2.Checked = True
End Sub

Private Sub danqu2_Click()
Form3.Option4.Value = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = True
   
      Form10.shunxu2.Checked = False
    Form10.shunxuxunhuan2.Checked = False
    Form10.shuiji2.Checked = False
    Form10.danqu2.Checked = True
End Sub

Private Sub delete_Click()
If Form1.Label40.Caption = "1" Then
   Form1.Label40.Caption = "2"
ElseIf Form1.Label40.Caption = "2" Then
   Form1.Label40.Caption = "1"
End If
End Sub

Private Sub delete2_Click()
If Form1.Label1.BackColor = &HFF8080 Then
Form1.Label46.Caption = "1"
ElseIf Form1.Label2.BackColor = &HFF8080 Then
Form1.Label46.Caption = "2"
ElseIf Form1.Label3.BackColor = &HFF8080 Then
Form1.Label46.Caption = "3"
ElseIf Form1.Label4.BackColor = &HFF8080 Then
Form1.Label46.Caption = "4"
ElseIf Form1.Label5.BackColor = &HFF8080 Then
Form1.Label46.Caption = "5"
ElseIf Form1.Label6.BackColor = &HFF8080 Then
Form1.Label46.Caption = "6"
ElseIf Form1.Label7.BackColor = &HFF8080 Then
Form1.Label46.Caption = "7"
ElseIf Form1.Label8.BackColor = &HFF8080 Then
Form1.Label46.Caption = "8"
ElseIf Form1.Label9.BackColor = &HFF8080 Then
Form1.Label46.Caption = "9"
ElseIf Form1.Label10.BackColor = &HFF8080 Then
Form1.Label46.Caption = "10"
End If
'�ҵ��ղ�
If Form1.Label23.BackColor = vbRed Then
Form1.Label46.Caption = "11"
ElseIf Form1.Label24.BackColor = vbRed Then
Form1.Label46.Caption = "12"
ElseIf Form1.Label25.BackColor = vbRed Then
Form1.Label46.Caption = "13"
ElseIf Form1.Label26.BackColor = vbRed Then
Form1.Label46.Caption = "14"
ElseIf Form1.Label27.BackColor = vbRed Then
Form1.Label46.Caption = "15"
ElseIf Form1.Label28.BackColor = vbRed Then
Form1.Label46.Caption = "16"
ElseIf Form1.Label29.BackColor = vbRed Then
Form1.Label46.Caption = "17"
ElseIf Form1.Label30.BackColor = vbRed Then
Form1.Label46.Caption = "18"
ElseIf Form1.Label31.BackColor = vbRed Then
Form1.Label46.Caption = "19"
ElseIf Form1.Label32.BackColor = vbRed Then
Form1.Label46.Caption = "20"
End If
End Sub

Private Sub deletemusic_Click()
If Form1.Label40.Caption = "1" Then
   Form1.Label40.Caption = "2"
ElseIf Form1.Label40.Caption = "2" Then
   Form1.Label40.Caption = "1"
End If
End Sub

Private Sub deletemusic2_Click()

End Sub

Private Sub Form_Load()
i = 0
Form1.Label60.Caption = i
End Sub

Private Sub load_Click()
If Form1.Label60.Caption <> "0" Then
  Form7.Timer1.Enabled = True
End If
Form7.Show 1
End Sub

Private Sub mycount_Click()
If Form1.Label42.Caption = "1" Then
   Form1.Label42.Caption = "2"
ElseIf Form1.Label42.Caption = "2" Then
   Form1.Label42.Caption = "1"
End If
End Sub

Private Sub pause_Click()
If Form1.Label58.Caption = "1" Then
   Form1.Label58.Caption = "2"
ElseIf Form1.Label58.Caption = "2" Then
   Form1.Label58.Caption = "1"
End If
End Sub

Private Sub play_Click()
If Form1.Label39.Caption = "1" Then
   Form1.Label39.Caption = "2"
ElseIf Form1.Label39.Caption = "2" Then
   Form1.Label39.Caption = "1"
End If
End Sub

Private Sub rename_Click()
If Form1.Label43.Caption = "1" Then
   Form1.Label43.Caption = "2"
ElseIf Form1.Label43.Caption = "2" Then
   Form1.Label43.Caption = "1"
End If
End Sub

Private Sub right_Click()
If Form1.Label1.BackColor = &HFF8080 Then
Form1.Label46.Caption = "1"
ElseIf Form1.Label2.BackColor = &HFF8080 Then
Form1.Label46.Caption = "2"
ElseIf Form1.Label3.BackColor = &HFF8080 Then
Form1.Label46.Caption = "3"
ElseIf Form1.Label4.BackColor = &HFF8080 Then
Form1.Label46.Caption = "4"
ElseIf Form1.Label5.BackColor = &HFF8080 Then
Form1.Label46.Caption = "5"
ElseIf Form1.Label6.BackColor = &HFF8080 Then
Form1.Label46.Caption = "6"
ElseIf Form1.Label7.BackColor = &HFF8080 Then
Form1.Label46.Caption = "7"
ElseIf Form1.Label8.BackColor = &HFF8080 Then
Form1.Label46.Caption = "8"
ElseIf Form1.Label9.BackColor = &HFF8080 Then
Form1.Label46.Caption = "9"
ElseIf Form1.Label10.BackColor = &HFF8080 Then
Form1.Label46.Caption = "10"
End If
'�ҵ��ղ�
If Form1.Label23.BackColor = vbRed Then
Form1.Label46.Caption = "11"
ElseIf Form1.Label24.BackColor = vbRed Then
Form1.Label46.Caption = "12"
ElseIf Form1.Label25.BackColor = vbRed Then
Form1.Label46.Caption = "13"
ElseIf Form1.Label26.BackColor = vbRed Then
Form1.Label46.Caption = "14"
ElseIf Form1.Label27.BackColor = vbRed Then
Form1.Label46.Caption = "15"
ElseIf Form1.Label28.BackColor = vbRed Then
Form1.Label46.Caption = "16"
ElseIf Form1.Label29.BackColor = vbRed Then
Form1.Label46.Caption = "17"
ElseIf Form1.Label30.BackColor = vbRed Then
Form1.Label46.Caption = "18"
ElseIf Form1.Label31.BackColor = vbRed Then
Form1.Label46.Caption = "19"
ElseIf Form1.Label32.BackColor = vbRed Then
Form1.Label46.Caption = "20"
End If
End Sub

Private Sub save_Click()
If Form1.Label44.Caption = "1" Then
   Form1.Label44.Caption = "2"
ElseIf Form1.Label44.Caption = "2" Then
   Form1.Label44.Caption = "1"
End If
End Sub

Private Sub set_Click()
Form3.Show 1
End Sub

Private Sub shuiji_Click()
   Form3.Option3.Value = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = True
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = False
    Form10.shunxuxunhuan2.Checked = False
    Form10.shuiji2.Checked = True
    Form10.danqu2.Checked = False
End Sub

Private Sub shuiji2_Click()
Form3.Option3.Value = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = True
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = False
    Form10.shunxuxunhuan2.Checked = False
    Form10.shuiji2.Checked = True
    Form10.danqu2.Checked = False
End Sub

Private Sub shunxu_Click()
   Form3.Option1.Value = True

   
   Form10.shunxu.Checked = True
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = True
    Form10.shunxuxunhuan2.Checked = False
    Form10.shuiji2.Checked = False
    Form10.danqu2.Checked = False
End Sub

Private Sub shunxu2_Click()
   Form3.Option1.Value = True

   
   Form10.shunxu.Checked = True
   Form10.shunxuxunhuan.Checked = False
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = True
    Form10.shunxuxunhuan2.Checked = False
    Form10.shuiji2.Checked = False
    Form10.danqu2.Checked = False
End Sub

Private Sub shunxuxunhuan_Click()
Form3.Option2.Value = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = True
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = False
    Form10.shunxuxunhuan2.Checked = True
    Form10.shuiji2.Checked = False
    Form10.danqu2.Checked = False
End Sub

Private Sub shunxuxunhuan2_Click()
Form3.Option2.Value = True

   
   Form10.shunxu.Checked = False
   Form10.shunxuxunhuan.Checked = True
   Form10.shuiji.Checked = False
   Form10.danqu.Checked = False
   
   Form10.shunxu2.Checked = False
    Form10.shunxuxunhuan2.Checked = True
    Form10.shuiji2.Checked = False
    Form10.danqu2.Checked = False
End Sub

Private Sub stop_Click()
If Form1.Label59.Caption = "1" Then
   Form1.Label59.Caption = "2"
ElseIf Form1.Label59.Caption = "2" Then
   Form1.Label59.Caption = "1"
End If
End Sub

Private Sub system_Click()
Form3.Show 1
End Sub

Private Sub tianjiabendi_Click()
If Form1.Label41.Caption = "1" Then
   Form1.Label41.Caption = "2"
ElseIf Form1.Label41.Caption = "2" Then
   Form1.Label41.Caption = "1"
End If
End Sub

Private Sub tianjiawangluo_Click()
Form4.Show 1
End Sub

Private Sub unload_Click()
Form1.Label20.Visible = False
Form1.Label21.Visible = True
Form1.Label22.Caption = ""
Form1.Label48.Visible = True
i = i + 1
Form1.Label60.Caption = i

With Form3
.Picture5.Top = 5790
.Picture6.Top = 5790
.Picture7.Top = 5790
.Label2.Visible = True
.Label13.Visible = True
.Option1.Visible = True
.Option2.Visible = True
.Option3.Visible = True
.Option4.Visible = True
.Check4.Visible = True

.Label5.Visible = False
.Label6.Visible = False
.Check1.Visible = False
.Check2.Visible = False
.Image1.Visible = False
.Picture15.Visible = False
.Label12.Visible = False
.Label8.Visible = False
.Label9.Visible = False
.Label10.Visible = False
.Text1.Visible = False


.Picture1.Visible = True
.Picture2.Visible = False

.Label1.Visible = False
.Label3.Visible = False
.Label4.Visible = False
End With

If Form1.Picture15.Visible = True Then
   Form1.Picture19.Visible = True
End If

End Sub
