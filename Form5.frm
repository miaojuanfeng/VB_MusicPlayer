VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����������"
   ClientHeight    =   2250
   ClientLeft      =   4950
   ClientTop       =   4455
   ClientWidth     =   5295
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command2 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ͬʱ���浽�ļ�"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim song, X, Y, dizhi, yinyue, leixin As String
Dim q1, q2, q3, q4, q5, q6 As String

Private Sub Command1_Click()
If Text1.Text <> "" And Text2.Text <> "" Then

            If Form1.Label37.Caption = 1 Then
                Form1.Label1.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 2 Then
                Form1.Label2.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 3 Then
                Form1.Label3.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 4 Then
                Form1.Label4.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 5 Then
                Form1.Label5.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 6 Then
                Form1.Label6.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 7 Then
                Form1.Label7.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 8 Then
                Form1.Label8.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 9 Then
                Form1.Label9.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 10 Then
                Form1.Label10.Caption = " " & Text1.Text & " - " & Text2.Text
                
            ElseIf Form1.Label37.Caption = 11 Then
                Form1.Label23.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 12 Then
                Form1.Label24.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 13 Then
                Form1.Label25.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 14 Then
                Form1.Label26.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 15 Then
                Form1.Label27.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 16 Then
                Form1.Label28.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 17 Then
                Form1.Label29.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 18 Then
                Form1.Label30.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 19 Then
                Form1.Label31.Caption = " " & Text1.Text & " - " & Text2.Text
            ElseIf Form1.Label37.Caption = 20 Then
                Form1.Label32.Caption = " " & Text1.Text & " - " & Text2.Text
            End If
'���ڲ�����ʾ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
       If Form1.WindowsMediaPlayer1.URL = Form1.Label35.Caption Then
            If Form1.Label37.Caption = 1 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label1.Caption
            ElseIf Form1.Label37.Caption = 2 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label2.Caption
            ElseIf Form1.Label37.Caption = 3 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label3.Caption
            ElseIf Form1.Label37.Caption = 4 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label4.Caption
            ElseIf Form1.Label37.Caption = 5 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label5.Caption
            ElseIf Form1.Label37.Caption = 6 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label6.Caption
            ElseIf Form1.Label37.Caption = 7 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label7.Caption
            ElseIf Form1.Label37.Caption = 8 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label8.Caption
            ElseIf Form1.Label37.Caption = 9 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label9.Caption
            ElseIf Form1.Label37.Caption = 10 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label10.Caption
                
            ElseIf Form1.Label37.Caption = 11 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label23.Caption
            ElseIf Form1.Label37.Caption = 12 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label24.Caption
            ElseIf Form1.Label37.Caption = 13 Then
               Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label25.Caption
            ElseIf Form1.Label37.Caption = 14 Then
               Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label26.Caption
            ElseIf Form1.Label37.Caption = 15 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label27.Caption
            ElseIf Form1.Label37.Caption = 16 Then
               Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label28.Caption
            ElseIf Form1.Label37.Caption = 17 Then
               Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label29.Caption
            ElseIf Form1.Label37.Caption = 18 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label30.Caption
            ElseIf Form1.Label37.Caption = 19 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label31.Caption
            ElseIf Form1.Label37.Caption = 20 Then
                Form1.Label16.Caption = "���ڲ��ţ�" & Form1.Label32.Caption
            End If
        End If
'���浽�ļ�+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    If Check1.Value = 1 Then
                        song = Form1.Label35.Caption
                        
                        X = Len(song)
                        Y = Mid(song, X, 1)
                                                          Do While Y <> "\"
                                                             X = X - 1
                                                             Y = Mid(song, X, 1)
                                                          Loop
                        dizhi = Mid(song, 1, X)
                          X = X + 1
                        yinyue = Mid(song, X)
                        
                         X = Len(song)
                         Y = Mid(song, X, 1)
                                                      Do While Y <> "."
                                                         X = X - 1
                                                         Y = Mid(song, X, 1)
                                                      Loop
                         leixin = Mid(song, X)
                      
                      Form1.Label36.Caption = dizhi & Text1.Text & " - " & Text2.Text & leixin
                      
                     If Dir(song) <> "" Then
                        Name song As Form1.Label36.Caption
                     ElseIf Dir(song) = "" Then
                        MsgBox "�ļ�������", , "�ݸ���ʾ"
                    End If
                End If
      Unload Me
      
Else
      MsgBox "�����������Ϣ", , "�ݸ���ʾ"
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub
