VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "百度°乐库"
   ClientHeight    =   7785
   ClientLeft      =   10020
   ClientTop       =   2070
   ClientWidth     =   7635
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "Form2.frx":038A
   ScaleHeight     =   7785
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7040
      Picture         =   "Form2.frx":AA7F
      ScaleHeight     =   270
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00DA8FCB&
      Caption         =   "下次不再显示"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   7530
      Width           =   1455
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7575
      ExtentX         =   13361
      ExtentY         =   12303
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "百度乐库"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_x, m_y As String
Private Sub Form_Load()
Form2.Left = Form1.Left + 4470
Form2.Top = Form1.Top
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   m_x = X
   m_y = Y
End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   Form2.Left = Form2.Left + X - m_x
   Form2.Top = Form2.Top + Y - m_y
End If

If 7020 < X And X < 7530 And 15 < Y And Y < 270 Then
  Picture1.Visible = True
Else
  Picture1.Visible = False
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  Picture1.Visible = False
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
unload Form2
End Sub

