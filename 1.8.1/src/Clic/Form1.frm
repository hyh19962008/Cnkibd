VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ժȡ"
   ClientHeight    =   3708
   ClientLeft      =   120
   ClientTop       =   768
   ClientWidth     =   5916
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.8
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3708
   ScaleWidth      =   5916
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   3120
   End
   Begin VB.TextBox Text1 
      Height          =   3756
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5910
   End
   Begin VB.Menu CAT 
      Caption         =   "�ı�"
      Begin VB.Menu SelA 
         Caption         =   "ȫѡ"
         Shortcut        =   ^A
      End
      Begin VB.Menu Sel0 
         Caption         =   "ժȡģʽ"
         Begin VB.Menu Sel1 
            Caption         =   "�ر�"
            Checked         =   -1  'True
         End
         Begin VB.Menu Sel2 
            Caption         =   "�س��ָ�"
         End
         Begin VB.Menu Sel3 
            Caption         =   "�ո�ָ�"
         End
         Begin VB.Menu Sel4 
            Caption         =   "�޷ָ���"
         End
      End
   End
   Begin VB.Menu abt1 
      Caption         =   "����"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ttt As String
Private Sub abt1_Click()
msg2 = MsgBox("v1.1 �ں�̨������¼����������" & vbCrLf & vbCrLf & "v1.12 �޸������屻ռ��ʱ�����Ĵ���", vbOKOnly)
End Sub

Private Sub Form_Load()
ttt = Clipboard.GetText
End Sub

Private Sub Sel1_Click()
If Sel1.Checked = False Then
    Sel1.Checked = True
End If
Sel2.Checked = False
Sel3.Checked = False
Sel4.Checked = False
Timer1 = False
End Sub

Private Sub Sel2_Click()
If Sel2.Checked = False Then
    Sel2.Checked = True
End If
Sel1.Checked = False
Sel3.Checked = False
Sel4.Checked = False
Timer1 = True
End Sub

Private Sub Sel3_Click()
If Sel3.Checked = False Then
    Sel3.Checked = True
End If
Sel1.Checked = False
Sel2.Checked = False
Sel4.Checked = False
Timer1 = True
End Sub

Private Sub Sel4_Click()
If Sel4.Checked = False Then
    Sel4.Checked = True
End If
Sel1.Checked = False
Sel2.Checked = False
Sel3.Checked = False
Timer1 = True
End Sub

Private Sub SelA_Click()
    With Text1
        .SelStart = 0
        .SelLength = Len(Text1.Text)
        .SetFocus
    End With
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Clipboard.GetText <> ttt Then
    ttt = Clipboard.GetText
    If Sel2.Checked = True Then Text1.SelText = ttt & vbCrLf
    If Sel3.Checked = True Then Text1.SelText = ttt & "  "
    If Sel4.Checked = True Then Text1.SelText = ttt
End If
End Sub
