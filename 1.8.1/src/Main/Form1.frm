VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ı�С����"
   ClientHeight    =   4488
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
   ScaleHeight     =   4488
   ScaleWidth      =   5916
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "ȥ�Ʊ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1890
      TabIndex        =   7
      Top             =   3744
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3570
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   3675
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1050
      Picture         =   "Form1.frx":02E0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   3675
      Width           =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ȥ���з�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3150
      TabIndex        =   4
      Top             =   3150
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���ķ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4410
      TabIndex        =   3
      Top             =   3150
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ȥ�ո�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1890
      TabIndex        =   2
      Top             =   3150
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȥ�س�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   630
      TabIndex        =   1
      Top             =   3150
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   3000
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

Private Sub abt1_Click()
msg2 = MsgBox("ȥ�س���ȥ����(��)" & vbCrLf & vbCrLf & "ȥ�ո�ÿ��ȥһ���ո�" & vbCrLf & vbCrLf & "���ķ��ţ���Ӣ�ķ���ת��Ϊ����" _
& vbCrLf & "(Ŀǰ֧�ֶ��ž��ð�źͷֺ�)" & vbCrLf & vbCrLf & "v1.1 ÿ������ִ��ǰ�Զ��Ӽ������ȡ�ı�" _
 & vbCrLf & vbCrLf & "v1.2 ÿ������ִ����Ϻ��Զ����Ƶ�������" & vbCrLf & vbCrLf & "v1.3 ���v1.2�е�һ������" & vbCrLf & vbCrLf & _
 "v1.4 �����С����ť" & vbCrLf & vbCrLf & "v1.5 ����ժȡģʽ���ں�̨�Զ����¼�����" & vbCrLf & Space(8) & "���ݵ��ı����У�����������ȡ�ı�" _
  & vbCrLf & vbCrLf & "v1.5.1 �޸��ı����Զ����¼����棬��ֹ�����" & vbCrLf & vbCrLf & "v1.5.2 �滻һ��Bug" _
   & vbCrLf & vbCrLf & "v1.6 ����ȥ���з����������ͼʾ" & vbCrLf & vbCrLf & "v1.7 ����ժȡģʽ" & vbCrLf & vbCrLf & "v1.8 ����ȥ�Ʊ��" & vbCrLf & vbCrLf & "v1.8.1 Bug�޸�", vbOKOnly, "˵��")
End Sub

Private Sub Command1_Click()
ZT
Clipboard.Clear
Dim a$, i%, b$
For i = 1 To Len(Text1)
    b = Mid(Text1, i, 2)
    If b = Chr(13) + Chr(10) Then
    Text1 = Left(Text1, i - 1) & Right(Text1, Len(Text1) - i - 1)
    End If
Next
Fuzi
End Sub


Private Sub Command2_Click()
ZT
Clipboard.Clear
Dim a$, i%, b$
For i = 1 To Len(Text1)
    b = Mid(Text1, i, 1)
    If b = " " Then
    Text1 = Left(Text1, i - 1) & Right(Text1, Len(Text1) - i)
    End If
Next
Fuzi
End Sub

Private Sub Command3_Click()
ZT
Clipboard.Clear
Dim a$, i%, b$
For i = 1 To Len(Text1)
    b = Mid(Text1, i, 1)
    If b = "," Then
        Text1.SelStart = i - 1
        Text1.SelLength = 1
        Text1.SelText = "��"
    ElseIf b = "." Then
        Text1.SelStart = i - 1
        Text1.SelLength = 1
        Text1.SelText = "��"
    ElseIf b = ";" Then
        Text1.SelStart = i - 1
        Text1.SelLength = 1
        Text1.SelText = "��"
    ElseIf b = ":" Then
        Text1.SelStart = i - 1
        Text1.SelLength = 1
        Text1.SelText = "��"
    End If
Next
Fuzi
End Sub

Private Sub Command4_Click()
ZT
Clipboard.Clear
Dim a$, i%, b$
For i = 1 To Len(Text1)
    b = Mid(Text1, i, 2)
    If b = Chr(10) + Chr(13) Then
    Text1 = Left(Text1, i - 1) & Right(Text1, Len(Text1) - i - 1)
    End If
Next
Fuzi
End Sub

Private Sub Command5_Click()
ZT
Clipboard.Clear
Dim a$, i%, b$
For i = 1 To Len(Text1)
    b = Mid(Text1, i, 1)
    If b = Chr(9) Then
    Text1 = Left(Text1, i - 1) & Right(Text1, Len(Text1) - i)
    End If
Next
Fuzi
End Sub

Private Sub SelA_Click()
    With Text1
        .SelStart = 0
        .SelLength = Len(Text1.Text)
        .SetFocus
    End With
End Sub

Sub Fuzi()
Clipboard.SetText (Text1.Text)
End Sub

Sub ZT()
Text1.Text = Clipboard.GetText
End Sub

Private Sub Text1_LostFocus()
Clipboard.SetText (Text1.Text)
End Sub
