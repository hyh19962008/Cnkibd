VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文本小工具"
   ClientHeight    =   4488
   ClientLeft      =   120
   ClientTop       =   768
   ClientWidth     =   5916
   BeginProperty Font 
      Name            =   "宋体"
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
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "去制表符"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "去换行符"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "中文符号"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "去空格"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "去回车"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "文本"
      Begin VB.Menu SelA 
         Caption         =   "全选"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu abt1 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub abt1_Click()
msg2 = MsgBox("去回车：去空行(符)" & vbCrLf & vbCrLf & "去空格：每次去一个空格" & vbCrLf & vbCrLf & "中文符号：将英文符号转换为中文" _
& vbCrLf & "(目前支持逗号句号冒号和分号)" & vbCrLf & vbCrLf & "v1.1 每个功能执行前自动从剪贴板获取文本" _
 & vbCrLf & vbCrLf & "v1.2 每个功能执行完毕后自动复制到剪贴板" & vbCrLf & vbCrLf & "v1.3 解决v1.2中的一个问题" & vbCrLf & vbCrLf & _
 "v1.4 添加最小化按钮" & vbCrLf & vbCrLf & "v1.5 增加摘取模式，在后台自动更新剪贴板" & vbCrLf & Space(8) & "内容到文本框中，用于连续获取文本" _
  & vbCrLf & vbCrLf & "v1.5.1 修改文本后自动更新剪贴版，防止误操作" & vbCrLf & vbCrLf & "v1.5.2 替换一处Bug" _
   & vbCrLf & vbCrLf & "v1.6 新增去换行符，加入符号图示" & vbCrLf & vbCrLf & "v1.7 分离摘取模式" & vbCrLf & vbCrLf & "v1.8 增加去制表符" & vbCrLf & vbCrLf & "v1.8.1 Bug修复", vbOKOnly, "说明")
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
        Text1.SelText = "，"
    ElseIf b = "." Then
        Text1.SelStart = i - 1
        Text1.SelLength = 1
        Text1.SelText = "。"
    ElseIf b = ";" Then
        Text1.SelStart = i - 1
        Text1.SelLength = 1
        Text1.SelText = "；"
    ElseIf b = ":" Then
        Text1.SelStart = i - 1
        Text1.SelLength = 1
        Text1.SelText = "："
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
