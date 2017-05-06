VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7260
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Caption         =   "Writetext"
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "WriteExcel"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "退出窗体"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "清空文本框"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "一次性导入(正常)"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "分割导入"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "逐行导入"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "一次性导入(出错)"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000013&
      Caption         =   "内容二："
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "内容一："
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000013&
      Caption         =   "导入文本文件中内容："
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""

Open "abc.txt" For Input As #1

Text1.Text = Input(LOF(1), 1)

Close #1
End Sub

Private Sub Command2_Click()
Text1.Text = ""

Open "abc.txt" For Input As #1

Do While Not EOF(1)

Line Input #1, InputData

Text1.Text = Text1.Text + InputData + vbCrLf

Loop

Close #1
End Sub

Private Sub Command3_Click()
Dim stra As String

Open "abc.txt" For Input As #1

stra = StrConv(InputB$(LOF(1), #1), vbUnicode)

Close #1

b = Split(stra, "，")
Text2.Text = b(0)
Text3.Text = b(1)
End Sub

Private Sub Command4_Click()

Text1.Text = ""

Open "abc.txt" For Input As #1

Text1.Text = StrConv(InputB$(LOF(1), #1), vbUnicode)


Close #1
End Sub

Private Sub Command5_Click()
Text1.Text = ""
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Form2.Visible = True
End Sub

Private Sub Command8_Click()
Form3.Visible = True

End Sub

