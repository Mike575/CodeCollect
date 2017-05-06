VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000013&
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   8415
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "写入txt文件"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Open "App.Path & aaa.txt" For Output As #1
Print #1, Text1.Text
Close #1
End Sub

