VERSION 5.00
Begin VB.Form form1 
   Caption         =   "简易计算器"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   4290
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command7 
      Caption         =   "清空"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "单/双"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "退出"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "结果"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
      Begin VB.Label Label1 
         Caption         =   "在此显示结果"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "÷"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "×"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "数据1"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "数据2"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   2520
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "单"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   1560
      Width           =   255
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer

Private Sub Command1_Click()
a = Val(Text1.Text)
b = Val(Text2.Text)
Label1.Caption = a1 + b1
End Sub

Private Sub Command2_Click()
a = Val(Text1.Text)
b = Val(Text2.Text)
Label1.Caption = a1 - b1
End Sub

Private Sub Command3_Click()
a = Val(Text1.Text)
b = Val(Text2.Text)
Label1.Caption = a1 * b1
End Sub

Private Sub Command4_Click()
a = Val(Text1.Text)
b = Val(Text2.Text)
Label1.Caption = a1 / b1
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
c = c + 1
If c Mod 2 = 0 Then
Label2.Caption = "单"
a1 = CSng(a)
b1 = CSng(b)
'd = "#.######"
Else
Label2.Caption = "双"
a1 = CDbl(a)
b1 = CDbl(b)
'd = "#.###############"
End If
End Sub

Private Sub Command7_Click()
Text1.Text = ""
Text2.Text = ""
Label1.Caption = "在此显示结果"
End Sub

Private Sub Form_Load()
Dim a As Single
Dim b As Single
a1 = a
b1 = b
c = 0
End Sub
