VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "超级计算器 - v.1.15(Pre)"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   5895
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5895
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame10 
      Caption         =   "数据存储"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton mmode 
         BackColor       =   &H0080C0FF&
         Caption         =   "写"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton mc 
         BackColor       =   &H00FFFFC0&
         Caption         =   "MC"
         Height          =   375
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox memoryshow 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   24
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton mb5 
         Caption         =   "M5"
         Height          =   375
         Left            =   2040
         TabIndex        =   23
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton mb4 
         Caption         =   "M4"
         Height          =   375
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton mb3 
         Caption         =   "M3"
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton mb2 
         Caption         =   "M2"
         Height          =   375
         Left            =   600
         TabIndex        =   20
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton mb1 
         Caption         =   "M1"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton clear 
      BackColor       =   &H00FFFFC0&
      Caption         =   "清空"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Caption         =   "结果(点击复制)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   7
      Top             =   1920
      Width           =   2415
      Begin VB.TextBox result 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton calc 
      BackColor       =   &H0080C0FF&
      Caption         =   "计算"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox cha 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox num1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "第一数字"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "运算符"
      Height          =   735
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3836
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "数字保留"
      TabPicture(0)   =   "Form1.frx":218A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "import"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "sswr"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cl"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "科学计数法"
      TabPicture(1)   =   "Form1.frx":218C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cl 
         BackColor       =   &H00FFFFC0&
         Caption         =   "清空"
         Height          =   495
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1440
         Width           =   855
      End
      Begin VB.Frame Frame9 
         Caption         =   "保留位数"
         Height          =   855
         Left            =   3960
         TabIndex        =   16
         Top             =   480
         Width           =   1335
         Begin VB.TextBox weishu 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   27
            Top             =   360
            Width           =   615
         End
         Begin VB.Label 位数 
            Caption         =   "位数:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "点击复制"
         Height          =   735
         Left            =   1560
         TabIndex        =   14
         Top             =   1320
         Width           =   2175
         Begin VB.TextBox sswrt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "原始数据"
         Height          =   735
         Left            =   1560
         TabIndex        =   12
         Top             =   480
         Width           =   2175
         Begin VB.TextBox importt 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.CommandButton sswr 
         Caption         =   "四舍五入"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton import 
         Caption         =   "导入结果"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   ":( 暂未开放"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73080
         TabIndex        =   30
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.TextBox num2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "第二数字"
      Height          =   735
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Menu 计算 
      Caption         =   "计算"
      Begin VB.Menu 导入数据 
         Caption         =   "导入数据(待开放)"
      End
   End
   Begin VB.Menu 帮助 
      Caption         =   "帮助"
   End
   Begin VB.Menu 关于 
      Caption         =   "关于"
      Begin VB.Menu 关于超级计算器 
         Caption         =   "关于""超级计算器"""
      End
      Begin VB.Menu fgx 
         Caption         =   "-"
      End
      Begin VB.Menu 检查更新 
         Caption         =   "检查更新"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##0.添加操作快捷键；1.最好能实现2pi=6.28...这样的功能3.更多三角函数
'###1.【i】提示！(o゜▽゜)o☆2.【×】错误！(っ °Д °;)っ3.【!】警告！Σ(っ °Д °;)っ
'20240410-Now
Dim a#, b#
'第一数字：作为底数、被转换数
'第二数字：作为真数、三角函数自变量
Dim mm% '定义数据存储模式
Dim m1#, m2#, m3#, m4#, m5#    '定义数据存储变量
Dim ws%  '定义保留位数
Dim blm%

Private Sub Form_Load()
mm = 1
    '初始化内存模式切换
    If mm Mod 2 = 1 Then '=1为写模式
        mmode.Caption = "写"
    Else
        mmode.Caption = "读"
    End If
blm = 1
    '初始化保留模式切换
    If blm Mod 2 = 1 Then '=1为写模式
        mmode.Caption = "写"
    Else
        mmode.Caption = "读"
    End If
End Sub

Private Sub mmode_Click()
mm = mm + 1
    If mm Mod 2 = 1 Then '单数写，双数读
        mmode.Caption = "写"
    Else
        mmode.Caption = "读"
    End If
End Sub '内存读写切换

Private Sub mb1_Click()
    If mm Mod 2 = 1 Then
        m1 = Val(result.Text)
    Else
        memoryshow.Text = m1
    End If
End Sub '数据存储1

Private Sub mb2_Click()
If mm Mod 2 = 1 Then
        m2 = Val(result.Text)
    Else
        memoryshow.Text = m2
    End If
End Sub '数据存储2

Private Sub mb3_Click()
If mm Mod 2 = 1 Then
        m3 = Val(result.Text)
    Else
        memoryshow.Text = m3
    End If
End Sub '数据存储3

Private Sub mb4_Click()
If mm Mod 2 = 1 Then
        m4 = Val(result.Text)
    Else
        memoryshow.Text = m4
    End If
End Sub '数据存储5

Private Sub mb5_Click()
If mm Mod 2 = 1 Then
        m5 = Val(result.Text)
    Else
        memoryshow.Text = m5
    End If
End Sub '数据存储6

Private Sub mc_Click()
m0 = 0: m1 = 0: m2 = 0: m3 = 0: m4 = 0: m5 = 0
End Sub '清除数据存储

Private Sub cha_change()    '运算符字体变换
    If Len(cha.Text) <= 3 Then
        cha.FontSize = 13
    ElseIf Len(cha.Text) = 4 Then
        cha.Font = "Arial Narrow"
        cha.FontBold = True
        cha.FontSize = 11
    ElseIf Len(cha.Text) > 4 Then
        cha.Font = "Arial Narrow"
        cha.FontBold = True
        cha.FontSize = 10
    End If
End Sub

Private Sub calc_Click()    '计算

    If num1.Text = "e" Then
        a = 2.71828182845905
    ElseIf num1.Text = "π" Or num1.Text = "Π" Or num1.Text = "pai" Or num1.Text = "pi" Then
        a = 3.14159265358979
    Else: a = Val(num1.Text)
    End If
    If num2.Text = "e" Then
        b = 2.71828182845905
    ElseIf num2.Text = "π" Or num2.Text = "Π" Or num2.Text = "pai" Or num2.Text = "pi" Then
        b = 3.14159265358979
    Else: b = Val(num2.Text)
    End If  '检测输入值是否为e,π
    
    If cha.Text = "+" Then
        result.Text = a + b
    ElseIf cha.Text = "-" Then
        result.Text = a - b
    ElseIf cha.Text = "*" Or cha.Text = "×" Then
        result.Text = a * b
    ElseIf cha.Text = "/" Or cha.Text = "÷" Then
        If b = 0 Then
        MsgBox "被除数不可为0！→ 请检查“第二数字”！", vbCritical, "【×】错误！(っ °Д °;)っ"  '检测/0
        Else: result.Text = a / b '
        End If
    ElseIf cha.Text = "//" Then
        result.Text = Str(a \ b) & "..." & Str(a Mod b)
    ElseIf cha.Text = "√" Or cha.Text = "r" Then
        If a = 0 Then
        MsgBox "方根数不可为0！→ 请检查“第一数字”！", vbCritical, "【×】错误！(っ °Д °;)っ"  '检测/0
        ElseIf num1.Text = "" Then  '不输方根数则默认为开平方
        result.Text = Sqr(b)
        Else: result.Text = b ^ (1 / a)
        End If
        
    ElseIf cha.Text = "^" Then
        result.Text = a ^ b
    ElseIf cha.Text = "log" Then
        result.Text = Log(b) / Log(a)   '对数运算换底公式，b=真数，a=底数
        
    ElseIf cha.Text = "o" Or cha.Text = "0" Or cha.Text = "°" Then
        result.Text = Str(a * (180 / 3.14159265358979)) & "°"
    ElseIf cha.Text = "rad" Then
        result.Text = Str(a * 3.14159265358979 / 180) & "rad"   '角度弧度互换
        
    ElseIf cha.Text = "sin" Then
        result.Text = Sin(b * 3.14159265358979 / 180)
    ElseIf cha.Text = "cos" Then
        result.Text = Cos(b * 3.14159265358979 / 180)
    ElseIf cha.Text = "tan" Then
        result.Text = Tan(b * 3.14159265358979 / 180)   '三角函数，输入全部是角度
            If Str(Int(Str((b / 180) - 0.5))) = Str((b / 180) - 0.5) Then
                MsgBox "被计算的角度不能为90°+k·180°, k∈Z → 请检查“第二数字”！", vbCritical, "【×】错误！(っ °Д °;)っ"    '定义域检测
            End If
    
    ElseIf cha.Text = "arcsin" Then
        If b < -1 Or b > 1 Then
            MsgBox "被计算的数不能小于-1或大于1 → 请检查“第二数字”！", vbCritical, "【×】错误！(っ °Д °;)っ"    '定义域检测
        ElseIf b = 1 Then
            result.Text = "90°"
        ElseIf b = -1 Then
            result.Text = "-90°"
        Else
            result.Text = Str((Atn(b / Sqr(-b * b + 1))) * (180 / 3.14159265358979)) & "°"
        End If
    ElseIf cha.Text = "arccos" Then
        If b < -1 Or b > 1 Then
            MsgBox "被计算的数不能小于-1或大于1 → 请检查“第二数字”！", vbCritical, "【×】错误！(っ °Д °;)っ"    '定义域检测
        ElseIf b = 1 Then
            result.Text = "0°"
        ElseIf b = -1 Then
            result.Text = "120°"
        Else
            result.Text = Str(Atn(Sqr(1 - b * b) / b) * (180 / 3.14159265358979)) & "°"
        End If
    ElseIf cha.Text = "arctan" Then
    result.Text = Str(Atn(b) * (180 / 3.14159265358979)) & "°"     '反三角函数区#可以继续添加
    
    
    
    
    End If  '★主endif
End Sub

Private Sub clear_Click()   '清除计算区
result.Text = ""
num1.Text = ""
num2.Text = ""
cha.Text = ""
End Sub

Private Sub cl_Click()
importt.Text = ""
sswrt.Text = ""

End Sub

Private Sub Frame4_Click()  '复制数据
Clipboard.SetText (result.Text)
End Sub

Private Sub Frame6_Click()  '复制数据
Clipboard.SetText (sswrt.Text)
End Sub

Private Sub Frame7_Click()  '复制数据
Clipboard.SetText (srfzt.Text)
End Sub

Private Sub Frame8_Click()  '复制数据
Clipboard.SetText (jwfzt.Text)
End Sub

Private Sub SSTab1_GotFocus()   '切换快捷键至Tab框
'还没写
End Sub

Private Sub import_Click()  '导入结果
importt.Text = result.Text
End Sub

Private Sub sswr_Click()    '计算四舍五入
ws = Val(weishu.Text)
sswrt.Text = Round(importt, ws)
End Sub


Private Sub weishu_LostFocus()  '检测位数0<ws<=16#考虑增加-位数：向前保留，使用科学计数法表示
    If ws > 16 Then
        MsgBox "保留的位数不能超过16位！→ 请检查“位数”！", vbCritical, "【×】错误！(っ °Д °;)っ"
    ElseIf ws < 0 Then
        MsgBox "保留的位数不能为负数！→ 请检查“位数”！", vbCritical, "【×】错误！(っ °Д °;)っ"
    End If
End Sub

Private Sub 帮助_Click()
MsgBox "即将打开网页（https://567z30m162.goho.co/text/sucalc.html）→ 本功能需要联网。", vbInformation, "【i】提示！(o゜▽゜)o☆"
Shell "explorer.exe https://567z30m162.goho.co/text/sucalc.html#help"
End Sub

Private Sub 关于超级计算器_Click()
MsgBox "即将打开网页（https://567z30m162.goho.co/text/sucalc.html）→ 本功能需要联网。", vbInformation, "【i】提示！(o゜▽゜)o☆"
Shell "explorer.exe https://567z30m162.goho.co/text/sucalc.html"
End Sub

Private Sub 检查更新_Click()
MsgBox "即将打开网页（https://567z30m162.goho.co/text/sucalc.html）→ 本功能需要联网。", vbInformation, "【i】提示！(o゜▽゜)o☆"
Shell "explorer.exe https://567z30m162.goho.co/text/sucalc.html#checknew"
End Sub
