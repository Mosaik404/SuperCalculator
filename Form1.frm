VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���������� - v.1.15(Pre)"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame10 
      Caption         =   "���ݴ洢"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton mmode 
         BackColor       =   &H0080C0FF&
         Caption         =   "д"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "���"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Caption         =   "���(�������)"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����"
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
      Caption         =   "��һ����"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "�����"
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "���ֱ���"
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
      TabCaption(1)   =   "��ѧ������"
      TabPicture(1)   =   "Form1.frx":218C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cl 
         BackColor       =   &H00FFFFC0&
         Caption         =   "���"
         Height          =   495
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1440
         Width           =   855
      End
      Begin VB.Frame Frame9 
         Caption         =   "����λ��"
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
         Begin VB.Label λ�� 
            Caption         =   "λ��:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "�������"
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
         Caption         =   "ԭʼ����"
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
         Caption         =   "��������"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton import 
         Caption         =   "������"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   ":( ��δ����"
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Caption         =   "�ڶ�����"
      Height          =   735
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu �������� 
         Caption         =   "��������(������)"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ���ڳ��������� 
         Caption         =   "����""����������"""
      End
      Begin VB.Menu fgx 
         Caption         =   "-"
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##0.��Ӳ�����ݼ���1.�����ʵ��2pi=6.28...�����Ĺ���3.�������Ǻ���
'###1.��i����ʾ��(o�b���b)o��2.����������(�� �㧥 ��;)��3.��!�����棡��(�� �㧥 ��;)��
'20240410-Now
Dim a#, b#
'��һ���֣���Ϊ��������ת����
'�ڶ����֣���Ϊ���������Ǻ����Ա���
Dim mm% '�������ݴ洢ģʽ
Dim m1#, m2#, m3#, m4#, m5#    '�������ݴ洢����
Dim ws%  '���屣��λ��
Dim blm%

Private Sub Form_Load()
mm = 1
    '��ʼ���ڴ�ģʽ�л�
    If mm Mod 2 = 1 Then '=1Ϊдģʽ
        mmode.Caption = "д"
    Else
        mmode.Caption = "��"
    End If
blm = 1
    '��ʼ������ģʽ�л�
    If blm Mod 2 = 1 Then '=1Ϊдģʽ
        mmode.Caption = "д"
    Else
        mmode.Caption = "��"
    End If
End Sub

Private Sub mmode_Click()
mm = mm + 1
    If mm Mod 2 = 1 Then '����д��˫����
        mmode.Caption = "д"
    Else
        mmode.Caption = "��"
    End If
End Sub '�ڴ��д�л�

Private Sub mb1_Click()
    If mm Mod 2 = 1 Then
        m1 = Val(result.Text)
    Else
        memoryshow.Text = m1
    End If
End Sub '���ݴ洢1

Private Sub mb2_Click()
If mm Mod 2 = 1 Then
        m2 = Val(result.Text)
    Else
        memoryshow.Text = m2
    End If
End Sub '���ݴ洢2

Private Sub mb3_Click()
If mm Mod 2 = 1 Then
        m3 = Val(result.Text)
    Else
        memoryshow.Text = m3
    End If
End Sub '���ݴ洢3

Private Sub mb4_Click()
If mm Mod 2 = 1 Then
        m4 = Val(result.Text)
    Else
        memoryshow.Text = m4
    End If
End Sub '���ݴ洢5

Private Sub mb5_Click()
If mm Mod 2 = 1 Then
        m5 = Val(result.Text)
    Else
        memoryshow.Text = m5
    End If
End Sub '���ݴ洢6

Private Sub mc_Click()
m0 = 0: m1 = 0: m2 = 0: m3 = 0: m4 = 0: m5 = 0
End Sub '������ݴ洢

Private Sub cha_change()    '���������任
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

Private Sub calc_Click()    '����

    If num1.Text = "e" Then
        a = 2.71828182845905
    ElseIf num1.Text = "��" Or num1.Text = "��" Or num1.Text = "pai" Or num1.Text = "pi" Then
        a = 3.14159265358979
    Else: a = Val(num1.Text)
    End If
    If num2.Text = "e" Then
        b = 2.71828182845905
    ElseIf num2.Text = "��" Or num2.Text = "��" Or num2.Text = "pai" Or num2.Text = "pi" Then
        b = 3.14159265358979
    Else: b = Val(num2.Text)
    End If  '�������ֵ�Ƿ�Ϊe,��
    
    If cha.Text = "+" Then
        result.Text = a + b
    ElseIf cha.Text = "-" Then
        result.Text = a - b
    ElseIf cha.Text = "*" Or cha.Text = "��" Then
        result.Text = a * b
    ElseIf cha.Text = "/" Or cha.Text = "��" Then
        If b = 0 Then
        MsgBox "����������Ϊ0���� ���顰�ڶ����֡���", vbCritical, "����������(�� �㧥 ��;)��"  '���/0
        Else: result.Text = a / b '
        End If
    ElseIf cha.Text = "//" Then
        result.Text = Str(a \ b) & "..." & Str(a Mod b)
    ElseIf cha.Text = "��" Or cha.Text = "r" Then
        If a = 0 Then
        MsgBox "����������Ϊ0���� ���顰��һ���֡���", vbCritical, "����������(�� �㧥 ��;)��"  '���/0
        ElseIf num1.Text = "" Then  '���䷽������Ĭ��Ϊ��ƽ��
        result.Text = Sqr(b)
        Else: result.Text = b ^ (1 / a)
        End If
        
    ElseIf cha.Text = "^" Then
        result.Text = a ^ b
    ElseIf cha.Text = "log" Then
        result.Text = Log(b) / Log(a)   '�������㻻�׹�ʽ��b=������a=����
        
    ElseIf cha.Text = "o" Or cha.Text = "0" Or cha.Text = "��" Then
        result.Text = Str(a * (180 / 3.14159265358979)) & "��"
    ElseIf cha.Text = "rad" Then
        result.Text = Str(a * 3.14159265358979 / 180) & "rad"   '�ǶȻ��Ȼ���
        
    ElseIf cha.Text = "sin" Then
        result.Text = Sin(b * 3.14159265358979 / 180)
    ElseIf cha.Text = "cos" Then
        result.Text = Cos(b * 3.14159265358979 / 180)
    ElseIf cha.Text = "tan" Then
        result.Text = Tan(b * 3.14159265358979 / 180)   '���Ǻ���������ȫ���ǽǶ�
            If Str(Int(Str((b / 180) - 0.5))) = Str((b / 180) - 0.5) Then
                MsgBox "������ĽǶȲ���Ϊ90��+k��180��, k��Z �� ���顰�ڶ����֡���", vbCritical, "����������(�� �㧥 ��;)��"    '��������
            End If
    
    ElseIf cha.Text = "arcsin" Then
        If b < -1 Or b > 1 Then
            MsgBox "�������������С��-1�����1 �� ���顰�ڶ����֡���", vbCritical, "����������(�� �㧥 ��;)��"    '��������
        ElseIf b = 1 Then
            result.Text = "90��"
        ElseIf b = -1 Then
            result.Text = "-90��"
        Else
            result.Text = Str((Atn(b / Sqr(-b * b + 1))) * (180 / 3.14159265358979)) & "��"
        End If
    ElseIf cha.Text = "arccos" Then
        If b < -1 Or b > 1 Then
            MsgBox "�������������С��-1�����1 �� ���顰�ڶ����֡���", vbCritical, "����������(�� �㧥 ��;)��"    '��������
        ElseIf b = 1 Then
            result.Text = "0��"
        ElseIf b = -1 Then
            result.Text = "120��"
        Else
            result.Text = Str(Atn(Sqr(1 - b * b) / b) * (180 / 3.14159265358979)) & "��"
        End If
    ElseIf cha.Text = "arctan" Then
    result.Text = Str(Atn(b) * (180 / 3.14159265358979)) & "��"     '�����Ǻ�����#���Լ������
    
    
    
    
    End If  '����endif
End Sub

Private Sub clear_Click()   '���������
result.Text = ""
num1.Text = ""
num2.Text = ""
cha.Text = ""
End Sub

Private Sub cl_Click()
importt.Text = ""
sswrt.Text = ""

End Sub

Private Sub Frame4_Click()  '��������
Clipboard.SetText (result.Text)
End Sub

Private Sub Frame6_Click()  '��������
Clipboard.SetText (sswrt.Text)
End Sub

Private Sub Frame7_Click()  '��������
Clipboard.SetText (srfzt.Text)
End Sub

Private Sub Frame8_Click()  '��������
Clipboard.SetText (jwfzt.Text)
End Sub

Private Sub SSTab1_GotFocus()   '�л���ݼ���Tab��
'��ûд
End Sub

Private Sub import_Click()  '������
importt.Text = result.Text
End Sub

Private Sub sswr_Click()    '������������
ws = Val(weishu.Text)
sswrt.Text = Round(importt, ws)
End Sub


Private Sub weishu_LostFocus()  '���λ��0<ws<=16#��������-λ������ǰ������ʹ�ÿ�ѧ��������ʾ
    If ws > 16 Then
        MsgBox "������λ�����ܳ���16λ���� ���顰λ������", vbCritical, "����������(�� �㧥 ��;)��"
    ElseIf ws < 0 Then
        MsgBox "������λ������Ϊ�������� ���顰λ������", vbCritical, "����������(�� �㧥 ��;)��"
    End If
End Sub

Private Sub ����_Click()
MsgBox "��������ҳ��https://567z30m162.goho.co/text/sucalc.html���� ��������Ҫ������", vbInformation, "��i����ʾ��(o�b���b)o��"
Shell "explorer.exe https://567z30m162.goho.co/text/sucalc.html#help"
End Sub

Private Sub ���ڳ���������_Click()
MsgBox "��������ҳ��https://567z30m162.goho.co/text/sucalc.html���� ��������Ҫ������", vbInformation, "��i����ʾ��(o�b���b)o��"
Shell "explorer.exe https://567z30m162.goho.co/text/sucalc.html"
End Sub

Private Sub ������_Click()
MsgBox "��������ҳ��https://567z30m162.goho.co/text/sucalc.html���� ��������Ҫ������", vbInformation, "��i����ʾ��(o�b���b)o��"
Shell "explorer.exe https://567z30m162.goho.co/text/sucalc.html#checknew"
End Sub
