VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '���� ����
   Caption         =   "���� Ÿ�̸� by ryush00"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   960
   End
   Begin VB.CommandButton ssbtn 
      Caption         =   "Ÿ�̸� ����"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox s 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox m 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "��                  ��"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sval As Integer
Private Sub ssbtn_Click()
On Error GoTo error
If s.Text = "" And m.Text = "" Then GoTo error
If ssbtn.Caption = "Ÿ�̸� ����" Then
    If s.Text = "" Then s.Text = "0"
    If m.Text = "" Then m.Text = "0"
    If m.Text = "0" And m.Text = "0" Then GoTo error
    If s.Text <> "" And m.Text <> "" And s.Text >= 0 And m.Text >= 0 Then
        sval = s.Text + m.Text * 60
        ssbtn.Caption = "Ÿ�̸� ����"
        s.Enabled = False
        m.Enabled = False
        Timer1.Enabled = True
        Exit Sub
    End If
Else
    Timer1.Enabled = False
        s.Enabled = True
        m.Enabled = True
    ssbtn.Caption = "Ÿ�̸� ����"
    Exit Sub
End If
error:
MsgBox "������ �߻��Ͽ����ϴ�."
End Sub
Private Sub Timer1_Timer()
If sval = 0 Then
    Timer1.Enabled = False
    s.Enabled = True
    m.Enabled = True
    ssbtn.Caption = "Ÿ�̸� ����"
    MsgBox "Ÿ�̸� ����"
Else
sval = sval - 1
m.Text = Int(sval / 60)
s.Text = sval Mod 60
End If
End Sub
'ryush00 ����
