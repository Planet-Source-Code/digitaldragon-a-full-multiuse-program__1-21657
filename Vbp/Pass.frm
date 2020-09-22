VERSION 5.00
Begin VB.Form Pass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Input"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2745
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   2745
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "Pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text2.Visible = False Then
    If Text1.Text = Label1.Caption Then
        Vault.Show
        Unload Me
    Else: MsgBox ("Daj pazi kaj pišeš!")
    End If
End If
If Text2.Visible = True Then
    If Text1.Text = Text2.Text Then
        Label1.Caption = Text1.Text
        Main.Comm3.Caption = "1"
        Open ".\Text Files\Comm\Password.txt" For Output As 1
        Dim v1
        v1 = Label1.Caption
        Write #1, v1
        Close
        Open ".\Text Files\Comm\Vault.txt" For Output As 2
        Dim c1
        c1 = Main.Comm3.Caption
        Write #2, c1
        Close
        Unload Me
    Else: MsgBox ("Daj pazi kaj pišeš!")
    End If
End If
End Sub

Private Sub Command2_Click()
If Text2.Visible = False Then Unload Me
If Text2.Visible = True Then
    Vault.LblPass.Caption = "   Enable pass"
    Main.Comm3.Caption = "0"
    Unload Me
End If
End Sub

Private Sub Form_Load()
Open ".\Text files\Comm\Password.txt" For Input As 1
Dim v1
Input #1, v1
Label1.Caption = v1
Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open ".\Text Files\Comm\Password.txt" For Output As 1
Dim v1
v1 = Label1.Caption
Write #1, v1
Close
End Sub
