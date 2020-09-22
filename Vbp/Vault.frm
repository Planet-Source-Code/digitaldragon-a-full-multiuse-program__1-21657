VERSION 5.00
Begin VB.Form Vault 
   BorderStyle     =   0  'None
   Caption         =   "Vault"
   ClientHeight    =   7095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
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
   ScaleHeight     =   7095
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5535
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label LblPass 
      BackStyle       =   0  'Transparent
      Caption         =   "   Enable Pass"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   6405
      Width           =   1335
   End
   Begin VB.Image CmdPass 
      Height          =   330
      Left            =   360
      Top             =   6360
      Width           =   1425
   End
   Begin VB.Label LblReturn 
      BackStyle       =   0  'Transparent
      Caption         =   "       Return"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   6405
      Width           =   1335
   End
   Begin VB.Image CmdReturn 
      Height          =   330
      Left            =   2640
      Top             =   6360
      Width           =   1425
   End
   Begin VB.Image MinSmall 
      Height          =   225
      Left            =   3795
      Top             =   60
      Width           =   225
   End
   Begin VB.Image BorderDL 
      Height          =   375
      Left            =   0
      Top             =   6720
      Width           =   1080
   End
   Begin VB.Image BorderDR 
      Height          =   645
      Left            =   3770
      Top             =   6450
      Width           =   570
   End
   Begin VB.Image ExitSmall 
      Height          =   225
      Left            =   4035
      Top             =   60
      Width           =   225
   End
   Begin VB.Image DragImgR 
      Height          =   600
      Left            =   3600
      Top             =   0
      Width           =   735
   End
   Begin VB.Label DragLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "GreenDragon"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   300
      Width           =   1815
   End
   Begin VB.Image DragImgL 
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Image BorderR 
      Height          =   6255
      Left            =   4185
      Stretch         =   -1  'True
      Top             =   480
      Width           =   150
   End
   Begin VB.Image BorderD 
      Height          =   420
      Left            =   1065
      Stretch         =   -1  'True
      Top             =   6675
      Width           =   2715
   End
   Begin VB.Image BorderL 
      Height          =   6495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   300
   End
   Begin VB.Image DragImg 
      Height          =   600
      Left            =   480
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Vault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DragImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call DragForm(Me)
End Sub

Private Sub DragLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call DragForm(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open ".\Text Files\Vault.txt" For Output As 1
Dim v1
v1 = Text1.Text
Print #1, Text1.Text
Close
End Sub

Private Sub LblAdd_Click()
AddNew.Show
End Sub

Private Sub Form_Load()
Open ".\Text Files\Vault.txt" For Input As 1
Do Until EOF(1)
Line Input #1, redak
Text1.Text = Text1.Text + redak + vbCrLf
Loop
Close
DragImg.Picture = Skinz.DragImg.Picture
MinSmall.Picture = Skinz.MinUP.Picture
ExitSmall.Picture = Skinz.ExitUP.Picture
BorderL.Picture = Skinz.BorderL.Picture
BorderR.Picture = Skinz.BorderR.Picture
BorderD.Picture = Skinz.BorderD.Picture
CmdPass.Picture = Skinz.Bup.Picture
CmdReturn.Picture = Skinz.Bup.Picture
Me.Picture = Skinz.Back.Picture
BorderDL.Picture = Skinz.BorderDL.Picture
BorderDR.Picture = Skinz.BorderDR.Picture
DragImgL.Picture = Skinz.DragImgL.Picture
DragImgR.Picture = Skinz.DragImgR.Picture
End Sub
Private Sub LblReturn_Click()
Unload Me
End Sub

Private Sub LblReturn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdReturn.Picture = Skinz.Bdn.Picture
LblReturn.ForeColor = &H8000000E
End Sub

Private Sub LblReturn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdReturn.Picture = Skinz.Bup.Picture
LblReturn.ForeColor = &H80000012
End Sub

Private Sub Lblpass_Click()
If Main.Comm3.Caption = "0" Then
    LblPass.Caption = "   Disable pass"
    Pass.Show
    Pass.Text2.Visible = True
End If
If Main.Comm3.Caption = "1" Then
    Main.Comm3.Caption = "0"
    LblPass.Caption = "   Enable pass"
    Open ".\Text Files\Comm\Comm1.txt" For Output As 1
    Dim v1
    v1 = Main.Comm3.Caption
    Write #1, v1
    Close
End If
End Sub

Private Sub Lblpass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdPass.Picture = Skinz.Bdn.Picture
LblPass.ForeColor = &H8000000E
End Sub

Private Sub Lblpass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdPass.Picture = Skinz.Bup.Picture
LblPass.ForeColor = &H80000012
End Sub

