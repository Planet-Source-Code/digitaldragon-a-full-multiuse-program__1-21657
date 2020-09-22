VERSION 5.00
Begin VB.Form Vicevi 
   BorderStyle     =   0  'None
   Caption         =   "Raspored"
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Vicevi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000009&
      Height          =   4110
      Left            =   360
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Naslov 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Vic 
      Height          =   2415
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label DragLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Vicevi"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   300
      Width           =   1695
   End
   Begin VB.Label LblEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "         Edit"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   4355
      Width           =   1335
   End
   Begin VB.Image CmdEdit 
      Height          =   330
      Left            =   2520
      Top             =   4320
      Width           =   1425
   End
   Begin VB.Label LblReturn 
      BackStyle       =   0  'Transparent
      Caption         =   "       Return"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   4355
      Width           =   1335
   End
   Begin VB.Image CmdReturn 
      Height          =   330
      Left            =   4080
      Top             =   4320
      Width           =   1425
   End
   Begin VB.Image ExitSmall 
      Height          =   225
      Left            =   5540
      Top             =   60
      Width           =   225
   End
   Begin VB.Image MinSmall 
      Height          =   225
      Left            =   5265
      Top             =   60
      Width           =   225
   End
   Begin VB.Image DragImgL 
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Image BorderL 
      Height          =   4335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   300
   End
   Begin VB.Image BorderR 
      Height          =   3735
      Left            =   5700
      Stretch         =   -1  'True
      Top             =   600
      Width           =   150
   End
   Begin VB.Image BorderD 
      Height          =   420
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   4515
      Width           =   4210
   End
   Begin VB.Image DragImg 
      Height          =   600
      Left            =   495
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4630
   End
   Begin VB.Image BorderDR 
      Height          =   645
      Left            =   5280
      Top             =   4290
      Width           =   570
   End
   Begin VB.Image DragImgR 
      Height          =   600
      Left            =   5120
      Top             =   0
      Width           =   735
   End
   Begin VB.Image BorderDL 
      Height          =   375
      Left            =   0
      Top             =   4560
      Width           =   1080
   End
   Begin VB.Label LblVic 
      BackStyle       =   0  'Transparent
      Height          =   2895
      Left            =   2640
      TabIndex        =   6
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label LblNew 
      BackStyle       =   0  'Transparent
      Caption         =   "     Add New"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   3875
      Width           =   1335
   End
   Begin VB.Label LblDelete 
      BackStyle       =   0  'Transparent
      Caption         =   "       Delete"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   3875
      Width           =   1335
   End
   Begin VB.Image CmdDelete 
      Height          =   330
      Left            =   2520
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Image CmdNew 
      Height          =   330
      Left            =   4080
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   3135
      Left            =   2520
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Vicevi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private broj1(1000), n As Integer

Private Sub CmdDelete_Click()
Call LblDelete_Click
End Sub

Private Sub CmdDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblDelete_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblDelete_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdEdit_Click()
Call LblEdit_Click
End Sub

Private Sub CmdEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblEdit_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblEdit_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdNew_Click()
Call LblNew_Click
End Sub

Private Sub CmdNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblNew_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblNew_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdReturn_Click()
Call LblReturn_Click
End Sub

Private Sub CmdReturn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblReturn_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdReturn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblReturn_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub DragImg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call DragForm(Me)
End Sub

Private Sub DragLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call DragForm(Me)
End Sub

Private Sub Form_Load()
Vicevi.Picture = Skinz.Back.Picture
DragImg.Picture = Skinz.DragImg.Picture
ExitSmall.Picture = Skinz.ExitUP.Picture
MinSmall.Picture = Skinz.MinUP.Picture
BorderL.Picture = Skinz.BorderL
BorderR.Picture = Skinz.BorderR
BorderD.Picture = Skinz.BorderD
CmdNew.Picture = Skinz.Bup.Picture
CmdReturn.Picture = Skinz.Bup.Picture
CmdDelete.Picture = Skinz.Bup.Picture
CmdEdit.Picture = Skinz.Bup.Picture
DragImgL.Picture = Skinz.DragImgL.Picture
DragImgR.Picture = Skinz.DragImgR.Picture
BorderDL.Picture = Skinz.BorderDL.Picture
BorderDR.Picture = Skinz.BorderDR.Picture
Open ".\Text Files\Vicevi.txt" For Input As 1
Do Until EOF(1)
n = n + 1
Line Input #1, nextline
List1.List(n - 1) = nextline
Line Input #1, nextline
broj1(n) = nextline
Loop
Close
End Sub

Private Sub ExitSmall_Click()
Unload Me
End Sub

Private Sub ExitSmall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then ExitSmall.Picture = Skinz.ExitDN.Picture
End Sub

Private Sub ExitSmall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then ExitSmall.Picture = Skinz.ExitUP.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open ".\Text Files\Vicevi.txt" For Output As 1
For i = 1 To List1.ListCount
    Print #1, List1.List(i - 1)
    Print #1, broj1(i)
Next
Close
n = 0
End Sub

Private Sub LblDelete_Click()
If List1.ListIndex = -1 Then MsgBox ("You can't delete nothing!")
If List1.ListIndex > -1 Then
ListPosition = List1.ListIndex
If List1.ListIndex > -1 Then List1.RemoveItem (List1.ListIndex)
For i = ListPosition + 1 To List1.ListCount
broj1(i) = broj1(i + 1)
Next
n = n - 1
LblVic.Caption = ""
End If
End Sub

Private Sub LblDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdDelete.Picture = Skinz.Bdn.Picture
LblDelete.ForeColor = &H8000000E
End Sub

Private Sub LblDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdDelete.Picture = Skinz.Bup.Picture
LblDelete.ForeColor = &H80000012
End Sub

Private Sub LblEdit_Click()
If List1.ListIndex = -1 Then MsgBox ("You must select the Vic that you wish to edit first!")
If LblEdit.Caption = "        Edit" And List1.ListIndex > 0 Then
    LblEdit.Caption = "       Do It!"
    Vic.Visible = True
    Naslov.Visible = True
    Naslov = List1.Text
    Vic = LblVic.Caption
    LblVic.Caption = ""
ElseIf LblEdit.Caption = "       Do It!" Then
    LblEdit.Caption = "        Edit"
    Naslov.Visible = False
    Vic.Visible = False
    List1.List(List1.ListIndex) = Naslov.Text
    broj1(List1.ListIndex + 1) = Vic.Text
    Naslov.Text = ""
    Vic.Text = ""
End If
End Sub

Private Sub LblEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdEdit.Picture = Skinz.Bdn.Picture
LblEdit.ForeColor = &H8000000E
End Sub

Private Sub LblEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdEdit.Picture = Skinz.Bup.Picture
LblEdit.ForeColor = &H80000012
End Sub

Private Sub LblNew_Click()
If LblNew.Caption = "     Add New" Then
    LblNew.Caption = "       Do It!"
    Naslov.Visible = True
    Vic.Visible = True
    LblVic.Caption = ""
ElseIf LblNew.Caption = "       Do It!" Then
    LblNew.Caption = "     Add New"
    n = n + 1
    List1.AddItem (Naslov.Text)
    broj1(n) = Vic.Text
    Naslov.Visible = False
    Vic.Visible = False
    Vic.Text = ""
    Naslov.Text = ""
End If
End Sub

Private Sub LblNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdNew.Picture = Skinz.Bdn.Picture
LblNew.ForeColor = &H8000000E
End Sub

Private Sub LblNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdNew.Picture = Skinz.Bup.Picture
LblNew.ForeColor = &H80000012
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

Private Sub List1_Click()
LblVic.Caption = broj1(List1.ListIndex + 1)
End Sub

Private Sub MinSmall_Click()
Vicevi.WindowState = 1
End Sub

Private Sub MinSmall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MinSmall.Picture = Skinz.MinDN.Picture
End Sub

Private Sub MinSmall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MinSmall.Picture = Skinz.MinUP.Picture
End Sub
