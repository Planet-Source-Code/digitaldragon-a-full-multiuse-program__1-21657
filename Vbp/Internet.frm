VERSION 5.00
Begin VB.Form Internet 
   BorderStyle     =   0  'None
   Caption         =   "Internet"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
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
   ScaleHeight     =   4455
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Naslov 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Text            =   "[Title]"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3240
      TabIndex        =   6
      Text            =   "[Address]"
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   3600
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label DragLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   300
      Width           =   1695
   End
   Begin VB.Label LblReturn 
      BackStyle       =   0  'Transparent
      Caption         =   "     Return"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   3875
      Width           =   1215
   End
   Begin VB.Image CmdReturn 
      Height          =   330
      Left            =   3960
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Image BorderDL 
      Height          =   375
      Left            =   0
      Top             =   4080
      Width           =   1080
   End
   Begin VB.Image DragImgL 
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Image MinSmall 
      Height          =   225
      Left            =   5115
      Top             =   60
      Width           =   225
   End
   Begin VB.Image ExitSmall 
      Height          =   225
      Left            =   5355
      Top             =   60
      Width           =   225
   End
   Begin VB.Image DragImgR 
      Height          =   600
      Left            =   4935
      Top             =   0
      Width           =   735
   End
   Begin VB.Image BorderDR 
      Height          =   645
      Left            =   5100
      Top             =   3810
      Width           =   570
   End
   Begin VB.Label LblEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "       Edit"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   3875
      Width           =   1215
   End
   Begin VB.Image CmdEdit 
      Height          =   330
      Left            =   2400
      Top             =   3840
      Width           =   1425
   End
   Begin VB.Image DragImg 
      Height          =   600
      Left            =   495
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4440
   End
   Begin VB.Image BorderD 
      Height          =   420
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   4040
      Width           =   4050
   End
   Begin VB.Label LblNew 
      BackStyle       =   0  'Transparent
      Caption         =   "    Add New"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   3395
      Width           =   1215
   End
   Begin VB.Label LblDelete 
      BackStyle       =   0  'Transparent
      Caption         =   "     Delete"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   3395
      Width           =   1215
   End
   Begin VB.Label Komentar 
      Height          =   2055
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Adresa 
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label InfoLbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label InfoLbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Adress"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Left            =   2400
      Top             =   600
      Width           =   3015
   End
   Begin VB.Image CmdDelete 
      Height          =   330
      Left            =   2400
      Top             =   3360
      Width           =   1425
   End
   Begin VB.Image CmdNew 
      Height          =   330
      Left            =   3960
      Top             =   3360
      Width           =   1425
   End
   Begin VB.Image BorderR 
      Height          =   3255
      Left            =   5520
      Stretch         =   -1  'True
      Top             =   600
      Width           =   150
   End
   Begin VB.Image BorderL 
      Height          =   3855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   300
   End
End
Attribute VB_Name = "Internet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private broj1(1000), broj2(1000), n As Integer

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

Private Sub DragImgL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call DragForm(Me)
End Sub

Private Sub DragImgR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call DragForm(Me)
End Sub

Private Sub DragLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call DragForm(Me)
End Sub

Private Sub MinSmall_Click()
Internet.WindowState = 1
End Sub

Private Sub MinSmall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MinSmall.Picture = Skinz.MinDN.Picture
End Sub

Private Sub MinSmall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MinSmall.Picture = Skinz.MinUP.Picture
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

Private Sub Form_Load()
Open ".\Text Files\Internet.txt" For Input As 1
Do Until EOF(1)
n = n + 1
Line Input #1, nextline
List1.List(n - 1) = nextline
Line Input #1, nextline
broj1(n) = nextline
Line Input #1, nextline
broj2(n) = nextline
Loop
Close
DragImg.Picture = Skinz.DragImg.Picture
MinSmall.Picture = Skinz.MinUP.Picture
ExitSmall.Picture = Skinz.ExitUP.Picture
BorderL.Picture = Skinz.BorderL.Picture
BorderR.Picture = Skinz.BorderR.Picture
BorderD.Picture = Skinz.BorderD.Picture
CmdNew.Picture = Skinz.Bup.Picture
CmdReturn.Picture = Skinz.Bup.Picture
CmdDelete.Picture = Skinz.Bup.Picture
CmdEdit.Picture = Skinz.Bup.Picture
Internet.Picture = Skinz.Back.Picture
DragImgL.Picture = Skinz.DragImgL.Picture
DragImgR.Picture = Skinz.DragImgR.Picture
BorderDL.Picture = Skinz.BorderDL.Picture
BorderDR.Picture = Skinz.BorderDR.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open ".\Text Files\Internet.txt" For Output As 1
For i = 1 To List1.ListCount
    Print #1, List1.List(i - 1)
    Print #1, broj1(i)
    Print #1, broj2(i)
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
broj2(i) = broj1(i + 1)
Next
n = n - 1
Adresa.Caption = ""
Komentar.Caption = ""
End If
End Sub

Private Sub LblDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdDelete.Picture = Skinz.Bdn.Picture
LblDelete.ForeColor = &H8000000E
End Sub

Private Sub LblDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdDelete.Picture = Skinz.Bup.Picture
LblDelete.ForeColor = &H80000012

End Sub

Private Sub LblEdit_Click()
If List1.ListIndex = -1 Then MsgBox ("In order to edit one of the internet adresses,first you must select one!")
If LblEdit.Caption = "       Edit" And List1.ListIndex > -1 Then
    LblEdit.Caption = "      Do It!"
    Text1.Visible = True
    Text2.Visible = True
    Naslov.Visible = True
    Naslov.Text = List1.Text
    Text1.Text = Adresa.Caption
    Text2.Text = Komentar.Caption
ElseIf LblEdit.Caption = "      Do It!" Then
    LblEdit.Caption = "       Edit"
    Naslov.Visible = False
    Text1.Visible = False
    Text2.Visible = False
    List1.List(List1.ListIndex) = Naslov.Text
    broj1(List1.ListIndex + 1) = Text1.Text
    broj2(List1.ListIndex + 1) = Text2.Text
    Naslov.Text = "[Title]"
    Text1.Text = "[Address]"
    Text2.Text = ""
End If
End Sub

Private Sub LblEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdEdit.Picture = Skinz.Bdn.Picture
LblEdit.ForeColor = &H8000000E
End Sub

Private Sub LblEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdEdit.Picture = Skinz.Bup.Picture
LblEdit.ForeColor = &H80000012
End Sub

Private Sub LblNew_Click()
If LblNew.Caption = "    Add New" Then
    LblNew.Caption = "      Do It!"
    Naslov.Visible = True
    Text1.Visible = True
    Text2.Visible = True
ElseIf LblNew.Caption = "      Do It!" Then
    LblNew.Caption = "    Add New"
    Naslov.Visible = False
    Text1.Visible = False
    Text2.Visible = False
    List1.AddItem (Naslov.Text)
    n = n + 1
    broj1(n) = Text1.Text
    broj2(n) = Text2.Text
    Naslov.Text = "[Title]"
    Text1.Text = "[Address]"
    Text2.Text = ""
End If
End Sub

Private Sub LblNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdNew.Picture = Skinz.Bdn.Picture
LblNew.ForeColor = &H8000000E
End Sub

Private Sub LblNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdNew.Picture = Skinz.Bup.Picture
LblNew.ForeColor = &H80000012
End Sub

Private Sub LblReturn_Click()
Unload Me
End Sub

Private Sub LblReturn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdReturn.Picture = Skinz.Bdn.Picture
LblReturn.ForeColor = &H8000000E
End Sub

Private Sub LblReturn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdReturn.Picture = Skinz.Bup.Picture
LblReturn.ForeColor = &H80000012
End Sub

Private Sub List1_Click()
Adresa.Caption = broj1(List1.ListIndex + 1)
Komentar.Caption = broj2(List1.ListIndex + 1)
End Sub

Private Sub Naslov_Click()
If Naslov.Text = "[Title]" Then Naslov.Text = ""
End Sub

Private Sub Text1_Click()
If Text1.Text = "[Address]" Then Text1.Text = ""
End Sub
