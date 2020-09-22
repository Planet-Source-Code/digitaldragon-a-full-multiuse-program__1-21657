VERSION 5.00
Begin VB.Form Testovi 
   BorderStyle     =   0  'None
   Caption         =   "Testovi"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
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
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   375
      ItemData        =   "Testovi.frx":0000
      Left            =   360
      List            =   "Testovi.frx":003D
      TabIndex        =   5
      Text            =   "Predmet:"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Text            =   "yyyy"
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Text            =   "mm"
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Text            =   "dd"
      Top             =   3360
      Width           =   375
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   2580
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   2115
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   2580
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label LblReturn 
      BackStyle       =   0  'Transparent
      Caption         =   " Return"
      Height          =   255
      Left            =   3060
      TabIndex        =   7
      Top             =   3875
      Width           =   735
   End
   Begin VB.Image CmdReturn 
      Height          =   330
      Left            =   2940
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   945
   End
   Begin VB.Image BorderDR 
      Height          =   645
      Left            =   3600
      Top             =   3810
      Width           =   570
   End
   Begin VB.Label DragLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Testovi"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   300
      Width           =   1815
   End
   Begin VB.Image ExitSmall 
      Height          =   225
      Left            =   3860
      Top             =   60
      Width           =   225
   End
   Begin VB.Image MinSmall 
      Height          =   225
      Left            =   3580
      Top             =   60
      Width           =   225
   End
   Begin VB.Image DragImgL 
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Label LblEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "   Edit"
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   3875
      Width           =   735
   End
   Begin VB.Image CmdEdit 
      Height          =   330
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   945
   End
   Begin VB.Image BorderR 
      Height          =   3255
      Left            =   4020
      Top             =   600
      Width           =   150
   End
   Begin VB.Image DragImgR 
      Height          =   600
      Left            =   3430
      Top             =   0
      Width           =   735
   End
   Begin VB.Image BorderD 
      Height          =   420
      Left            =   1065
      Stretch         =   -1  'True
      Top             =   4035
      Width           =   2545
   End
   Begin VB.Image BorderL 
      Height          =   3855
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
      Width           =   2970
   End
   Begin VB.Image BorderDL 
      Height          =   375
      Left            =   0
      Top             =   4080
      Width           =   1080
   End
   Begin VB.Label LblDelete 
      BackStyle       =   0  'Transparent
      Caption         =   " Delete"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   3395
      Width           =   735
   End
   Begin VB.Label LblNew 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New"
      Height          =   255
      Left            =   3060
      TabIndex        =   6
      Top             =   3395
      Width           =   735
   End
   Begin VB.Image CmdDelete 
      Height          =   330
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   945
   End
   Begin VB.Image CmdNew 
      Height          =   330
      Left            =   2940
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   945
   End
End
Attribute VB_Name = "Testovi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private n As Integer

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
Testovi.Picture = Skinz.Back.Picture
DragImgL.Picture = Skinz.DragImgL.Picture
DragImgR.Picture = Skinz.DragImgR.Picture
BorderDL.Picture = Skinz.BorderDL.Picture
BorderDR.Picture = Skinz.BorderDR.Picture
Open ".\Text Files\Testovi.txt" For Input As 1
Do Until EOF(1)
n = n + 1
Line Input #1, nextline
List1.List(n - 1) = nextline
Line Input #1, nextline
List2.List(n - 1) = nextline
Loop
Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open ".\Text Files\Testovi.txt" For Output As 1
For i = 1 To List1.ListCount
    Print #1, List1.List(i - 1)
    Print #1, List2.List(i - 1)
Next
Close
n = 0
End Sub

Private Sub LblDelete_Click()
If List1.ListIndex = -1 Then MsgBox ("You can't delete nothing!")
If List1.ListIndex > -1 Then
If List1.ListIndex > -1 Then List2.RemoveItem (List1.ListIndex)
If List1.ListIndex > -1 Then List1.RemoveItem (List1.ListIndex)
n = n + 1
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
If List1.ListIndex = -1 Then MsgBox ("You must select a test before you can edit it!")
If LblEdit.Caption = "   Edit" And List1.ListIndex > -1 Then
    LblEdit.Caption = "  Do It!"
    Text1.Text = Left(List1.Text, 2)
    Text2.Text = Mid(List1.Text, 6, 2)
    Text3.Text = Mid(List1.Text, 11, 4)
    Combo1.Text = List2.Text
ElseIf LblEdit.Caption = "  Do It!" Then
    LblEdit.Caption = "   Edit"
    List1.List(List1.ListIndex) = Text1.Text + " . " + Text2.Text + " . " + Text3.Text
    List2.List(List1.ListIndex) = Combo1.Text
    Text1.Text = "dd"
    Text2.Text = "mm"
    Text3.Text = "yyyy"
    Combo1.Text = "Predmet:"
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
If Combo1.Text = "Predmet:" Then MsgBox ("Èuj,prvo moraš odabrat predmet!")
If Not Combo1.Text = "Predmet:" Then
    List1.AddItem (Text1.Text + " . " + Text2.Text + " . " + Text3.Text)
    List2.AddItem (Combo1.Text)
    Text1.Text = "dd"
    Text2.Text = "mm"
    Text3.Text = "yyyy"
    Combo1.Text = "Predmet:"
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
List2.Selected(List1.ListIndex) = True
End Sub

Private Sub List2_Click()
List1.Selected(List2.ListIndex) = True
End Sub

Private Sub Text1_Click()
If Text1.Text = "dd" Then Text1.Text = ""
End Sub

Private Sub Text2_Click()
If Text2.Text = "mm" Then Text2.Text = ""
End Sub

Private Sub Text3_Click()
If Text3.Text = "yyyy" Then Text3.Text = ""
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

Private Sub MinSmall_Click()
Testovi.WindowState = 1
End Sub

Private Sub MinSmall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MinSmall.Picture = Skinz.MinDN.Picture
End Sub

Private Sub MinSmall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MinSmall.Picture = Skinz.MinUP.Picture
End Sub

