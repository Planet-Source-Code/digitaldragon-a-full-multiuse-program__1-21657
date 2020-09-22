VERSION 5.00
Begin VB.Form Adresar 
   BorderStyle     =   0  'None
   Caption         =   "Adresar"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
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
   ScaleHeight     =   5415
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Naslov 
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
      Left            =   360
      TabIndex        =   13
      Text            =   "[Title]"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text5 
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
      Left            =   3360
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text4 
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
      Left            =   3360
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text3 
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
      Left            =   3360
      TabIndex        =   16
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
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
      Height          =   330
      Left            =   3360
      TabIndex        =   15
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
      Left            =   3360
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000009&
      Height          =   4110
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label LblMail 
      BackStyle       =   0  'Transparent
      Caption         =   "   Send Mail"
      Height          =   255
      Left            =   780
      TabIndex        =   24
      Top             =   4845
      Width           =   1215
   End
   Begin VB.Image CmdMail 
      Height          =   330
      Left            =   660
      Top             =   4800
      Width           =   1425
   End
   Begin VB.Label LblReturn 
      BackStyle       =   0  'Transparent
      Caption         =   "     Return"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   4835
      Width           =   1215
   End
   Begin VB.Image CmdReturn 
      Height          =   330
      Left            =   4080
      Top             =   4800
      Width           =   1425
   End
   Begin VB.Label LblEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "       Edit"
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   4835
      Width           =   1215
   End
   Begin VB.Image CmdEdit 
      Height          =   330
      Left            =   2520
      Top             =   4800
      Width           =   1425
   End
   Begin VB.Label DragLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Adresar"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   300
      Width           =   1695
   End
   Begin VB.Image BorderDL 
      Height          =   375
      Left            =   0
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Image BorderDR 
      Height          =   645
      Left            =   5220
      Top             =   4770
      Width           =   570
   End
   Begin VB.Image MinSmall 
      Height          =   225
      Left            =   5235
      Top             =   60
      Width           =   225
   End
   Begin VB.Image ExitSmall 
      Height          =   225
      Left            =   5475
      Top             =   60
      Width           =   225
   End
   Begin VB.Image DragImgR 
      Height          =   600
      Left            =   5060
      Top             =   0
      Width           =   735
   End
   Begin VB.Image DragImgL 
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Image BorderD 
      Height          =   420
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   5000
      Width           =   4155
   End
   Begin VB.Image DragImg 
      Height          =   600
      Left            =   490
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label LblNew 
      BackStyle       =   0  'Transparent
      Caption         =   "    Add New"
      Height          =   255
      Left            =   4200
      TabIndex        =   21
      Top             =   4355
      Width           =   1215
   End
   Begin VB.Label LblDelete 
      BackStyle       =   0  'Transparent
      Caption         =   "     Delete"
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   4355
      Width           =   1215
   End
   Begin VB.Label Komentar 
      Height          =   1575
      Left            =   3360
      TabIndex        =   12
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Telefon 
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Adresa 
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Mail 
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Prezime 
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Ime 
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label InfoLbl6 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label InfoLbl5 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefon"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label InfoLbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "Adresa"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label InfoLbl3 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label InfoLbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "Prezime"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label InfoLbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ime"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   3615
      Left            =   2520
      Top             =   600
      Width           =   3015
   End
   Begin VB.Image CmdDelete 
      Height          =   330
      Left            =   2520
      Top             =   4320
      Width           =   1425
   End
   Begin VB.Image CmdNew 
      Height          =   330
      Left            =   4080
      Top             =   4320
      Width           =   1425
   End
   Begin VB.Image BorderR 
      Height          =   4335
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   600
      Width           =   150
   End
   Begin VB.Image BorderL 
      Height          =   4815
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   300
   End
End
Attribute VB_Name = "Adresar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private broj1(1000), broj2(1000), broj3(1000), broj4(1000), broj5(1000), broj6(1000), n As Integer

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

Private Sub CmdMail_Click()
Call LblMail_Click
End Sub

Private Sub CmdMail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblMail_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdMail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblMail_MouseUp(Button, Shift, X, Y)
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

Private Sub Form_Load()
DragImg.Picture = Skinz.DragImg.Picture
MinSmall.Picture = Skinz.MinUP.Picture
ExitSmall.Picture = Skinz.ExitUP.Picture
BorderL.Picture = Skinz.BorderL.Picture
BorderR.Picture = Skinz.BorderR.Picture
BorderD.Picture = Skinz.BorderD.Picture
CmdReturn.Picture = Skinz.Bup.Picture
CmdNew.Picture = Skinz.Bup.Picture
CmdDelete.Picture = Skinz.Bup.Picture
CmdEdit.Picture = Skinz.Bup.Picture
CmdMail.Picture = Skinz.Bup.Picture
Adresar.Picture = Skinz.Back.Picture
DragImgL.Picture = Skinz.DragImgL.Picture
DragImgR.Picture = Skinz.DragImgR.Picture
BorderDL.Picture = Skinz.BorderDL.Picture
BorderDR.Picture = Skinz.BorderDR.Picture
Open ".\Text Files\Adresar.txt" For Input As 1
Do Until EOF(1)
n = n + 1
Line Input #1, nextline
List1.List(n - 1) = nextline
Line Input #1, nextline
broj1(n) = nextline
Line Input #1, nextline
broj2(n) = nextline
Line Input #1, nextline
broj3(n) = nextline
Line Input #1, nextline
broj4(n) = nextline
Line Input #1, nextline
broj5(n) = nextline
Line Input #1, nextline
broj6(n) = nextline
Loop
Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open ".\Text Files\Adresar.txt" For Output As 1
For i = 1 To List1.ListCount
    Print #1, List1.List(i - 1)
    Print #1, broj1(i)
    Print #1, broj2(i)
    Print #1, broj3(i)
    Print #1, broj4(i)
    Print #1, broj5(i)
    Print #1, broj6(i)
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
broj2(i) = broj2(i + 1)
broj3(i) = broj3(i + 1)
broj4(i) = broj4(i + 1)
broj5(i) = broj5(i + 1)
broj6(i) = broj6(i + 1)
Next
Ime.Caption = ""
Prezime.Caption = ""
Mail.Caption = ""
Adresa.Caption = ""
Telefon.Caption = ""
Komentar.Caption = ""
n = n - 1
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
If List1.ListIndex = -1 Then MsgBox ("You must select the Vic that you wish to edit first!")
If LblEdit.Caption = "       Edit" And List1.ListIndex > -1 Then
    LblEdit.Caption = "      Do It!"
    Text1.Visible = True
    Text2.Visible = True
    Text3.Visible = True
    Text4.Visible = True
    Text5.Visible = True
    Text6.Visible = True
    Naslov.Visible = True
    Naslov.Text = List1.Text
    Text1.Text = Ime.Caption
    Text2.Text = Prezime.Caption
    Text3.Text = Mail.Caption
    Text4.Text = Adresa.Caption
    Text5.Text = Telefon.Caption
    Text6.Text = Komentar.Caption
ElseIf LblEdit.Caption = "      Do It!" Then
    LblEdit.Caption = "       Edit"
    Naslov.Visible = False
    Text1.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Text5.Visible = False
    Text6.Visible = False
    List1.List(List1.ListIndex) = Naslov.Text
    broj1(List1.ListIndex + 1) = Text1.Text
    broj2(List1.ListIndex + 1) = Text2.Text
    broj3(List1.ListIndex + 1) = Text3.Text
    broj4(List1.ListIndex + 1) = Text4.Text
    broj5(List1.ListIndex + 1) = Text5.Text
    broj6(List1.ListIndex + 1) = Text6.Text
    Naslov.Text = "[Title]"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
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

Private Sub LblMail_Click()
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
sTo = Mail.Caption
sCC = ""
sBCC = ""
sSubject = "(No subject)"
sBody = ""
If List1.ListIndex = -1 Then MsgBox ("You first must select a person to send him/her E-Mail!") Else ret = Shell("Start.exe " _
        & "mailto:" & """" & sTo & """" _
        & "?Subject=" & """" & sSubject & """" _
        & "&cc=" & """" & sCC & """" _
        & "&bcc=" & """" & sBCC & """" _
        & "&Body=" & """" & sBody & """" _
        & "&File=" & """" & "c:\autoexec.bat" & """" _
        , 0)
End Sub

Private Sub LblMail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdMail.Picture = Skinz.Bdn.Picture
LblMail.ForeColor = &H8000000E
End Sub

Private Sub LblMail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdMail.Picture = Skinz.Bup.Picture
LblMail.ForeColor = &H80000012
End Sub

Private Sub LblNew_Click()
If LblNew.Caption = "    Add New" Then
    LblNew.Caption = "      Do It!"
    Naslov.Visible = True
    Text1.Visible = True
    Text2.Visible = True
    Text3.Visible = True
    Text4.Visible = True
    Text5.Visible = True
    Text6.Visible = True
ElseIf LblNew.Caption = "      Do It!" Then
    n = n + 1
    LblNew.Caption = "    Add New"
    List1.AddItem (Naslov.Text)
    broj1(n) = Text1.Text
    broj2(n) = Text2.Text
    broj3(n) = Text3.Text
    broj4(n) = Text4.Text
    broj5(n) = Text5.Text
    broj6(n) = Text6.Text
    Naslov.Visible = False
    Text1.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Text5.Visible = False
    Text6.Visible = False
    Naslov.Text = "[Title]"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
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
Ime.Caption = broj1(List1.ListIndex + 1)
Prezime.Caption = broj2(List1.ListIndex + 1)
Mail.Caption = broj3(List1.ListIndex + 1)
Adresa.Caption = broj4(List1.ListIndex + 1)
Telefon.Caption = broj5(List1.ListIndex + 1)
Komentar.Caption = broj6(List1.ListIndex + 1)
End Sub

Private Sub MinSmall_Click()
Adresar.WindowState = 1
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
