VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "GreenDragon"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Comm3 
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   4200
      Width           =   255
   End
   Begin VB.CommandButton Comm1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   255
   End
   Begin VB.CommandButton Comm2 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image MainImg 
      Height          =   2550
      Left            =   1920
      Top             =   720
      Width           =   1650
   End
   Begin VB.Image MinSmall 
      Height          =   225
      Left            =   3200
      Top             =   60
      Width           =   225
   End
   Begin VB.Image ExitSmall 
      Height          =   225
      Left            =   3440
      Top             =   60
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "GreenDragon"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   300
      Width           =   1815
   End
   Begin VB.Image DragImgL 
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Image BorderDL 
      Height          =   375
      Left            =   0
      Top             =   3600
      Width           =   1080
   End
   Begin VB.Label LblExit 
      BackStyle       =   0  'Transparent
      Caption         =   "         Exit"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   3395
      Width           =   1335
   End
   Begin VB.Image CmdExit 
      Height          =   330
      Left            =   2040
      Top             =   3360
      Width           =   1425
   End
   Begin VB.Image BorderDR 
      Height          =   645
      Left            =   3165
      Top             =   3330
      Width           =   570
   End
   Begin VB.Label LblSalac 
      BackStyle       =   0  'Transparent
      Caption         =   "      Options"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3275
      Width           =   1335
   End
   Begin VB.Image CmdSalac 
      Height          =   330
      Left            =   360
      Top             =   3240
      Width           =   1425
   End
   Begin VB.Label LblTestovi 
      BackStyle       =   0  'Transparent
      Caption         =   "        Tests"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2915
      Width           =   1335
   End
   Begin VB.Label LblVicevi 
      BackStyle       =   0  'Transparent
      Caption         =   "        Jokes"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2555
      Width           =   1335
   End
   Begin VB.Label LblInternet 
      BackStyle       =   0  'Transparent
      Caption         =   "     Internet"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2195
      Width           =   1335
   End
   Begin VB.Label LblAdresar 
      BackStyle       =   0  'Transparent
      Caption         =   "     Addresses"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1835
      Width           =   1335
   End
   Begin VB.Label LblKemija 
      BackStyle       =   0  'Transparent
      Caption         =   "        To do"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label LblOcjene 
      BackStyle       =   0  'Transparent
      Caption         =   "       Grades"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1115
      Width           =   1335
   End
   Begin VB.Label LblRaspored 
      BackStyle       =   0  'Transparent
      Caption         =   "      Shedule"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   755
      Width           =   1335
   End
   Begin VB.Image DragImg 
      Height          =   600
      Left            =   480
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2535
   End
   Begin VB.Image BorderL 
      Height          =   3735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   300
   End
   Begin VB.Image CmdRaspored 
      Height          =   330
      Left            =   360
      Top             =   720
      Width           =   1425
   End
   Begin VB.Image CmdOcjene 
      Height          =   330
      Left            =   360
      Top             =   1080
      Width           =   1425
   End
   Begin VB.Image CmdKemija 
      Height          =   330
      Left            =   360
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Image CmdAdresar 
      Height          =   330
      Left            =   360
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Image CmdInternet 
      Height          =   330
      Left            =   360
      Top             =   2160
      Width           =   1425
   End
   Begin VB.Image CmdVicevi 
      Height          =   330
      Left            =   360
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Image CmdTestovi 
      Height          =   330
      Left            =   360
      Top             =   2880
      Width           =   1425
   End
   Begin VB.Image BorderD 
      Height          =   420
      Left            =   1065
      Stretch         =   -1  'True
      Top             =   3555
      Width           =   2110
   End
   Begin VB.Image DragImgR 
      Height          =   600
      Left            =   3000
      Top             =   0
      Width           =   735
   End
   Begin VB.Image BorderR 
      Height          =   3735
      Left            =   3590
      Top             =   240
      Width           =   150
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdresar_Click()
Call LblAdresar_Click
End Sub

Private Sub CmdAdresar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblAdresar_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdAdresar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblAdresar_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdInternet_Click()
Call LblInternet_Click
End Sub

Private Sub CmdInternet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblInternet_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdInternet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblInternet_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdKemija_Click()
Call LblKemija_Click
End Sub

Private Sub CmdKemija_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblKemija_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdKemija_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblKemija_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdOcjene_Click()
Call LblOcjene_Click
End Sub

Private Sub CmdOcjene_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblOcjene_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdOcjene_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblOcjene_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdRaspored_Click()
Call LblRaspored_Click
End Sub

Private Sub CmdRaspored_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblRaspored_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdRaspored_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblRaspored_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdSalac_Click()
Call LblSalac_Click
End Sub

Private Sub CmdSalac_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblSalac_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdSalac_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblSalac_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdTestovi_Click()
Call LblTestovi_Click
End Sub

Private Sub CmdTestovi_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblTestovi_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdTestovi_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblTestovi_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdVicevi_Click()
Call LblVicevi_Click
End Sub

Private Sub CmdVicevi_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblVicevi_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdVicevi_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblVicevi_MouseUp(Button, Shift, X, Y)
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
MainImg.Picture = LoadPicture(".\Skins\Main.jpg")
Skinz.Back.Picture = LoadPicture(".\Skins\Back.jpg")
Skinz.Bdn.Picture = LoadPicture(".\Skins\Bdn.jpg")
Skinz.Bup.Picture = LoadPicture(".\Skins\Bup.jpg")
Skinz.BorderL.Picture = LoadPicture(".\Skins\BorderL.jpg")
Skinz.BorderR.Picture = LoadPicture(".\Skins\BorderR.jpg")
Skinz.ExitUP.Picture = LoadPicture(".\Skins\ExitUP.jpg")
Skinz.ExitDN.Picture = LoadPicture(".\Skins\ExitDN.jpg")
Skinz.MinUP.Picture = LoadPicture(".\Skins\MinUP.jpg")
Skinz.MinDN.Picture = LoadPicture(".\Skins\MinDN.jpg")
Skinz.DragImg.Picture = LoadPicture(".\Skins\DragImgC.jpg")
Skinz.DragImgL.Picture = LoadPicture(".\Skins\DragImgL.jpg")
Skinz.DragImgR.Picture = LoadPicture(".\Skins\DragImgR.jpg")
Skinz.BorderD.Picture = LoadPicture(".\Skins\BorderD.jpg")
Skinz.BorderDL.Picture = LoadPicture(".\Skins\BorderDL.jpg")
Skinz.BorderDR.Picture = LoadPicture(".\Skins\BorderDR.jpg")
DragImg.Picture = Skinz.DragImg.Picture
MinSmall.Picture = Skinz.MinUP.Picture
ExitSmall.Picture = Skinz.ExitUP.Picture
BorderL.Picture = Skinz.BorderL.Picture
BorderR.Picture = Skinz.BorderR.Picture
BorderD.Picture = Skinz.BorderD.Picture
CmdRaspored.Picture = Skinz.Bup.Picture
CmdOcjene.Picture = Skinz.Bup.Picture
CmdKemija.Picture = Skinz.Bup.Picture
CmdAdresar.Picture = Skinz.Bup.Picture
CmdInternet.Picture = Skinz.Bup.Picture
CmdVicevi.Picture = Skinz.Bup.Picture
CmdTestovi.Picture = Skinz.Bup.Picture
CmdSalac.Picture = Skinz.Bup.Picture
CmdExit.Picture = Skinz.Bup.Picture
Main.Picture = Skinz.Back.Picture
BorderDL.Picture = Skinz.BorderDL.Picture
BorderDR.Picture = Skinz.BorderDR.Picture
DragImgL.Picture = Skinz.DragImgL.Picture
DragImgR.Picture = Skinz.DragImgR.Picture
Open ".\Text Files\Comm\Vault.txt" For Input As 3
Dim AssHole
Input #3, AssHole
Comm3.Caption = AssHole
Close
Open ".\Text Files\Comm\Ocjene.txt" For Input As 1
Dim v1
Input #1, v1
Comm1.Caption = v1
Close
Open ".\Text Files\Comm\Raspored.txt" For Input As 2
Dim c1
Input #2, c1
Comm2.Caption = c1
Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Skinz
Unload Vicevi
Unload ToDo
Unload Adresar
Unload Internet
Unload Testovi
Unload School
Unload Raspored
Unload Vault
Open ".\Text Files\Comm\Ocjene.txt" For Output As 1
Dim v1
v1 = Comm1.Caption
Write #1, v1
Close
Open ".\Text Files\Comm\Raspored.txt" For Output As 2
Dim c1
c1 = Comm2.Caption
Write #2, c1
Close
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Call DragForm(Me)
End Sub

Private Sub LblAdresar_Click()
Adresar.Show
End Sub

Private Sub LblAdresar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdAdresar.Picture = Skinz.Bdn.Picture
LblAdresar.ForeColor = &H8000000E
End Sub

Private Sub LblAdresar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdAdresar.Picture = Skinz.Bup.Picture
LblAdresar.ForeColor = &H80000012
End Sub

Private Sub LblExit_Click()
Unload Me
End Sub

Private Sub LblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdExit.Picture = Skinz.Bdn.Picture
LblExit.ForeColor = &H8000000E
End Sub

Private Sub LblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdExit.Picture = Skinz.Bup.Picture
LblExit.ForeColor = &H80000012
End Sub

Private Sub LblInternet_Click()
Internet.Show
End Sub

Private Sub LblInternet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdInternet.Picture = Skinz.Bdn.Picture
LblInternet.ForeColor = &H8000000E
End Sub

Private Sub LblInternet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdInternet.Picture = Skinz.Bup.Picture
LblInternet.ForeColor = &H80000012
End Sub

Private Sub LblKemija_Click()
MsgBox ("Under Construction")
End Sub

Private Sub LblKemija_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdKemija.Picture = Skinz.Bdn.Picture
LblKemija.ForeColor = &H8000000E
End Sub

Private Sub LblKemija_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdKemija.Picture = Skinz.Bup.Picture
LblKemija.ForeColor = &H80000012
End Sub

Private Sub LblOcjene_Click()
School.Show
End Sub

Private Sub LblOcjene_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdOcjene.Picture = Skinz.Bdn.Picture
LblOcjene.ForeColor = &H8000000E
End Sub

Private Sub LblOcjene_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdOcjene.Picture = Skinz.Bup.Picture
LblOcjene.ForeColor = &H80000012
End Sub

Private Sub LblRaspored_Click()
Raspored.Show
End Sub

Private Sub LblRaspored_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdRaspored.Picture = Skinz.Bdn.Picture
LblRaspored.ForeColor = &H8000000E
End Sub

Private Sub LblRaspored_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdRaspored.Picture = Skinz.Bup.Picture
LblRaspored.ForeColor = &H80000012
End Sub

Private Sub LblSalac_Click()
MsgBox ("Under construction")
End Sub

Private Sub LblSalac_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdSalac.Picture = Skinz.Bdn.Picture
LblSalac.ForeColor = &H8000000E
End Sub

Private Sub LblSalac_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdSalac.Picture = Skinz.Bup.Picture
LblSalac.ForeColor = &H80000012
End Sub

Private Sub LblTestovi_Click()
Testovi.Show
End Sub

Private Sub LblTestovi_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdTestovi.Picture = Skinz.Bdn.Picture
LblTestovi.ForeColor = &H8000000E
End Sub

Private Sub LblTestovi_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdTestovi.Picture = Skinz.Bup.Picture
LblTestovi.ForeColor = &H80000012
End Sub

Private Sub LblVicevi_Click()
Vicevi.Show
End Sub

Private Sub LblVicevi_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdVicevi.Picture = Skinz.Bdn.Picture
LblVicevi.ForeColor = &H8000000E
End Sub

Private Sub LblVicevi_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdVicevi.Picture = Skinz.Bup.Picture
LblVicevi.ForeColor = &H80000012
End Sub

Private Sub MainImg_Click()
If Comm3.Caption = "0" Then
    Vault.Show
    Vault.LblPass.Caption = "   Enable pass"
ElseIf Comm3.Caption = "1" Then
    Pass.Show
    Vault.LblPass.Caption = "   Disable pass"
End If
End Sub

Private Sub MinSmall_Click()
Main.WindowState = 1
End Sub

Private Sub MinSmall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MinSmall.Picture = Skinz.MinDN.Picture
End Sub

Private Sub MinSmall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MinSmall.Picture = Skinz.MinUP.Picture
End Sub
