VERSION 5.00
Begin VB.Form School 
   BorderStyle     =   0  'None
   Caption         =   "Ocjene"
   ClientHeight    =   4695
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
   ScaleHeight     =   4695
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   375
      ItemData        =   "Ocjene.frx":0000
      Left            =   2040
      List            =   "Ocjene.frx":0013
      TabIndex        =   11
      Text            =   "5"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   3855
      ItemData        =   "Ocjene.frx":0026
      Left            =   360
      List            =   "Ocjene.frx":0057
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Image ExitSmall 
      Height          =   225
      Left            =   5475
      Top             =   60
      Width           =   225
   End
   Begin VB.Image MinSmall 
      Height          =   225
      Left            =   5200
      Top             =   60
      Width           =   225
   End
   Begin VB.Label DragLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Ocjene"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   52
      Top             =   300
      Width           =   1695
   End
   Begin VB.Image DragImgL 
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Image BorderL 
      Height          =   4095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   300
   End
   Begin VB.Image BorderR 
      Height          =   3495
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   600
      Width           =   150
   End
   Begin VB.Image DragImg 
      Height          =   600
      Left            =   490
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image BorderD 
      Height          =   420
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   4275
      Width           =   4155
   End
   Begin VB.Image DragImgR 
      Height          =   600
      Left            =   5060
      Top             =   0
      Width           =   735
   End
   Begin VB.Image BorderDR 
      Height          =   645
      Left            =   5220
      Top             =   4050
      Width           =   570
   End
   Begin VB.Image BorderDL 
      Height          =   375
      Left            =   0
      Top             =   4320
      Width           =   1080
   End
   Begin VB.Label DeleteGN2 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   3600
      TabIndex        =   51
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label DeleteGN1 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   3600
      TabIndex        =   50
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Tester2 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   49
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Tester1 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   48
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label Calculator2 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   47
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Calculator1 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   46
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   29
      Left            =   2040
      TabIndex        =   45
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   28
      Left            =   2040
      TabIndex        =   44
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   27
      Left            =   2040
      TabIndex        =   43
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   26
      Left            =   2040
      TabIndex        =   42
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   25
      Left            =   2040
      TabIndex        =   41
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   24
      Left            =   1680
      TabIndex        =   40
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   23
      Left            =   1680
      TabIndex        =   39
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   22
      Left            =   1680
      TabIndex        =   38
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   21
      Left            =   1680
      TabIndex        =   37
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   20
      Left            =   1680
      TabIndex        =   36
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   19
      Left            =   1320
      TabIndex        =   35
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   18
      Left            =   1320
      TabIndex        =   34
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   17
      Left            =   1320
      TabIndex        =   33
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   16
      Left            =   1320
      TabIndex        =   32
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000006&
      Height          =   255
      Index           =   15
      Left            =   1320
      TabIndex        =   31
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   14
      Left            =   960
      TabIndex        =   30
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   13
      Left            =   960
      TabIndex        =   29
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   12
      Left            =   960
      TabIndex        =   28
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   11
      Left            =   960
      TabIndex        =   27
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   10
      Left            =   960
      TabIndex        =   26
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   9
      Left            =   600
      TabIndex        =   25
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   24
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   23
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   22
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   21
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   20
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   18
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Ocjene 
      BackColor       =   &H80000001&
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label LblReturn 
      BackStyle       =   0  'Transparent
      Caption         =   "      Return"
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   3645
      Width           =   1215
   End
   Begin VB.Label LblDelete 
      BackStyle       =   0  'Transparent
      Caption         =   "      Delete"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   3645
      Width           =   1215
   End
   Begin VB.Label LblAdd 
      BackStyle       =   0  'Transparent
      Caption         =   "    Add Mark"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   3045
      Width           =   1215
   End
   Begin VB.Label LblPolugo 
      BackStyle       =   0  'Transparent
      Caption         =   " 2.Polugodište"
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   2205
      Width           =   1215
   End
   Begin VB.Image CmdPolugo 
      Height          =   330
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1185
   End
   Begin VB.Image CmdDelete 
      Height          =   330
      Left            =   1920
      Top             =   3600
      Width           =   1425
   End
   Begin VB.Image CmdReturn 
      Height          =   330
      Left            =   3480
      Top             =   3600
      Width           =   1425
   End
   Begin VB.Image CmdAdd 
      Height          =   330
      Left            =   3480
      Top             =   3000
      Width           =   1425
   End
   Begin VB.Label FrameLbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add-New"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   2745
      Width           =   735
   End
   Begin VB.Label FrameLbl 
      BackColor       =   &H0080C0FF&
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   1920
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Finale2 
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Finale1 
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2055
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Ocjene2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Ocjene1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Predmet2 
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4455
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Predmet1 
      BackStyle       =   0  'Transparent
      Caption         =   "XX"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4455
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Line Line2 
      X1              =   4440
      X2              =   4440
      Y1              =   720
      Y2              =   1800
   End
   Begin VB.Line Line1 
      X1              =   3240
      X2              =   3240
      Y1              =   720
      Y2              =   1800
   End
   Begin VB.Label InfoLbl2 
      BackStyle       =   0  'Transparent
      Caption         =   "2.Polugodište"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label InfoLbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "1.Polugodište"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Shape Shape5 
      Height          =   615
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2055
      Left            =   1920
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "School"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private broj1(30), n, gn, polugo As Integer

Private Sub CmdAdd_Click()
Call LblAdd_Click
End Sub

Private Sub CmdAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblAdd_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblAdd_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdDelete_Click()
Call LblDelete_Click
End Sub

Private Sub CmdDelete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblDelete_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdDelete_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblDelete_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdPolugo_Click()
Call LblPolugo_Click
End Sub

Private Sub CmdPolugo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblPolugo_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdPolugo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblPolugo_MouseUp(Button, Shift, X, Y)
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
ExitSmall.Picture = Skinz.ExitUP.Picture
MinSmall.Picture = Skinz.MinUP.Picture
BorderL.Picture = Skinz.BorderL.Picture
BorderR.Picture = Skinz.BorderR.Picture
BorderD.Picture = Skinz.BorderD.Picture
CmdPolugo.Picture = Skinz.Bup.Picture
CmdAdd.Picture = Skinz.Bup.Picture
CmdReturn.Picture = Skinz.Bup.Picture
CmdDelete.Picture = Skinz.Bup.Picture
School.Picture = Skinz.Back.Picture
DragImgL.Picture = Skinz.DragImgL.Picture
DragImgR.Picture = Skinz.DragImgR.Picture
BorderDL.Picture = Skinz.BorderDL.Picture
BorderDR.Picture = Skinz.BorderDR.Picture
Open ".\Text Files\Ocjene.txt" For Input As 1
Do Until EOF(1)
n = n + 1
Line Input #1, nextline
Ocjene(n - 1) = nextline
Line Input #1, nextline
broj1(n) = nextline
Loop
Close
n = 0
If Main.Comm1.Caption = "1" Then
    polugo = 0
    LblPolugo.Caption = " 2.Polugodište"
ElseIf Main.Comm1.Caption = "2" Then
    polugo = 1
    LblPolugo.Caption = " 1.Polugodište"
End If
Call Calculate(X)
Call ApplyLabels(X)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open ".\Text Files\Ocjene.txt" For Output As 1
For i = 1 To 30
Print #1, Ocjene(i - 1)
Print #1, broj1(i)
Next
Close
End Sub

Private Sub LblAdd_Click()
If List1.ListIndex > -1 Then
If polugo = 0 And Not Ocjene1.Caption = "" Then broj1(List1.ListIndex + 1) = broj1(List1.ListIndex + 1) + "," + Combo1.Text
If polugo = 1 And Not Ocjene2.Caption = "" Then broj1(List1.ListIndex + 16) = broj1(List1.ListIndex + 16) + "," + Combo1.Text
If polugo = 0 And Ocjene1.Caption = "" Then broj1(List1.ListIndex + 1) = broj1(List1.ListIndex + 1) + Combo1.Text
If polugo = 1 And Ocjene2.Caption = "" Then broj1(List1.ListIndex + 16) = broj1(List1.ListIndex + 16) + Combo1.Text
If polugo = 0 Then Ocjene1.Caption = broj1(List1.ListIndex + 1)
If polugo = 1 Then Ocjene2.Caption = broj1(List1.ListIndex + 16)
Call Calculate(X)
Call ApplyLabels(X)
Call Calculate(X)
Call ApplyLabels(X)
End If
If List1.ListIndex = -1 Then MsgBox ("You first need to select a subject to view,edit or delete its contects!")
End Sub

Private Sub LblAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdAdd.Picture = Skinz.Bdn.Picture
LblAdd.ForeColor = &H8000000E
End Sub

Private Sub LblAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdAdd.Picture = Skinz.Bup.Picture
LblAdd.ForeColor = &H80000012
End Sub

Private Sub LblDelete_Click()
If Predmet1.Caption = "XX" And polugo = 0 Then MsgBox ("You cant delete nothing!")
If Predmet2.Caption = "XX" And polugo = 1 Then MsgBox ("You cant delete nothing!")
If Len(Ocjene1.Caption) > 1 And polugo = 0 Then
    DeleteGN1.Caption = broj1(List1.ListIndex + 1)
    DeleteGN1.Caption = Left(DeleteGN1.Caption, Len(DeleteGN1) - 2)
    broj1(List1.ListIndex + 1) = DeleteGN1.Caption
    Ocjene1.Caption = broj1(List1.ListIndex + 1)
    Call Calculate(X)
    Call ApplyLabels(X)
    Call Calculate(X)
    Call ApplyLabels(X)
ElseIf Len(Ocjene1.Caption) = 1 And polugo = 0 Then
    DeleteGN1.Caption = broj1(List1.ListIndex + 1)
    DeleteGN1.Caption = Left(DeleteGN1.Caption, Len(DeleteGN1) - 1)
    broj1(List1.ListIndex + 1) = DeleteGN1.Caption
    Ocjene1.Caption = broj1(List1.ListIndex + 1)
    Call Calculate(X)
    Call ApplyLabels(X)
    Call Calculate(X)
    Call ApplyLabels(X)
End If
If Len(Ocjene2.Caption) > 1 And polugo = 1 Then
    DeleteGN2.Caption = broj1(List1.ListIndex + 16)
    DeleteGN2.Caption = Left(DeleteGN2.Caption, Len(DeleteGN2) - 2)
    broj1(List1.ListIndex + 16) = DeleteGN2.Caption
    Ocjene2.Caption = broj1(List1.ListIndex + 16)
    Call Calculate(X)
    Call ApplyLabels(X)
    Call Calculate(X)
    Call ApplyLabels(X)
ElseIf Len(Ocjene2.Caption) = 1 And polugo = 1 Then
    DeleteGN2.Caption = broj1(List1.ListIndex + 16)
    DeleteGN2.Caption = Left(DeleteGN2.Caption, Len(DeleteGN2) - 1)
    broj1(List1.ListIndex + 16) = DeleteGN2.Caption
    Ocjene2.Caption = broj1(List1.ListIndex + 16)
    Call Calculate(X)
    Call ApplyLabels(X)
    Call Calculate(X)
    Call ApplyLabels(X)
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

Private Sub LblPolugo_Click()
If polugo = 0 Then
    polugo = 1
    LblPolugo.Caption = " 1.Polugodište"
    Main.Comm1.Caption = "2"
ElseIf polugo = 1 Then
    polugo = 0
    LblPolugo.Caption = " 2.Polugodište"
    Main.Comm1.Caption = "1"
End If
End Sub

Private Sub LblPolugo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdPolugo.Picture = Skinz.Bdn.Picture
LblPolugo.ForeColor = &H8000000E
End Sub

Private Sub LblPolugo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then CmdPolugo.Picture = Skinz.Bup.Picture
LblPolugo.ForeColor = &H80000012
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
Ocjene1.Caption = broj1(List1.ListIndex + 1)
Ocjene2.Caption = broj1(List1.ListIndex + 16)
Call Calculate(X)
Call ApplyLabels(X)
End Sub

Function Calculate(X)
Calculator1.Caption = Len("," + Ocjene1.Caption) / 2
Calculator2.Caption = Len("," + Ocjene2.Caption) / 2
Tester1.Caption = (Val(Left(Ocjene1.Caption, 1)) + Val(Mid(Ocjene1.Caption, 3, 1)) + Val(Mid(Ocjene1.Caption, 5, 1)) + Val(Mid(Ocjene1.Caption, 7, 1)) + Val(Mid(Ocjene1.Caption, 9, 1)) + Val(Mid(Ocjene1.Caption, 11, 1)) + Val(Mid(Ocjene1.Caption, 13, 1)) + Val(Mid(Ocjene1.Caption, 15, 1)) + Val(Mid(Ocjene1.Caption, 17, 1)) + Val(Mid(Ocjene1.Caption, 19, 1)) + Val(Mid(Ocjene1.Caption, 21, 1)) + Val(Mid(Ocjene1.Caption, 23, 1)) + Val(Mid(Ocjene1.Caption, 25, 1)) + Val(Mid(Ocjene1.Caption, 27, 1)) + Val(Mid(Ocjene1.Caption, 29, 1)) + Val(Mid(Ocjene1.Caption, 31, 1)) + Val(Mid(Ocjene1.Caption, 33, 1)) + Val(Mid(Ocjene1.Caption, 35, 1)) + Val(Mid(Ocjene1.Caption, 37, 1)) + Val(Mid(Ocjene1.Caption, 39, 1))) / Calculator1.Caption
Tester2.Caption = (Val(Left(Ocjene2.Caption, 1)) + Val(Mid(Ocjene2.Caption, 3, 1)) + Val(Mid(Ocjene2.Caption, 5, 1)) + Val(Mid(Ocjene2.Caption, 7, 1)) + Val(Mid(Ocjene2.Caption, 9, 1)) + Val(Mid(Ocjene2.Caption, 11, 1)) + Val(Mid(Ocjene2.Caption, 13, 1)) + Val(Mid(Ocjene2.Caption, 15, 1)) + Val(Mid(Ocjene2.Caption, 17, 1)) + Val(Mid(Ocjene2.Caption, 19, 1)) + Val(Mid(Ocjene2.Caption, 21, 1)) + Val(Mid(Ocjene2.Caption, 23, 1)) + Val(Mid(Ocjene2.Caption, 25, 1)) + Val(Mid(Ocjene2.Caption, 27, 1)) + Val(Mid(Ocjene2.Caption, 29, 1)) + Val(Mid(Ocjene2.Caption, 31, 1)) + Val(Mid(Ocjene2.Caption, 33, 1)) + Val(Mid(Ocjene2.Caption, 35, 1)) + Val(Mid(Ocjene2.Caption, 37, 1)) + Val(Mid(Ocjene2.Caption, 39, 1))) / Calculator2.Caption
Predmet1.Caption = Replace(Predmet1.Caption, ",", ".")
Predmet2.Caption = Replace(Predmet2.Caption, ",", ".")
Finale1.Caption = Left((Val(Ocjene(0).Caption) + Val(Ocjene(1).Caption) + Val(Ocjene(2).Caption) + Val(Ocjene(3).Caption) + Val(Ocjene(4).Caption) + Val(Ocjene(5).Caption) + Val(Ocjene(6).Caption) + Val(Ocjene(7).Caption) + Val(Ocjene(8).Caption) + Val(Ocjene(9).Caption) + Val(Ocjene(10).Caption) + Val(Ocjene(11).Caption) + Val(Ocjene(12).Caption) + Val(Ocjene(13).Caption) + Val(Ocjene(14).Caption)) / 15, 3)
Finale2.Caption = Left((Val(Ocjene(15).Caption) + Val(Ocjene(16).Caption) + Val(Ocjene(17).Caption) + Val(Ocjene(18).Caption) + Val(Ocjene(19).Caption) + Val(Ocjene(20).Caption) + Val(Ocjene(21).Caption) + Val(Ocjene(22).Caption) + Val(Ocjene(23).Caption) + Val(Ocjene(24).Caption) + Val(Ocjene(25).Caption) + Val(Ocjene(26).Caption) + Val(Ocjene(27).Caption) + Val(Ocjene(28).Caption) + Val(Ocjene(29).Caption)) / 15, 3)
End Function

Function ApplyLabels(X)
tcap1 = Val(Replace(Tester1.Caption, ",", "."))
tcap2 = Val(Replace(Tester2.Caption, ",", "."))
If tcap1 >= 4.5 And tcap1 <= 5 Then Ocjene(List1.ListIndex).Caption = "5"
If tcap1 >= 3.5 And tcap1 <= 4.49 Then Ocjene(List1.ListIndex).Caption = "4"
If tcap1 >= 2.5 And tcap1 <= 3.49 Then Ocjene(List1.ListIndex).Caption = "3"
If tcap1 >= 1.5 And tcap1 <= 2.49 Then Ocjene(List1.ListIndex).Caption = "2"
If tcap1 >= 1 And tcap1 <= 1.49 Then Ocjene(List1.ListIndex).Caption = "1"
If tcap2 >= 4.5 And tcap2 <= 5 Then Ocjene(List1.ListIndex + 15).Caption = "5"
If tcap2 >= 3.5 And tcap2 <= 4.49 Then Ocjene(List1.ListIndex + 15).Caption = "4"
If tcap2 >= 2.5 And tcap2 <= 3.49 Then Ocjene(List1.ListIndex + 15).Caption = "3"
If tcap2 >= 1.5 And tcap2 <= 2.49 Then Ocjene(List1.ListIndex + 15).Caption = "2"
If tcap2 >= 1 And tcap2 <= 1.49 Then Ocjene(List1.ListIndex + 15).Caption = "1"
Predmet1.Caption = Left(Tester1.Caption, 3)
Predmet2.Caption = Left(Tester2.Caption, 3)
If Finale1.Caption = "5" Then Finale1.Caption = "5.0"
If Finale1.Caption = "4" Then Finale1.Caption = "4.0"
If Finale1.Caption = "3" Then Finale1.Caption = "3.0"
If Finale1.Caption = "2" Then Finale1.Caption = "2.0"
If Finale1.Caption = "1" Then Finale1.Caption = "1.0"
If Finale1.Caption = "0" Then Finale1.Caption = "XX"
If Finale2.Caption = "5" Then Finale2.Caption = "5.0"
If Finale2.Caption = "4" Then Finale2.Caption = "4.0"
If Finale2.Caption = "3" Then Finale2.Caption = "3.0"
If Finale2.Caption = "2" Then Finale2.Caption = "2.0"
If Finale2.Caption = "1" Then Finale2.Caption = "1.0"
If Finale2.Caption = "0" Then Finale2.Caption = "XX"
If Predmet1.Caption = "5" Then Predmet1.Caption = "5.0"
If Predmet1.Caption = "4" Then Predmet1.Caption = "4.0"
If Predmet1.Caption = "3" Then Predmet1.Caption = "3.0"
If Predmet1.Caption = "2" Then Predmet1.Caption = "2.0"
If Predmet1.Caption = "1" Then Predmet1.Caption = "1.0"
If Predmet1.Caption = "0" Then Predmet1.Caption = "XX"
If Predmet2.Caption = "5" Then Predmet2.Caption = "5.0"
If Predmet2.Caption = "4" Then Predmet2.Caption = "4.0"
If Predmet2.Caption = "3" Then Predmet2.Caption = "3.0"
If Predmet2.Caption = "2" Then Predmet2.Caption = "2.0"
If Predmet2.Caption = "1" Then Predmet2.Caption = "1.0"
If Predmet2.Caption = "0" Then Predmet2.Caption = "XX"
End Function
Private Sub MinSmall_Click()
School.WindowState = 1
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

