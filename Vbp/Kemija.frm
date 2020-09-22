VERSION 5.00
Begin VB.Form ToDo 
   BorderStyle     =   0  'None
   Caption         =   "Kemija"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Kemija.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   2775
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1575
      Begin VB.Label Label5 
         Caption         =   "Ugljik"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Ljuske 
         Caption         =   "(2,4,1)"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label MaLbl 
         Caption         =   "16,78"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label ArLbl 
         Caption         =   "14.01"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label ProLbl 
         Caption         =   "6"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.Label InfoLbl1 
         BackColor       =   &H80000001&
         Caption         =   "Pro"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Element 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cu"
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label InfoLbl5 
         BackColor       =   &H80000001&
         Caption         =   "Naziv"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label InfoLbl4 
         BackColor       =   &H80000001&
         Caption         =   "Ljuske"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label InfoLbl3 
         BackColor       =   &H80000001&
         Caption         =   "Ma"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label InfoLbl2 
         BackColor       =   &H80000001&
         Caption         =   "Ar"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label DragLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Kemija"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   45
      Width           =   3135
   End
   Begin VB.Image CmdExit 
      Height          =   330
      Left            =   420
      Top             =   2520
      Width           =   1425
   End
   Begin VB.Image CmdKroz 
      Height          =   330
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   945
   End
   Begin VB.Image CmdUdio 
      Height          =   330
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   945
   End
   Begin VB.Image CmdMr 
      Height          =   330
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   945
   End
   Begin VB.Image CmdMf 
      Height          =   330
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   945
   End
   Begin VB.Image CmdAr 
      Height          =   330
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   960
      Width           =   945
   End
   Begin VB.Image CmdMa 
      Height          =   330
      Left            =   120
      Stretch         =   -1  'True
      Top             =   960
      Width           =   945
   End
   Begin VB.Image CmdDoIt 
      Height          =   330
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   480
      Width           =   945
   End
   Begin VB.Image BorderD 
      Height          =   75
      Left            =   0
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4020
   End
   Begin VB.Image ExitSmall 
      Height          =   225
      Left            =   3760
      Top             =   60
      Width           =   225
   End
   Begin VB.Image MinSmall 
      Height          =   225
      Left            =   3520
      Top             =   60
      Width           =   225
   End
   Begin VB.Image DragImg 
      Height          =   330
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4020
   End
   Begin VB.Image BorderR 
      Height          =   3735
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   240
      Width           =   60
   End
   Begin VB.Image BorderL 
      Height          =   3735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   60
   End
End
Attribute VB_Name = "ToDo"
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
Kemija.Picture = Skinz.Back.Picture
DragImg.Picture = Skinz.DragImg.Picture
MinSmall.Picture = Skinz.MinUP.Picture
ExitSmall.Picture = Skinz.ExitUP.Picture
BorderL.Picture = Skinz.BorderL.Picture
BorderR.Picture = Skinz.BorderR.Picture
BorderD.Picture = Skinz.BorderD.Picture
CmdDoIt.Picture = Skinz.Bup.Picture
CmdAr.Picture = Skinz.Bup.Picture
CmdMa.Picture = Skinz.Bup.Picture
CmdMf.Picture = Skinz.Bup.Picture
CmdMr.Picture = Skinz.Bup.Picture
CmdUdio.Picture = Skinz.Bup.Picture
CmdKroz.Picture = Skinz.Bup.Picture
CmdExit.Picture = Skinz.Bup.Picture
End Sub

Private Sub MinSmall_Click()
Kemija.WindowState = 1
End Sub

Private Sub MinSmall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MinSmall.Picture = Skinz.MinDN.Picture
End Sub

Private Sub MinSmall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then MinSmall.Picture = Skinz.MinUP.Picture
End Sub

