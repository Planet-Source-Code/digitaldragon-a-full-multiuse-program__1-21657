VERSION 5.00
Begin VB.Form Skinz 
   BorderStyle     =   0  'None
   Caption         =   "Skinz"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   2985
   ScaleWidth      =   1800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Back 
      Height          =   495
      Left            =   1200
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Bdn 
      Height          =   495
      Left            =   600
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image Bup 
      Height          =   495
      Left            =   0
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image BorderR 
      Height          =   495
      Left            =   1200
      Top             =   1800
      Width           =   495
   End
   Begin VB.Image BorderL 
      Height          =   495
      Left            =   1200
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image BorderDR 
      Height          =   495
      Left            =   1200
      Top             =   600
      Width           =   495
   End
   Begin VB.Image DragImgR 
      Height          =   495
      Left            =   1200
      Top             =   0
      Width           =   495
   End
   Begin VB.Image ExitDN 
      Height          =   495
      Left            =   600
      Top             =   1800
      Width           =   495
   End
   Begin VB.Image MinDN 
      Height          =   495
      Left            =   600
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image BorderDL 
      Height          =   495
      Left            =   600
      Top             =   600
      Width           =   495
   End
   Begin VB.Image DragImgL 
      Height          =   495
      Left            =   600
      Top             =   0
      Width           =   495
   End
   Begin VB.Image ExitUP 
      Height          =   495
      Left            =   0
      Top             =   1800
      Width           =   495
   End
   Begin VB.Image MinUP 
      Height          =   495
      Left            =   0
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image BorderD 
      Height          =   495
      Left            =   0
      Top             =   600
      Width           =   495
   End
   Begin VB.Image DragImg 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Skinz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NAME   WIDTH   HEIGHT
'
'Button 1425    330
'
'Drager 7500    330
'
'Small  225     225
'
'vert           75
'
'horz   60
