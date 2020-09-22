VERSION 5.00
Begin VB.Form Raspored 
   BorderStyle     =   0  'None
   Caption         =   "Raspored"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
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
   ScaleHeight     =   5655
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   60
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   59
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ujutro"
      Height          =   3855
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   4335
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   39
         Left            =   3480
         TabIndex        =   55
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   38
         Left            =   3480
         TabIndex        =   54
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   37
         Left            =   3480
         TabIndex        =   53
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   36
         Left            =   3480
         TabIndex        =   52
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   35
         Left            =   3480
         TabIndex        =   51
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   34
         Left            =   3480
         TabIndex        =   50
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   33
         Left            =   3480
         TabIndex        =   49
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   32
         Left            =   3480
         TabIndex        =   48
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   31
         Left            =   2760
         TabIndex        =   47
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   30
         Left            =   2760
         TabIndex        =   46
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   29
         Left            =   2760
         TabIndex        =   45
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   28
         Left            =   2760
         TabIndex        =   44
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   27
         Left            =   2760
         TabIndex        =   43
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   26
         Left            =   2760
         TabIndex        =   42
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   25
         Left            =   2760
         TabIndex        =   41
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   24
         Left            =   2760
         TabIndex        =   40
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   23
         Left            =   2040
         TabIndex        =   39
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   22
         Left            =   2040
         TabIndex        =   38
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   21
         Left            =   2040
         TabIndex        =   37
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   20
         Left            =   2040
         TabIndex        =   36
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   19
         Left            =   2040
         TabIndex        =   35
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   18
         Left            =   2040
         TabIndex        =   34
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   17
         Left            =   2040
         TabIndex        =   33
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   16
         Left            =   2040
         TabIndex        =   32
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   15
         Left            =   1320
         TabIndex        =   31
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   14
         Left            =   1320
         TabIndex        =   30
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   13
         Left            =   1320
         TabIndex        =   29
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   12
         Left            =   1320
         TabIndex        =   28
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   11
         Left            =   1320
         TabIndex        =   27
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   10
         Left            =   1320
         TabIndex        =   26
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   9
         Left            =   1320
         TabIndex        =   25
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   8
         Left            =   1320
         TabIndex        =   24
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   23
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   22
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   5
         Left            =   600
         TabIndex        =   21
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   20
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   19
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   18
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Sat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Sat8 
         BackStyle       =   0  'Transparent
         Caption         =   " 8"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Sat7 
         BackStyle       =   0  'Transparent
         Caption         =   " 7"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Sat6 
         BackStyle       =   0  'Transparent
         Caption         =   " 6"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Sat5 
         BackStyle       =   0  'Transparent
         Caption         =   " 5"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Sat4 
         BackStyle       =   0  'Transparent
         Caption         =   " 4"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Sat3 
         BackStyle       =   0  'Transparent
         Caption         =   " 3"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Sat2 
         BackStyle       =   0  'Transparent
         Caption         =   " 2"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Sat1 
         BackStyle       =   0  'Transparent
         Caption         =   " 1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.Label PetLbl 
         BackStyle       =   0  'Transparent
         Caption         =   " Pet"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label CetLbl 
         BackStyle       =   0  'Transparent
         Caption         =   " Èet"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label SriLbl 
         BackStyle       =   0  'Transparent
         Caption         =   " Sri"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label UtoLbl 
         BackStyle       =   0  'Transparent
         Caption         =   " Uto"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label PonLbl 
         BackStyle       =   0  'Transparent
         Caption         =   " Pon"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.Label SatLbl 
         BackStyle       =   0  'Transparent
         Caption         =   " S"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000E&
         Height          =   400
         Left            =   120
         Top             =   240
         Width           =   4095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000E&
         Height          =   3495
         Left            =   120
         Top             =   240
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000E&
         FillColor       =   &H8000000E&
         Height          =   3495
         Left            =   120
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Label LblReturn 
      BackStyle       =   0  'Transparent
      Caption         =   "       Return"
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   5075
      Width           =   1335
   End
   Begin VB.Label LblEdit 
      BackStyle       =   0  'Transparent
      Caption         =   "         Edit"
      Height          =   255
      Left            =   360
      TabIndex        =   56
      Top             =   5075
      Width           =   1335
   End
   Begin VB.Image CmdEdit 
      Height          =   330
      Left            =   360
      Top             =   5040
      Width           =   1425
   End
   Begin VB.Image CmdReturn 
      Height          =   330
      Left            =   3240
      Top             =   5040
      Width           =   1425
   End
   Begin VB.Image ExitSmall 
      Height          =   225
      Left            =   4635
      Top             =   60
      Width           =   225
   End
   Begin VB.Image MinSmall 
      Height          =   225
      Left            =   4360
      Top             =   60
      Width           =   225
   End
   Begin VB.Label DragLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Raspored"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   61
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
      Height          =   5055
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   300
   End
   Begin VB.Image BorderR 
      Height          =   4455
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   600
      Width           =   150
   End
   Begin VB.Image BorderD 
      Height          =   420
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   5235
      Width           =   3320
   End
   Begin VB.Image DragImg 
      Height          =   600
      Left            =   495
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3720
   End
   Begin VB.Image BorderDR 
      Height          =   645
      Left            =   4380
      Top             =   5010
      Width           =   570
   End
   Begin VB.Image DragImgR 
      Height          =   600
      Left            =   4215
      Top             =   0
      Width           =   735
   End
   Begin VB.Image BorderDL 
      Height          =   375
      Left            =   0
      Top             =   5280
      Width           =   1080
   End
   Begin VB.Label LblTurnus 
      BackStyle       =   0  'Transparent
      Caption         =   "      Popodne"
      Height          =   255
      Left            =   360
      TabIndex        =   58
      Top             =   4595
      Width           =   1335
   End
   Begin VB.Image CmdTurnus 
      Height          =   330
      Left            =   360
      Top             =   4560
      Width           =   1425
   End
   Begin VB.Label LblClear 
      BackStyle       =   0  'Transparent
      Caption         =   "        Clear"
      Height          =   255
      Left            =   3240
      TabIndex        =   57
      Top             =   4595
      Width           =   1335
   End
   Begin VB.Image CmdClear 
      Height          =   330
      Left            =   3240
      Top             =   4560
      Width           =   1425
   End
End
Attribute VB_Name = "Raspored"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private n, polugo, MotherFucker As Integer

Private Sub CmdApply_Click()
If n = -1 Then MsgBox ("Ahhh,like i said first select one before u can edit it,ok?")
If n > -1 Then
Sat(n).Caption = Text1.Text
If polugo = 0 Then
Open ".\Text Files\Raspored1.txt" For Output As 1
For i = 0 To 39
Print #1, Sat(i)
Next
Close
ElseIf polugo = 1 Then
Open ".\Text Files\Raspored2.txt" For Output As 2
For i = 0 To 39
Print #2, Sat(i)
Next
Close
End If
End If
End Sub

Private Sub CmdClear_Click()
Call LblClear_Click
End Sub

Private Sub CmdClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblClear_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblClear_MouseUp(Button, Shift, X, Y)
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

Private Sub CmdReturn_Click()
Call LblReturn_Click
End Sub

Private Sub CmdReturn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblReturn_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdReturn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblReturn_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub CmdTurnus_Click()
Call LblTurnus_Click
End Sub

Private Sub CmdTurnus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblTurnus_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub CmdTurnus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call LblTurnus_MouseUp(Button, Shift, X, Y)
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
ExitSmall.Picture = Skinz.ExitUP.Picture
MinSmall.Picture = Skinz.MinUP.Picture
BorderL.Picture = Skinz.BorderL.Picture
BorderR.Picture = Skinz.BorderR.Picture
BorderD.Picture = Skinz.BorderD.Picture
CmdReturn.Picture = Skinz.Bup.Picture
CmdClear.Picture = Skinz.Bup.Picture
CmdTurnus.Picture = Skinz.Bup.Picture
CmdEdit.Picture = Skinz.Bup.Picture
Raspored.Picture = Skinz.Back.Picture
DragImgL.Picture = Skinz.DragImgL.Picture
DragImgR.Picture = Skinz.DragImgR.Picture
BorderDL.Picture = Skinz.BorderDL.Picture
BorderDR.Picture = Skinz.BorderDR.Picture
Open ".\Text Files\Raspored1.txt" For Input As 1
Do Until EOF(1)
    n = n + 1
    Line Input #1, nextline
    Sat(n - 1) = nextline
Loop
Close
n = 0
If Main.Comm2.Caption = "1" Then
    LblTurnus.Caption = "      Popodne"
    Call LblTurnus_Click
ElseIf Main.Comm2.Caption = "2" Then LblTurnus.Caption = "       Ujutro"
    Call LblTurnus_Click
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
n = 0
End Sub

Private Sub LblClear_Click()
If CmdApply.Visible = False Then MsgBox ("This button can only be used when editing!")
If CmdApply.Visible = True Then
For i = 0 To 39
Sat(i).BackStyle = 0
Sat(i).ForeColor = &H80000018
Next
n = -1
Text1.Text = ""
End If
End Sub

Private Sub LblClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdClear.Picture = Skinz.Bdn.Picture
LblClear.ForeColor = &H8000000E
End Sub

Private Sub LblClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdClear.Picture = Skinz.Bup.Picture
LblClear.ForeColor = &H80000012
End Sub

Private Sub LblEdit_Click()
If LblEdit.Caption = "         Edit" Then
    n = -1
    LblEdit.Caption = "        Close"
    CmdApply.Visible = True
    Text1.Visible = True
ElseIf LblEdit.Caption = "        Close" Then
    LblEdit.Caption = "         Edit"
    CmdApply.Visible = False
    Text1.Visible = False
    For i = 0 To 39
    Sat(i).BackStyle = 0
    Sat(i).ForeColor = &H80000018
    Next
    n = -1
    Text1.Text = ""
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

Private Sub LblTurnus_Click()
If LblTurnus.Caption = "      Popodne" Then
    Sat1.Caption = " 0"
    Sat2.Caption = " 0"
    Sat3.Caption = " 1"
    Sat4.Caption = " 2"
    Sat5.Caption = " 3"
    Sat6.Caption = " 4"
    Sat7.Caption = " 5"
    Sat8.Caption = " 6"
    LblTurnus.Caption = "       Ujutro"
    Frame1.Caption = "Popodne"
    If CmdApply.Visible = True Then Call LblEdit_Click
    n = 0
    Open ".\Text Files\Raspored2.txt" For Input As 1
    Do Until EOF(1)
        n = n + 1
        Line Input #1, nextline
        Sat(n - 1) = nextline
    Loop
    Close
    n = 0
    polugo = 1
    Main.Comm2.Caption = "1"
ElseIf LblTurnus.Caption = "       Ujutro" Then
    Sat1.Caption = " 1"
    Sat2.Caption = " 2"
    Sat3.Caption = " 3"
    Sat4.Caption = " 4"
    Sat5.Caption = " 5"
    Sat6.Caption = " 6"
    Sat7.Caption = " 7"
    Sat8.Caption = " 8"
        LblTurnus.Caption = "      Popodne"
    Frame1.Caption = "Ujutro"
    If CmdApply.Visible = True Then Call LblEdit_Click
    n = 0
    Open ".\Text Files\Raspored1.txt" For Input As 1
    Do Until EOF(1)
        n = n + 1
        Line Input #1, nextline
        Sat(n - 1) = nextline
    Loop
    Close
    n = 0
    polugo = 0
    Main.Comm2.Caption = "2"
End If
End Sub

Private Sub LblTurnus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdTurnus.Picture = Skinz.Bdn.Picture
LblTurnus.ForeColor = &H8000000E
End Sub

Private Sub LblTurnus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CmdTurnus.Picture = Skinz.Bup.Picture
LblTurnus.ForeColor = &H80000012
End Sub

Private Sub Sat_Click(Index As Integer)
If CmdApply.Visible = True Then
Text1.Text = Sat(Index).Caption
n = Sat(Index).Index
For i = 0 To 39
Sat(i).BackStyle = 0
Sat(i).ForeColor = &H80000018
Next
Sat(Index).BackStyle = 1
Sat(Index).BackColor = &H80000018
Sat(Index).ForeColor = &H80000001
End If
End Sub

Private Sub MinSmall_Click()
Raspored.WindowState = 1
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
