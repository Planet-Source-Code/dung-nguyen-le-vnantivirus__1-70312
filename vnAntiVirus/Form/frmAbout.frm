VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Author"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblVer 
      Caption         =   "Update : 3/19/2008"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "vnAntivirus 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblSend 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label lblOS 
      BackStyle       =   0  'Transparent
      Caption         =   "vnAntivirus is a open source software"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label lblSup 
      BackStyle       =   0  'Transparent
      Caption         =   "Support (ID Yahoo) : dungcoivb"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblSam 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You can send sample virus to : dungcoivb@gmail.com (With Pass is : injected )"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "Email : dungcoivb@gmail.com"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Author : Dung Le Nguyen"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image ima 
      Height          =   720
      Left            =   120
      Picture         =   "frmAbout.frx":058A
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 1.0

'Author : Dung Le Nguyen
'Email : dungcoivb@gmail.com
'This is a software open source

Private Sub cmdBack_Click()
    frmMain.Show
    Unload Me
End Sub
Private Sub Form_Load()
    Language Me
End Sub
Private Sub lblSam_Click()
    Shell "EXPLORER.EXE " & "http://www.vietvirus.info"
End Sub
