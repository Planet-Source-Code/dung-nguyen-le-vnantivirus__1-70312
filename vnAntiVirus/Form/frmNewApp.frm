VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNewApp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New app"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "frmNewApp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   5220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Back"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdBro 
      Caption         =   "..."
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "C:\Sampleworm\RealWorm.exe"
      Top             =   120
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1440
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   5040
      X2              =   2880
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblOpen 
      BackStyle       =   0  'Transparent
      Caption         =   "Open :"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmNewApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 1.0

'Author : Dung Le Nguyen
'Email : dungcoivb@gmail.com
'This is a software open source

Private Sub cmdBro_Click()
    cd.DialogTitle = "Choose a file ..."
    cd.Filter = "Portable files (*.exe)|*.exe"
    Dim Path As String
    cd.ShowOpen
    If cd.Filename <> "" Then txtPath.Text = cd.Filename
End Sub

Private Sub cmdCancel_Click()
    frmPro.Show
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If FileExists(txtPath.Text) = True Then Shell txtPath.Text
End Sub

Private Sub Form_Load()
        Language Me
End Sub
