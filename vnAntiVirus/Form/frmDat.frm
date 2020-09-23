VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View you data"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   3540
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ima 
      Left            =   1680
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "And"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.Frame frmIcon 
      Caption         =   "Detect with icon :"
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3615
      Begin MSComctlLib.ListView LVI 
         Height          =   3225
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5689
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "Ima"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnAvant"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tªn"
            Object.Width           =   5821
         EndProperty
      End
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   5040
      Width           =   300
   End
   Begin VB.FileListBox File 
      Height          =   300
      Left            =   0
      Pattern         =   "*.ico"
      TabIndex        =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This function to manager icons being use when scan with method detect icon"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   3375
   End
End
Attribute VB_Name = "frmDat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 1.0

'Author : Dung Le Nguyen
'Email : dungcoivb@gmail.com
'This is a software open source

Private Sub cmdAdd_Click()
    frmAdd.Show
End Sub
Private Sub cmdBack_Click()
    frmMain.Show
    Unload Me
End Sub
Private Sub Form_Load()
    Language Me
    GetInfo
End Sub
Public Sub GetInfo()
'Load icon
File.Path = PathApp & "\Dat\Icon"
File.Refresh
    ThietLap LVI, ima, Pic
If File.ListCount <> 0 Then

    Dim i As Integer
    For i = 0 To File.ListCount - 1
        LVI.ListItems.Add , , Left(File.List(i), Len(File.List(i)) - 4)
    Next
    
        For i = 0 To File.ListCount - 1
            Pic.Cls
            Pic.Picture = LoadPicture(PathApp & "\Dat\Icon" & "\" & File.List(i))
            ima.ListImages.Add i + 1, , Pic.Image
        Next
        
    With LVI
      .SmallIcons = ima
      For Each lsv In .ListItems
        lsv.SmallIcon = lsv.Index
      Next
    End With
End If

End Sub
Private Sub LVI_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu frmMnu.mnud0
End Sub
