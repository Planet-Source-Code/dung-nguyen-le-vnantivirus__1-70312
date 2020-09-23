VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add key startup"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBro 
      Caption         =   "..."
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Add"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2640
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   3855
   End
   Begin VB.ComboBox cmbKey 
      Appearance      =   0  'Flat
      Height          =   330
      ItemData        =   "frmAddKey.frx":058A
      Left            =   1080
      List            =   "frmAddKey.frx":0594
      TabIndex        =   1
      Text            =   "HKEY_LOCAL_MACHINE"
      Top             =   120
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label lblPath 
      Caption         =   "Path :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Key name :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblType 
      Caption         =   "Type :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmAddKey"
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
    cd.Filter = "Protable files (*.pif;*.exe)|*.exe;*.pif|All Files (*.*)|*.*"
    Dim Path As String
    cd.ShowOpen
    If cd.Filename <> "" Then txtPath.Text = cd.Filename
End Sub

Private Sub cmdCancel_Click()
    frmSta.Show
    Unload Me
End Sub

Private Sub cmdOk_Click()

Dim GiaTri As String
If (txtName.Text <> "") And (txtPath.Text <> "") Then
    If cmbKey.Text = "HKEY_CURRENT_USER" Then
    
        GiaTri = GetString(HKEY_CURRENT_USER, Pathkey, txtName.Text)
        If GiaTri = "" Then SaveString HKEY_CURRENT_USER, Pathkey, txtName.Text, txtPath.Text: ThongBao "vnAntiVirus", GetStr("MesComAdd")
    ElseIf cmbKey.Text = "HKEY_LOCAL_MACHINE" Then
    
        GiaTri = GetString(HKEY_LOCAL_MACHINE, Pathkey, txtName.Text)
        If GiaTri = "" Then SaveString HKEY_LOCAL_MACHINE, Pathkey, txtName.Text, txtPath.Text: ThongBao "vnAntiVirus", GetStr("MesComAdd")
    End If
Else
    ThongBao "vnAntiVirus", GetStr("MesNF")
End If
End Sub

Private Sub Form_Load()
        Language Me
End Sub
