VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add data worm"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBro 
      Caption         =   "..."
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "&Add"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame frmSet 
      Caption         =   "Setting :"
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   4695
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   275
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   275
      End
      Begin VB.OptionButton optFile 
         Caption         =   "MD5 code"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   3735
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "Scan with icon"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblIcon 
         BackStyle       =   0  'Transparent
         Caption         =   "Detect with method scan icon"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "Dectect with method scan compare MD5 code"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   1080
         Width           =   3375
      End
   End
   Begin VB.FileListBox File 
      Height          =   510
      Left            =   1440
      Pattern         =   "*.ico"
      TabIndex        =   3
      Top             =   3360
      Width           =   615
   End
   Begin VB.PictureBox picCom 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   960
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   300
   End
   Begin MSComctlLib.ImageList ima 
      Left            =   120
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1080
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label cmdInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "With function, user can add code detect virus :)"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Virus name :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Caption         =   "Path sample :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmAdd"
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
    cd.Filter = "Portable files (*.pif;*.exe)|*.exe;*.pif|All Files (*.*)|*.*"
    Dim Path As String
    cd.ShowOpen
    If cd.Filename <> "" Then txtPath.Text = cd.Filename: GetIcon txtPath.Text, Pic: Pic.Visible = True
End Sub
Private Sub cmdCancel_Click()
    frmDat.Show
    Unload Me
End Sub

Private Sub cmdUp_Click()
Dim kq As String
If txtName.Text <> "" Then
    If optIcon.Value = True Then
        If FileExists(PathApp & "\Dat\Icon\" & txtName.Text & ".ico") = True Then
            ThongBao "vnAntiVirus", GetStr("MesTDL")
        Else
            kq = KiemTraIcon
            If kq = "" Then
                SavePicture Pic.Image, PathApp & "\Dat\Icon\" & txtName.Text & ".ico"
                frmDat.GetInfo
                ThongBao "vnAntiVirus", GetStr("MesComUI") & txtName.Text
            Else
                ThongBao "vnAntiVirus", GetStr("MesIAs") & kq
            End If
        End If
    ElseIf optFile.Value = True Then

        Dim tt As Boolean
        Dim tt1 As Boolean
        
        tt = False
        tt1 = False
        
        Dim tn As String
        Dim MD5 As String
        MD5 = GetMD5(txtPath.Text)
            If MD5 <> "0" Then
                Dim PathTmp As String
                PathTmp = PathApp & "\Dat\Sign\" & GetExt(txtPath.Text) & "\" & Left(MD5, 2) & ".vnd"
                Dim InputData As String
                Open PathTmp For Input As #1
                    Do While Not EOF(1)
                        Line Input #1, InputData
                        If MD5 = Split(InputData, "|", , vbBinaryCompare)(0) Then tt = True
                        If txtName.Text = Split(InputData, "|", , vbBinaryCompare)(1) Then tt1 = True
                    Loop
                Close #1
                If tt = True Then ThongBao "vnAntiVirus", "Virus na2y d9a4 d9u7o75c ca65p nha65t": GoTo KetThuc1
                If tt1 = True Then ThongBao "vnAntiVirus", "Te6n virus bi5 tru2ng": GoTo KetThuc1
                
                AddToFile MD5 & "|" & txtName.Text, PathTmp
                                ThongBao "vnAntiVirus", GetStr("MesComUD") & " : " & txtName.Text
            Else
                ThongBao "vnAntiVirus", "Du74 lie65u nha65p va2o co1 lo64i"
KetThuc1:
            End If

    End If
'End If
Else
    ThongBao "vnAntiVirus", GetStr("MesNN")
End If

End Sub
Private Sub Form_Load()
    Language Me
End Sub
Private Sub txtPath_Change()
If FileExists(txtPath.Text) = True Then GetIcon txtPath.Text, Pic: Pic.Visible = True
End Sub
Private Function KiemTraIcon() As String
    KiemTraIcon = ""
    'Xu ly thao tac kiem tra Icon xem co ton tai trong data chua
    ima.ListImages.Clear
    File.Path = PathApp & "\Dat\Icon"
    File.Refresh
    Dim i As Integer
    Dim re As Byte
    Dim kq As String
    kq = ""
    For i = 0 To File.ListCount - 1
        picCom.Cls
        picCom.Picture = LoadPicture(PathApp & "\Dat\Icon\" & File.List(i))
        PaP picCom, Pic, Pic.Width, Pic.Height, 15, re
        If re = 100 Then KiemTraIcon = Left(File.List(i), Len(File.List(i)) - 4)
    Next
End Function
