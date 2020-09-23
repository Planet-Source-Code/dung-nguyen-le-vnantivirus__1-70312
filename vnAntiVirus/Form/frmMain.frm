VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vnAntivirus"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2760
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   2760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMain 
      Caption         =   "Main :"
      ForeColor       =   &H8000000D&
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   2535
      End
      Begin VB.CommandButton cmdOpt 
         Caption         =   "Data of virus of user"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   2535
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Option"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CommandButton cmdAu 
         Caption         =   "Author"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3480
         Width           =   2535
      End
      Begin VB.CommandButton cmdPro 
         Caption         =   "Processes"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan with sample"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton cmdScanSys 
         Caption         =   "Scan system"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton cmdSta 
         Caption         =   "Startup"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000D&
         X1              =   120
         X2              =   2640
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000D&
         X1              =   120
         X2              =   2640
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000D&
         X1              =   120
         X2              =   2640
         Y1              =   1200
         Y2              =   1200
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 1.0

'Author : Dung Le Nguyen
'Email : dungcoivb@gmail.com
'This is a software open source

Private Sub cmdAu_Click()
    frmAbout.Show
    Unload Me
End Sub
Private Sub cmdHide_Click()
    frmMnu.Show
    frmMnu.Hide
    Unload Me
End Sub
Private Sub cmdOpt_Click()
    frmOpt.Show
    Unload Me
End Sub
Private Sub cmdPro_Click()
    frmPro.Show
    Unload Me
End Sub
Private Sub cmdScan_Click()
    frmPht.Show
    Unload Me
End Sub
Private Sub cmdScanSys_Click()
Dim sOutPut
    sOutPut = ""
    sOutPut = GetFolder(Me.hwnd, "Scan Path : ", WindowsDir)
    If sOutPut <> "" Then
        PathWScan = sOutPut
        frmScan.Show
    Else
        ThongBao "vnAntiVirus", GetStr("MesSe")
    End If
End Sub
Private Sub cmdSta_Click()
    frmSta.Show
    Unload Me
End Sub
Private Sub cmdUp_Click()
    frmDat.Show
    Unload Me
End Sub
Private Sub Form_Load()

    App.TaskVisible = False
    PathApp = App.Path
    If Right(PathApp, 1) = "\" Then PathApp = Left(Path, Len(PathApp) - 1)
If FileExists(PathApp & "\Data.ini") = False Then
    ThongBao "vnAntiVirus", GetStr("MesNS")
    Me.Show
Else
    GetOpt
    If ichkShow = True Then
        Me.Show
    Else
        Me.Hide
    End If
    If ichkSystemTray = True Then Load frmMnu
End If

If (FileExists(PathApp & "\Language\VietNam.lng") = False) Or (FileExists(PathApp & "\Language\EngLish.lng") = False) Then
    bLang = False
    ThongBao "vnAntiVirus", GetStr("MesFL")
Else
    bLang = True
    Language Me
    Language frmMnu
End If

'Tinh nang Monitor hien nay lam viec chua on dinh nen tam thoi chua co mat
'If LoadMon = False Then
'    If FileExists(PathApp & "\Mon\Mon.exe") = True Then
'        Shell PathApp & "\Mon\Mon.exe " & PathDec
'        LoadMon = True
'    Else
'        ThongBao "vnAntiVirus", GetStr("MesNF") & " " & PathApp & "\Mon\Mon.exe"
'    End If
'End If
    SeeSta = False
End Sub
