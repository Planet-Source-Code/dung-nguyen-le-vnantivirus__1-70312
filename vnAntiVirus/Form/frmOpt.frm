VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame frmAutoScan 
      Caption         =   "Auto detect and auto scan"
      Height          =   3135
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdBrowFolder 
         Caption         =   "..."
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "C:\Windows"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CheckBox chkDec 
         Caption         =   "Detect file add in system"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox chkSam 
         Caption         =   "Scan with sample"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkAutoIT 
         Caption         =   "Scan file write on AutoIt"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkScanI 
         Caption         =   "Scan icon"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkUSB 
         Caption         =   "Auto scan USB disk"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Value           =   1  'Checked
         Width           =   3015
      End
   End
   Begin VB.Frame frmSta 
      Caption         =   "Startup"
      Height          =   3135
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox chkSystemTray 
         Caption         =   "Show in Systemtray"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "Show program"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   2775
      End
      Begin VB.CheckBox chkStartup 
         Caption         =   "Work when startup"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   2775
      End
   End
   Begin VB.Frame frmLang 
      Caption         =   "Language :"
      Height          =   3135
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton optVie 
         Caption         =   "Vietnammese"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton optEng 
         Caption         =   "English"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ima 
      Left            =   1440
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpt.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpt.frx":11DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpt.frx":1E2E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1920
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   5741
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ima"
      ForeColor       =   12582912
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnAvant"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 1.0

'Author : Dung Le Nguyen
'Email : dungcoivb@gmail.com
'This is a software open source

Private Sub chkAutoIT_Click()
    cmdOk.Enabled = True
    chkUSB.Value = Checked
    If (chkAutoIT.Value = Unchecked) And (chkScanI.Value = Unchecked) And (chkSam.Value = Unchecked) Then chkUSB.Value = Unchecked
End Sub

Private Sub chkDec_Click()
    cmdOk.Enabled = True
    If chkDec.Value = Checked Then
        cmdBrowFolder.Enabled = True
        txtPath.Enabled = True
    Else
        cmdBrowFolder.Enabled = False
        txtPath.Enabled = False
    End If
End Sub
Private Sub chkSam_Click()
    cmdOk.Enabled = True
    chkUSB.Value = Checked
    If (chkAutoIT.Value = Unchecked) And (chkScanI.Value = Unchecked) And (chkSam.Value = Unchecked) Then chkUSB.Value = Unchecked
End Sub
Private Sub chkScanI_Click()
    cmdOk.Enabled = True
    chkUSB.Value = Checked
    If (chkAutoIT.Value = Unchecked) And (chkScanI.Value = Unchecked) And (chkSam.Value = Unchecked) Then chkUSB.Value = Unchecked
End Sub
Private Sub chkShow_Click()
    cmdOk.Enabled = True
    If (chkSystemTray.Value = Unchecked) And (chkShow.Value = Unchecked) Then chkSystemTray.Value = Checked
End Sub
Private Sub chkStartup_Click()
    cmdOk.Enabled = True
    If chkStartup.Value = Checked Then
        SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "vnAntivirus", PathApp & "\" & App.exename
    Else
        DelSetting HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "vnAntivirus"
    End If
End Sub

Private Sub chkSystemTray_Click()
    cmdOk.Enabled = True
    If (chkSystemTray.Value = Unchecked) And (chkShow.Value = Unchecked) Then chkShow.Value = Checked
End Sub

Private Sub chkUSB_Click()
    cmdOk.Enabled = True
    If chkUSB.Value = Unchecked Then
        chkScanI.Enabled = False
        chkAutoIT.Enabled = False
        chkSam.Enabled = False
        chkScanI.Value = Unchecked
        chkAutoIT.Value = Unchecked
        chkSam.Value = Unchecked
    Else
        chkScanI.Enabled = True
        chkAutoIT.Enabled = True
        chkSam.Enabled = True
        chkScanI.Value = Checked
        chkAutoIT.Value = Checked
        chkSam.Value = Checked
    End If
End Sub

Private Sub cmdBrowFolder_Click()
Dim sOutPut
    sOutPut = ""
    sOutPut = GetFolder(Me.hwnd, "Scan Path : ", txtPath.Text)
    If sOutPut <> "" Then
        txtPath.Text = sOutPut
    Else
        If Len(txtPath.Text) <> 0 Then
            Else
            ThongBao "vnAntiVirus", GetStr("MesSe")
    End If
    End If
End Sub

Private Sub cmdCancel_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub cmdOk_Click()

        cmdOk.Enabled = False
        ControlSet chkUSB, ichkUSB
        ControlSet chkScanI, ichkScanI
        ControlSet chkAutoIT, ichkAutoIT
        ControlSet chkSam, ichkSam
        ControlSet chkDec, ichkDec
        
        ControlSet chkShow, ichkShow
        ControlSet chkSystemTray, ichkSystemTray
        
        ioptVie = optVie.Value
        
    WriteINI PathApp & "\Data.ini", "Option", "chkUSB", QuyDoi1(ichkUSB)
    WriteINI PathApp & "\Data.ini", "Option", "chkScanI", QuyDoi1(ichkScanI)
    WriteINI PathApp & "\Data.ini", "Option", "chkAutoIT", QuyDoi1(ichkAutoIT)
    WriteINI PathApp & "\Data.ini", "Option", "chkSam", QuyDoi1(ichkSam)
    WriteINI PathApp & "\Data.ini", "Option", "chkDec", QuyDoi1(ichkDec)
    WriteINI PathApp & "\Data.ini", "Option", "PathDec", txtPath.Text

    WriteINI PathApp & "\Data.ini", "Option", "chkShow", QuyDoi1(ichkShow)
    WriteINI PathApp & "\Data.ini", "Option", "chkSystemTray", QuyDoi1(ichkSystemTray)
    WriteINI PathApp & "\Data.ini", "Option", "optVie", QuyDoi1(ioptVie)
        ResetLV
        Language Me
        Language frmMnu
    'frmMain.Show
    'frmPro.Show
End Sub

Private Sub Form_Load()

    Language Me
    'MsgBox ReadINI("D:\MySoft\LovePN\Language\EngLish.lng", "frmOpt", "chkUSB")
        SetControl chkUSB, ichkUSB
        SetControl chkScanI, ichkScanI
        SetControl chkAutoIT, ichkAutoIT
        SetControl chkSam, ichkSam
        SetControl chkDec, ichkDec
        ResetLV
        If ichkDec = True Then
            txtPath.Enabled = True
            cmdBrowFolder.Enabled = True
        Else
            txtPath.Enabled = False
            cmdBrowFolder.Enabled = False
        End If
        txtPath.Text = PathDec
        
        SetControl chkShow, ichkShow
        SetControl chkSystemTray, ichkSystemTray
        
        If GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "vnAntivirus") = PathApp & "\" & App.exename Then
            chkStartup.Value = Checked
        Else
            chkStartup.Value = Unchecked
        End If
        
        optVie.Value = ioptVie
        optEng.Value = Not (ioptVie)
        cmdOk.Enabled = False
End Sub

'Hi hi, doan code nay hoi "Ngo"
Private Sub SetControl(chkOK As CheckBox, iValue As Boolean)
If iValue = True Then
    chkOK.Value = Checked
Else
    chkOK.Value = Unchecked
End If
End Sub
Private Sub ControlSet(chkOK As CheckBox, iValue As Boolean)
If chkOK.Value = Checked Then
    iValue = True
Else
    iValue = False
End If
End Sub
Private Sub LamMoi(frmOk As Frame)

Dim frm As Control
Dim tmp As String
Dim tmp1 As String
tmp1 = frmOk.Name
For Each i In Me.Controls
    tmp = i.Name
    If Left(tmp, 3) = "frm" Then
        i.Visible = False
        If tmp1 = tmp Then i.Visible = True
    End If
    
Next
End Sub

Private Sub LV_Click()
If LV.SelectedItem.Index = 1 Then
    LamMoi frmSta
ElseIf LV.SelectedItem.Index = 2 Then
    LamMoi frmLang
ElseIf LV.SelectedItem.Index = 3 Then
    LamMoi frmAutoScan
End If
End Sub
Private Sub LV_KeyUp(KeyCode As Integer, Shift As Integer)
If LV.SelectedItem.Index = 1 Then
    LamMoi frmSta
ElseIf LV.SelectedItem.Index = 2 Then
    LamMoi frmLang
ElseIf LV.SelectedItem.Index = 3 Then
    LamMoi frmAutoScan
End If
End Sub
Private Sub optEng_Click()
    cmdOk.Enabled = True
End Sub
Private Sub optVie_Click()
    cmdOk.Enabled = True
End Sub
Private Sub txtPath_Change()
    cmdOk.Enabled = True
End Sub
Private Sub ResetLV()
    LV.ListItems.Clear
    LV.ListItems.Add , , GetStrOther("Sta"), 1
    LV.ListItems.Add , , GetStrOther("Lan"), 2
    LV.ListItems.Add , , GetStrOther("Aut"), 3
    LV.Refresh
End Sub
