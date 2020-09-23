VERSION 5.00
Begin VB.Form frmScan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scan"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "frmScan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox chkCWI 
      Caption         =   "Scan with index"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CheckBox chkIndex 
      Caption         =   "Make index"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1800
      Width           =   2295
   End
   Begin vnAntivirus.XP_ProgressBar pro 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   6956042
      Orientation     =   1
      Scrolling       =   3
   End
   Begin VB.ListBox lstStr 
      Height          =   1035
      Left            =   1800
      TabIndex        =   5
      Top             =   5280
      Width           =   855
   End
   Begin VB.ListBox lstSDec 
      Height          =   1035
      Left            =   2640
      TabIndex        =   4
      Top             =   5280
      Width           =   855
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
      Height          =   615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   5280
   End
   Begin VB.ListBox lstName 
      Height          =   1035
      Left            =   720
      TabIndex        =   1
      Top             =   5280
      Width           =   855
   End
   Begin VB.ListBox lstDat 
      Height          =   1035
      Left            =   -120
      TabIndex        =   0
      Top             =   5280
      Width           =   855
   End
   Begin vnAntivirus.ucFirefoxWait ffw 
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Label lblPH 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   840
      TabIndex        =   19
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblDQ 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblTG 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblPhH 
      BackStyle       =   0  'Transparent
      Caption         =   "Detected :"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblDaQuet 
      BackStyle       =   0  'Transparent
      Caption         =   "Scaned :"
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time :"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblPathText 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label lblNP 
      BackStyle       =   0  'Transparent
      Caption         =   "Scaning..."
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblPathScan 
      BackStyle       =   0  'Transparent
      Caption         =   "Scan folder :"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 1.0

'Author : Dung Le Nguyen
'Email : dungcoivb@gmail.com
'This is a software open source

Option Explicit
Private WithEvents SP As cScanPath
Attribute SP.VB_VarHelpID = -1
Public dq As Long
Public tg As Long
Dim FullPath As String
Dim Scaning As Boolean
Dim ScaningIn As Boolean

Dim iIndex As Boolean
Dim iCWI As Boolean
Dim strTMP As String
Dim DemFile As Long
Dim i As Long
Dim ScaIndex As Boolean

Private Sub ThietLapForm()
    'pro.Color = RGB(226, 233, 246)
    'GetData App.Path & "\Dat\Sign.vnd", lstDat, lstName, True
    'GetData App.Path & "\Dat\SW\YourSign.vnd", lstDat, lstName, False
    'GetData App.Path & "\Dat\String.vnd", lstStr, lstSDec, True

'Note:    'vnd= VnAntivirus data ;-)
    ' dung co dich la "Viet Nam Dong" nhe, Dung Coi ko tham lam tien ...lam dau
    
End Sub

Private Sub chkCWI_Click()
    If chkCWI.Value = Checked Then chkIndex.Value = Unchecked
End Sub
Private Sub chkIndex_Click()
    If chkIndex.Value = Checked Then chkCWI.Value = Unchecked
End Sub
Private Sub cmdCancel_Click()
    frmMain.Show
    Unload Me
End Sub
Private Sub cmdScan_Click()
    'On Error Resume Next
If (Scaning = False) And (ScaningIn = True) Then GoTo CWI

If (Scaning = False) And (ScaningIn = False) Then
    cmdCancel.Enabled = False
    'frmDetect.LV.Visible = False
    ScaIndex = False
    Scaning = True
    dq = 0
    ph = 0
    tg = 0
    DemFile = 0
    'MsgBox strTmp & "\Index.vnd"
    Timer.Enabled = True
    If chkIndex.Value = Checked Then iIndex = True
    If chkCWI.Value = Checked Then
        ScaIndex = True
        Scaning = False
        GoTo CWI
    End If
    ffw.PlayWait
    lblPH.Caption = "0 file"
    cmdScan.Caption = GetStr("MesScanT")
    cmdCancel.Enabled = False
    Scaning = True
    'dq = 0
    'ph = 0
    'tg = 0
    Timer.Enabled = True
    'Tien hanh cac thao tac quet process
    SAll = True
    frmPro.Show
    DoEvents
    frmPro.ScanPro
    DoEvents
    
    frmSta.Show
    DoEvents
    frmSta.ScanSta
    DoEvents
    'Tien hanh quet cac file
    Set SP = New cScanPath
        
        If FileExists(PathApp & "\indexTmp.vnd") = True Then XoaFile PathApp & "\indexTmp.vnd"
        
        With SP
            .Archive = True
            .Compressed = True
            .Hidden = True
            .Normal = True
            .ReadOnly = True
            .System = True
            
            .Filter = "*.exe;*.pif;*.com;*.vbs;*.bat;*.asp;*.bin;*.chm;*.php;*.dll;*.eml;*.hta;*.htm*.class;*.mht;*.js;*.ocx"
            '"*.exe;*.pif;*.com;*.vbs;*.bat;*.asp;*.bin;*.chm;*.cpl;*.dll;*.drv;*.eml;*.hta;*.htm*.drv;*.mht;*.mp3;*.ocx;*.sys;*.url"
            .StartScan PathWScan, True, True
        End With
        'Debug.Print "Okie 2"
        'MsgBox "Xong liet ke"
        'Okie, sau khi da tao bang chi muc file xong
        'Exit Sub
    If Right(PathWScan, 1) = "\" Then PathWScan = Left(PathWScan, Len(PathWScan) - 1)
    If iIndex = True Then
        If FileExists(PathWScan & "\Index.vnd") = True Then XoaFile PathWScan & "\Index.vnd"
        FileCopy PathApp & "\indexTmp.vnd", PathWScan & "\Index.vnd"
        WriteINI PathWScan & "\Index.vnd", "Info", "File", CStr(DemFile)
    End If
    
    ScanIndex PathApp & "\indexTmp.vnd"

        Timer.Enabled = False
        cmdCancel.Enabled = True
        Scaning = False
        ffw.StopWait
        cmdScan.Caption = GetStr("MesScanF")
        ThongBao "vnAntiVirus", GetStr("MesComScan")
        lblDQ.Caption = dq
ElseIf (Scaning = True) And (ScaningIn = False) Then
    SP.StopScan
    ScaIndex = False
    cmdScan.Caption = GetStr("MesScanF")
    cmdCancel.Enabled = True
    Scaning = False
    ffw.StopWait
    ThongBao "vnAntiVirus", GetStr("MesStoStop")
End If

    Scaning = False
    Exit Sub
CWI:
If ScaningIn = False Then
    ScaningIn = True
    dq = 0
    ph = 0
    
    SAll = True
    frmPro.Show
    DoEvents
    frmPro.ScanPro
    DoEvents
    
    frmSta.Show
    DoEvents
    frmSta.ScanSta
    DoEvents
    
    cmdScan.Caption = GetStr("MesScanT")
    DemFile = Val(ReadINI(strTMP & "\Index.vnd", "Info", "File"))
    ScanIndex strTMP & "\Index.vnd"
    ThongBao "vnAntiVirus", GetStr("MesComScan")
Else
    ScaningIn = False
    cmdScan.Caption = GetStr("MesScanF")
    ThongBao "vnAntiVirus", GetStr("MesComScan")
End If
End Sub

Private Sub Form_Load()
    Language Me
    Scaning = False
    ScaningIn = False
    lblPathText.Caption = PathWScan
    If Right(PathWScan, 1) = "\" Then
        strTMP = Left(PathWScan, Len(PathWScan) - 1)
    Else
        strTMP = PathWScan
    End If
    If FileExists(strTMP & "\index.vnd") = True Then
        chkCWI.Visible = True
    Else
        chkCWI.Visible = False
    End If
    iIndex = False
    iCWI = False
    ThietLapForm
    ScaIndex = True
End Sub

Private Sub SP_FileMatch(Filename As String, Path As String)
'Dung luong tap tin duoc quet se nho hon 6000000 byte (Gan 6Mb)
If DungLuong(Path & Filename) > 6000000 Then GoTo KetThuc
    Dim KetQua As String
    FullPath = Path & Filename
    KetQua = GetMD5(FullPath)
        If Len(KetQua) < 8 Then
            Select Case Len(KetQua)
            'Dung coi nghi, viec xet truong hop se nhanh hon viec su dung For
                Case 7
                    KetQua = "0" & KetQua
                Case 6
                    KetQua = "00" & KetQua
                Case 5
                    KetQua = "000" & KetQua
                Case 4
                    KetQua = "0000" & KetQua
                   Case 3
                    KetQua = "00000" & KetQua
                Case 2
                    KetQua = "000000" & KetQua
                Case 4
                    KetQua = "0000000" & KetQua
            End Select
        End If
    AddToFile KetQua & "|" & Path & Filename, PathApp & "\indexTMP.vnd"
    DemFile = DemFile + 1
KetThuc:
End Sub
Private Sub Timer_Timer()
    'tg = tg + 1
    'lblTG.Caption = tg
    lblDQ.Caption = dq & " file"
    txtPath.Text = FullPath
End Sub
Public Sub ScanIndex(PathFileIndex As String)
    On Error Resume Next
Dim DatTmp As String
Dim MD5Tmp As String
Dim strPath As String

Open PathFileIndex For Input As #2
    Do While Not EOF(2) And (ScaIndex = True)
        Line Input #2, DatTmp
        If ScaningIn = False Then GoTo TheEnd
        If dq >= DemFile Then GoTo TheEnd
        'Dong lenh tren nham trach loi khi quet file index.vnd trong thu muc
        dq = dq + 1
        pro.Value = Int(dq / DemFile * pro.Max)
        MD5Tmp = Split(DatTmp, "|", , vbBinaryCompare)(0)
        strPath = Split(DatTmp, "|", , vbBinaryCompare)(1)
    'Thong qua Test, Dung Coi xac dinh duoc rang, neu thuc hien Check Icon thi toc do quet se giam 1.5 lan
    'ScanFile Path & Filename, lstDat, lstName, lstStr, lstSDec, True, True, True, False, frmMnu.Ima, frmMnu.Pic, frmMnu.pic1

        If ScanMD5Main(MD5Tmp, strPath) = True Then GoTo KetThucByCRC
        DoEvents
        
        'Tien hanh check virus qua chuoi String
        Dim BoDem As String
            Open strPath For Binary As #1
                BoDem = Space(LOF(1))
                Get #1, , BoDem
            Close #1
            
        Dim InputData As String
        
        Dim strCodeS As String
        Dim strDecS As String
        
        Open App.Path & "\Dat\String.vnd" For Input As #1
        Do While Not EOF(1)
            Line Input #1, InputData
            
            strCodeS = Split(InputData, "|", , vbBinaryCompare)(0)
            If InStr(1, BoDem, strCodeS, vbBinaryCompare) <> 0 Then
            'Yeah, da nhan ra virus roi nhe
            'Nhan dang theo ky thuat nhan dang chuoi String
                strDecS = Split(InputData, "|", , vbBinaryCompare)(1)
                Detect GetStr("DecFile"), strDecS, strPath
                ph = ph + 1
                lblPH.Caption = ph
                GoTo KetThuc
            End If
        Loop
            Close #1
            
            'Nhan dang virus thong qua chuoi string (Khong the xac dinh virus qua MD5)
        For i = 0 To frmMnu.lstSVir.ListCount - 1
            If InStr(1, BoDem, frmMnu.lstSVir.List(i), vbBinaryCompare) <> 0 Then
                Detect GetStr("DecVir"), frmMnu.lstVirNa.List(i), strPath, frmMnu.lstVirDat.List(i)
                ph = ph + 1
                lblPH.Caption = ph
                GoTo KetThuc
            End If
       Next
       
       BoDem = vbNullString

KetThuc:
    BoDem = vbNullString
KetThucByCRC:
    Loop
TheEnd:
Close #2

ScaIndex = True
frmDetect.GetIDProcess

End Sub

