VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frmPro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRe 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan processes"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   2040
      Top             =   4080
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSComctlLib.ImageList ima 
      Left            =   1320
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3945
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6959
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ima"
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tªn tiÕn tr×nh"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "§­êng dÉn tiÕn tr×nh"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ChØ sè"
         Object.Width           =   1676
      EndProperty
   End
End
Attribute VB_Name = "frmPro"
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

Private Sub cmdRe_Click()
    GetProcess LV, ima, Pic
End Sub
Public Sub ScanPro()
Dim i As Integer
With frmMnu
    For i = 1 To LV.ListItems.Count
        If FileExists(LV.ListItems(i).SubItems(1)) = True Then
            ScanFile LV.ListItems(i).SubItems(1), True, True, True, .ima, .Pic, .pic1
        End If
    Next
End With
    If tb = True Then frmDetect.GetIDProcess
    ThongBao "vnAntiVirus", GetStr("MesComScanMe")
    If SAll = True Then Unload Me: SPro = True
End Sub

Private Sub cmdScan_Click()
    ScanPro
End Sub

Private Sub Form_Load()
    Language Me
    GetProcess LV, ima, Pic
End Sub
Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu frmMnu.mnuc0
End Sub
Private Sub Timer_Timer()
  Dim i As Integer
  Dim theloop As Long
  Dim proc As PROCESSENTRY32
  Dim snap As Long
    i = 0
   snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0)
   proc.dwSize = Len(proc)
   theloop = ProcessFirst(snap, proc)
   While theloop <> 0

      theloop = ProcessNext(snap, proc)
      If ProcessPathByPID(proc.th32ProcessID) <> "SYSTEM" Then
              i = i + 1
              If i > LV.ListItems.Count Then GoTo KetThuc
            If LV.ListItems(i).SubItems(1) <> ProcessPathByPID(proc.th32ProcessID) Then GoTo KetThuc
      End If
   Wend
   CloseHandle snap
Exit Sub
KetThuc:
    GetProcess LV, ima, Pic
End Sub
