VERSION 5.00
Begin VB.Form frmMas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MsgBox"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdScan 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1800
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   1200
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   1200
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000002&
      FillColor       =   &H80000008&
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label lblMes 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMas.frx":0000
      DragMode        =   1  'Automatic
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3705
   End
   Begin VB.Label lblCaption 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblPath 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "frmMas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'vnAntiVirus 1.0

'Author : Dung Le Nguyen
'Email : dungcoivb@gmail.com
'This is a software open source

'Code this form from PSC

Private Const SPI_GETWORKAREA As Long = 48&
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
    (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Type OSVersionInfo
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion  As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVersionInfo) As Long

Private m_iChangeSpeed    As Long         '/* The window's display speed
Private m_iCounter        As Long         '/* Display time in milliseconds
Private m_iScrnBottom     As Long         '/* Height of the screen - taskbar (if it is on the bottom)
Private m_bOnTop          As Boolean      '/* Form Z-Order Flag
Private m_iWindowCount    As Long         '/* Screen stop position multiplier (displaying more then 1 at a time)
Private m_bManualClose    As Boolean      '/* Manual close Flag
Private m_bCodeClose      As Boolean      '/* Prevent user close option
Private m_bFade           As Boolean      '/* Fade or move Flag
Private m_iOSver          As Byte         '/* OS 1=Win98/ME; 2=Win2000/XP
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Dim tg As Byte
'Dim formCap As String
'Dim mas As String
Dim co As Byte
'Dim com As String

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdScan_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Language Me
    Me.Caption = lblCaption.Caption
  Dim rc         As RECT
  Dim scrnRight  As Long
  Dim OSV        As OSVersionInfo
        OSV.OSVSize = Len(OSV)
        
    If GetVersionEx(OSV) = 1 Then
        If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then m_iOSver = 1 '/* Win 98/ME
        If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then m_iOSver = 2  '/* Win 2000/XP
    End If
    
    '/* Get Screen and TaskBar size
    Call SystemParametersInfo(SPI_GETWORKAREA, 0&, rc, 0&)
    
    '/* Screen Height - Taskbar Height (if is is located at the bottom of the screen)
    m_iScrnBottom = rc.Bottom * Screen.TwipsPerPixelY
    
    '/* Is the taskbar is located on the right side of the screen? (scrnRight < Screen.width)
    scrnRight = (rc.Right * Screen.TwipsPerPixelX)
    
    '/* Locate Form to bottom right and set default size
    Me.Move scrnRight - Me.Width, m_iScrnBottom, lblMes.Left + lblMes.Width + 100, cmdScan.Top + 700
    
    'Me.Move scrnRight - Me.Width, m_iScrnBottom - Me.Height, txtMas.Left + txtMas.Width + 100, cmdScan.Top + 800
    
    Timer2.Enabled = True
    tg = 10
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub Timer_Timer()
If tg - co > 0 Then
    co = co + 1
    Me.Caption = lblCaption.Caption & " (" & tg - co & ")"
Else
    Timer1.Enabled = True
End If
End Sub
Private Sub Timer1_Timer()
    If Me.Top > m_iScrnBottom Then
        Unload Me
    Else
        Me.Move Me.Left, Me.Top + 25, lblMes.Left + lblMes.Width + 100, cmdScan.Top + 700
    End If
End Sub

Private Sub Timer2_Timer()
    If Me.Top < m_iScrnBottom - Me.Height Then
        Timer.Enabled = True
        Timer2.Enabled = False
    Else
        Me.Caption = lblCaption.Caption
        Me.Move Me.Left, Me.Top - 25, lblMes.Left + lblMes.Width + 100, cmdScan.Top + 700
    End If
End Sub
