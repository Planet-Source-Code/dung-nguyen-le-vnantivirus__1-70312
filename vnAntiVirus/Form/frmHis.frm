VERSION 5.00
Begin VB.Form frmHis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtHis 
      Appearance      =   0  'Flat
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label lblHis 
      BackStyle       =   0  'Transparent
      Caption         =   "History scan :"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmHis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
