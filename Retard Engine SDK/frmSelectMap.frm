VERSION 5.00
Begin VB.Form frmSelectMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select map"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.DirListBox dirMap 
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.FileListBox fleMap 
      Height          =   1845
      Left            =   2640
      Pattern         =   "*.rcm"
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Select map ..."
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2760
      Width           =   4455
   End
End
Attribute VB_Name = "frmSelectMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuit_Click()
  End
End Sub

Private Sub cmdStart_Click()
  If fleMap.FileName = "" Then Exit Sub
  MapFile = fleMap.Path & "\" & fleMap.FileName
  Me.lblStatus = "Starting game ..."
  DoEvents
  Load frmGame
End Sub

Private Sub dirMap_Change()
  fleMap.Path = dirMap.Path
End Sub

Private Sub fleMap_DblClick()
  cmdStart_Click
End Sub

Private Sub Form_Load()
  dirMap.Path = App.Path
End Sub
