VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMapSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map settings"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3225
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picAmbient 
      Height          =   255
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   120
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label5 
      Caption         =   "Map Light:"
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Top             =   1320
      Width           =   750
   End
   Begin VB.Line Line2 
      X1              =   600
      X2              =   3120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label4 
      Caption         =   "Lights"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   420
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   3120
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Labels"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "Author:"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   510
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   465
   End
End
Attribute VB_Name = "frmMapSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Zrus As Byte, Color As Long

Private Sub cmdCancel_Click()
  Zrus = 0
  Unload Me
End Sub

Private Sub cmdOk_Click()
  EdMapSet.Author = txtAuthor
  EdMapSet.MapName = txtName
  EdMapSet.AmbientLight = Color
  
  Zrus = 0
  Unload Me
End Sub

Private Sub Form_Load()
  Show
  DoEvents
  Color = EdMapSet.AmbientLight
  Me.picAmbient.Line (0, 0)-(2000, 1000), Color, BF
  Me.txtAuthor = EdMapSet.Author
  Me.txtName = EdMapSet.MapName
  Zrus = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Cancel = Zrus
  frmEditor.Enabled = True
  frmEditor.Show
  frmEditor.SetFocus
End Sub

Private Sub picAmbient_Click()
  CMD.ShowColor
  
  Color = CMD.Color
  Me.picAmbient.Line (0, 0)-(2000, 1000), Color, BF
End Sub
