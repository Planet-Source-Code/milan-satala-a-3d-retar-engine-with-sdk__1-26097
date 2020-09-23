VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditor 
   Caption         =   "Retard Editor - Mental Soft"
   ClientHeight    =   6510
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   434
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTextury 
      Height          =   1575
      Index           =   2
      Left            =   7800
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   16
      Top             =   3120
      Width           =   1575
   End
   Begin VB.PictureBox picTextury 
      Height          =   1575
      Index           =   1
      Left            =   7800
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   15
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox picTextury 
      Height          =   1575
      Index           =   0
      Left            =   7800
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.Frame fraTextury 
      Height          =   1695
      Left            =   7800
      TabIndex        =   10
      Top             =   4680
      Width           =   1590
      Begin VB.CommandButton cmdDel 
         Caption         =   "Delete Texture"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdCTex 
         Caption         =   "Change Texture"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdSDown 
         Caption         =   "Scroll Down"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdSUp 
         Caption         =   "Scroll Up"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CMD 
      Left            =   4440
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmTools 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1620
      Begin VB.PictureBox picProcess 
         Height          =   6255
         Left            =   0
         ScaleHeight     =   6195
         ScaleWidth      =   1515
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdKoniec 
         Caption         =   "Koniec"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton cmdPridaj 
         Caption         =   "Add Buffer"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblPos 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox pic2D 
      Height          =   2895
      Index           =   2
      Left            =   4920
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   4
      Top             =   3480
      Width           =   2655
   End
   Begin VB.PictureBox pic3D 
      Height          =   2895
      Left            =   4920
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.PictureBox pic2D 
      Height          =   2895
      Index           =   1
      Left            =   1800
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   2
      Top             =   3480
      Width           =   2655
   End
   Begin VB.PictureBox pic2D 
      Height          =   2895
      Index           =   0
      Left            =   1800
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu mnuSubory 
      Caption         =   "File"
      Begin VB.Menu mnuNovy 
         Caption         =   "New"
      End
      Begin VB.Menu mnuNahrat 
         Caption         =   "Load"
      End
      Begin VB.Menu mnuUlozit 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuComp 
         Caption         =   "Compile"
      End
      Begin VB.Menu nic 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKoniec 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TexturyIndex As Integer
Dim pTextura() As StdPicture


Dim Turning As Boolean
Dim ZacX As Integer
Dim ZacY As Integer
Dim stx, sty
Dim WhatDo As String

Private Sub cmdCTex_Click()
  CMD.Filter = "Windows bitmap (*.bmp)|*.bmp|All Files (*.*)|*.*"
  CMD.ShowOpen
  
  If CMD.FileName <> "" Then
   If EdTextury(VybrataTextura) <> "" Then Kill App.Path & "\Data\" & EdTextury(VybrataTextura)
   Set pTextura(VybrataTextura) = LoadPicture(CMD.FileName)
   EdTextury(VybrataTextura) = Dir(CMD.FileName)
   Set Textures(VybrataTextura) = D3DX.CreateTextureFromFile(D3DDevice, CMD.FileName)
   If Not CMD.FileName = App.Path & "\Data\" & Dir(CMD.FileName) Then
   FileCopy CMD.FileName, App.Path & "\Data\" & Dir(CMD.FileName)
   End If
   NakresliTextury
   RefreshAll
   If VybrataTextura = EditorSet.PocetTextur Then
    EditorSet.PocetTextur = EditorSet.PocetTextur + 1
    ReDim Preserve pTextura(EditorSet.PocetTextur)
    ReDim Preserve EdTextury(EditorSet.PocetTextur)
    ReDim Preserve Textures(EditorSet.PocetTextur)
   End If
  End If
End Sub

Private Sub cmdDel_Click()
  Kill App.Path & "\Data\" & EdTextury(VybrataTextura)
  EdTextury(VybrataTextura) = ""
  Set Textures(VybrataTextura) = Nothing
  Set pTextura(VybrataTextura) = Nothing
  NakresliTextury
  RefreshAll
End Sub

Private Sub cmdKoniec_Click()
  Unload Me
End Sub

Private Sub cmdPridaj_Click()
  Dim a As Integer
  WhatDo = ""
  
  For a = 1 To PocetStena
   If Stena(a).Zivy = False Then
    PridajCube a
    Exit Sub
   End If
  Next a
  
  PocetStena = PocetStena + 1
  PridajCube PocetStena
End Sub

Private Sub PridajCube(Cislo As Integer)
  With Stena(Cislo)
   .Pos(1).Y = 5
   .Pos(3).X = 5
   .Pos(2).X = 5
   .Pos(2).Y = 5
   .Zivy = True
   .Textura = VybrataTextura
   VypocitajPriem Cislo
  End With
  RefreshAll
End Sub

Private Sub cmdSDown_Click()
  If TexturyIndex + 2 < EditorSet.PocetTextur Then
   TexturyIndex = TexturyIndex + 1
   NakresliTextury
  End If
End Sub

Private Sub cmdSUp_Click()
  If TexturyIndex > 0 Then
   TexturyIndex = TexturyIndex - 1
   NakresliTextury
  End If
End Sub

Private Sub Form_Load()
  Dim a As Integer
  
  
  Open App.Path & "\Editor.cfg" For Binary Access Read As #1
   Get #1, , EditorSet
   ReDim EdTextury(EditorSet.PocetTextur)
   ReDim pTextura(EditorSet.PocetTextur)
   Get #1, , EdTextury
  Close
  
  If EditorSet.PocetTextur < 2 Then
   EditorSet.PocetTextur = 2
   ReDim Preserve EdTextury(EditorSet.PocetTextur)
   ReDim pTextura(EditorSet.PocetTextur)
  End If
  
  MapSet.TextureCount = EditorSet.PocetTextur
  ReDim Textures(MapSet.TextureCount)
      
  EdMapSet.AmbientLight = vbWhite
  MapVelkost(0) = 1
  MapVelkost(1) = 1
  MapVelkost(2) = 1
  Pos3D.Z = -1
  
  Show
  
  InitEngine pic3D.Hwnd, 1
  For a = 0 To EditorSet.PocetTextur
   If EdTextury(a) <> "" Then
    Set pTextura(a) = LoadPicture(App.Path & "\Data\" & EdTextury(a))
    Set Textures(a) = D3DX.CreateTextureFromFile(D3DDevice, App.Path & "\Data\" & EdTextury(a))
   End If
  Next a
  
  D3DDevice.EndScene
End Sub

Sub NakresliTextury()
  Dim a As Byte, b As Integer
  Dim Index As Integer
  Index = VybrataTextura - TexturyIndex
  
  For a = 0 To 2
   b = TexturyIndex + a
   If EdTextury(b) <> "" Then
    Me.picTextury(a).Picture = pTextura(b)
   Else
    Me.picTextury(a).Line (0, 0)-(1000, 1000), 0, BF
   End If
  Next a
  
  If Index >= 0 And Index < 3 Then
  Me.picTextury(Index).Circle (50, 50), (20), vbWhite
  End If
End Sub

Private Sub Form_Paint()
  Show
  NakresliTextury
  RefreshAll
End Sub

Private Sub Form_Resize()
  Dim VelX As Integer, VelY As Integer, a As Byte
  VelX = (Me.ScaleWidth - 250) / 2
  VelY = (Me.ScaleHeight - 28) / 2
  
  Me.pic2D(0).Height = VelY
  Me.pic2D(0).Width = VelX
  
  Me.pic2D(1).Top = VelY + 18
  Me.pic2D(1).Height = VelY
  Me.pic2D(1).Width = VelX
  
  Me.pic3D.Left = 130 + VelX
  Me.pic3D.Width = VelX
  Me.pic3D.Height = VelY
  
  Me.pic2D(2).Left = 130 + VelX
  Me.pic2D(2).Width = VelX
  Me.pic2D(2).Top = VelY + 18
  Me.pic2D(2).Height = VelY
  
  Me.fraTextury.Left = Me.ScaleWidth - 110
  
  For a = 0 To 2
   Me.picTextury(a).Left = Me.ScaleWidth - 110
   Me.picTextury(a).DrawWidth = 10
  Next a
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Open App.Path & "\Editor.cfg" For Binary Access Write As #1
   Put #1, , EditorSet
   Put #1, , EdTextury
  Close
     
  ExitEngine
End Sub

Private Sub mnuComp_Click()
  
  CMD.Filter = "Retard Compiled Map (*.rcm)|*.rcm|All Files (*.*)|*.*"
  CMD.FileName = SuborComp
  CMD.ShowSave
  SuborComp = CMD.FileName
  On Error Resume Next
  Kill SuborComp
  
  Me.picProcess.Visible = True
  Me.picProcess.Cls

  If CompileMap = True Then MsgBox "File Saved !" Else MsgBox "Compiling Not Finished !"
  
  Me.picProcess.Visible = False
End Sub

Private Sub PosunBuffer(Zac As Integer)
  Dim a As Integer
  PocetStena = PocetStena - 1
  For a = Zac To PocetStena
   Stena(a) = Stena(a + 1)
  Next a
End Sub

Private Sub mnuKoniec_Click()
  End
End Sub

Private Sub mnuNahrat_Click()
  Dim TMPSet As reMapSettings
  CMD.Filter = "Retard Editor Map (*.rem)|*.rem|All Files (*.*)|*.*"
  CMD.FileName = SuborSave
  CMD.ShowSave
  SuborSave = CMD.FileName

  Open SuborSave For Binary Access Read Write As #1
   Get #1, , EdMapSet
   Get #1, , Stena
   Get #1, , EdEntity
  Close
  
  PocetStena = EdMapSet.BufferCount
  PocetEntity = EdMapSet.EntityCount
  RefreshAll
End Sub

Private Sub mnuUlozit_Click()
  Dim TMPSet As reMapSettings
  CMD.Filter = "Retard Editor Map (*.rem)|*.rem|All Files (*.*)|*.*"
  CMD.FileName = SuborSave
  CMD.ShowSave
  SuborSave = CMD.FileName
  
  TMPSet = EdMapSet
  With TMPSet
   .BufferCount = PocetStena
   .EntityCount = PocetEntity
  End With
  
   Open SuborSave For Binary Access Read Write As #1
    Put #1, , TMPSet
    Put #1, , Stena
    Put #1, , EdEntity
   Close
End Sub

Private Sub pic2D_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim PX(3) As Integer, PY(3) As Integer
  Dim a  As Integer
    
  Select Case Button
  Case 1 And VybrateCislo > 0
  Select Case WhatDo
   Case ""
    With Stena(VybrateCislo)
     For a = 0 To 3
      If Index = 0 Or Index = 1 Then
       PX(a) = (MapX(Index) + MapZoom(Index) + .Pos(a).X) * MapVelkost(Index)
      Else
       PX(a) = (MapX(Index) + MapZoom(Index) + .Pos(a).Y) * MapVelkost(Index)
      End If
      If Index = 0 Then
       PY(a) = (MapY(Index) + MapZoom(Index) + .Pos(a).Y) * MapVelkost(Index)
      Else
       PY(a) = (MapY(Index) + MapZoom(Index) + .Pos(a).Z) * MapVelkost(Index)
      End If
     Next a
    
     For a = 0 To 3
      If GetDist(PX(a), PY(a), X, Y) < 4 Then
       DrzimCo = a
       Drzim = True
       Exit Sub
      End If
     Next a
    

     Drzim = True
     DrzimCo = 4
     For a = 0 To 3
      TMPStena(a).X = .Pos(a).X - MysX
      TMPStena(a).Y = .Pos(a).Y - MysY
      TMPStena(a).Z = .Pos(a).Z - MysZ
     Next a
    
    End With
   Case "EMove"
    With EdEntity(VybrateCislo)
     If Index = 0 Or Index = 1 Then .Pos.X = MysX Else .Pos.Y = MysY
     If Index = 0 Then .Pos.Y = MysY Else .Pos.Z = MysZ
    
     RefreshAll
    End With
  End Select
  Case 2
   ZacX = X
   ZacY = Y
 End Select
End Sub

Private Sub pic2D_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim PX As Integer, PY As Integer
  Dim a As Integer
  
  If Index = 0 Or Index = 1 Then
   MysX = -MapX(Index) + Int(X / MapVelkost(Index)) - MapZoom(Index)
  Else
   MysY = -MapX(Index) + Int(X / MapVelkost(Index)) - MapZoom(Index)
  End If
  If Index = 0 Then
   MysY = -MapY(Index) + Int(Y / MapVelkost(Index)) - MapZoom(Index)
  Else
   MysZ = -MapY(Index) + Int(Y / MapVelkost(Index)) - MapZoom(Index)
  End If
  
  lblPos = "X: " & MysX & " Y: " & MysY & " Z: " & MysZ
  
  Select Case Button
   Case 1
    Select Case WhatDo
     Case ""
      If Drzim = True Then
       With Stena(VybrateCislo)
        If DrzimCo = 4 Then
         For a = 0 To 3
          If Index = 0 Or Index = 1 Then
           .Pos(a).X = MysX + TMPStena(a).X
          End If
          If Index = 0 Or Index = 2 Then
           .Pos(a).Y = MysY + TMPStena(a).Y
          End If
          If Index = 1 Or Index = 2 Then
           .Pos(a).Z = MysZ + TMPStena(a).Z
          End If
         Next a
        Else
         If Index = 0 Or Index = 1 Then
          .Pos(DrzimCo).X = MysX
         End If
         If Index = 0 Or Index = 2 Then
          .Pos(DrzimCo).Y = MysY
         End If
         If Index = 1 Or Index = 2 Then
          .Pos(DrzimCo).Z = MysZ
         End If
        End If
       End With
       VypocitajPriem VybrateCislo
       RefreshAll
      End If
     Case "EMove"
      With EdEntity(VybrateCislo)
       If Index = 0 Or Index = 1 Then .Pos.X = -MysX Else .Pos.Y = MysY
       If Index = 0 Then .Pos.Y = MysY Else .Pos.Z = MysZ
       RefreshAll
      End With
    End Select
   Case 2
    If X <> ZacX Or Y <> ZacY Then
     Scroll = True
     MapX(Index) = MapX(Index) + (-X + ZacX) / 4
     MapY(Index) = MapY(Index) + (-Y + ZacY) / 4
     
     ZacX = X
     ZacY = Y
     
     Refresh2D (Index)
    End If
  End Select
End Sub

Private Sub pic2D_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim a As Integer
  Dim Vz As Integer, NajVz As Integer, Naj As Integer
  NajVz = 10000
    
  Select Case Button
   Case 1
    Drzim = False
   Case 2
    If Scroll = False Then
     For a = 1 To PocetStena
      With Stena(a)
       If Index = 0 Then
        Vz = GetDist(.PriemX, .PriemY, MysX, MysY)
       ElseIf Index = 0 Then
        Vz = GetDist(.PriemX, .PriemZ, MysX, MysZ)
       Else
        Vz = GetDist(.PriemY, .PriemZ, MysY, MysZ)
       End If
       If Vz < NajVz And Not (a = VybrateCislo And CoVybrate = "") Then
        NajVz = Vz
        Naj = a
        CoVybrate = ""
        WhatDo = ""
       End If
      End With
     Next a
  
     For a = 1 To PocetEntity
      With EdEntity(a)
       If Index = 0 Then
        Vz = GetDist(.Pos.X, .Pos.Y, MysX, MysY)
       ElseIf Index = 0 Then
        Vz = GetDist(.Pos.X, .Pos.Z, MysX, MysZ)
       Else
        Vz = GetDist(.Pos.Y, .Pos.Z, MysY, MysZ)
       End If
       If Vz < NajVz And Not (a = VybrateCislo And CoVybrate = "E") Then
        NajVz = Vz
        Naj = a
        CoVybrate = "E"
        WhatDo = "EMove"
       End If
      End With
     Next a
  
     VybrateCislo = Naj
     VybrataTextura = Stena(Naj).Textura
     RefreshAll
     
    Else
     Scroll = False
    End If
   End Select
End Sub

Private Sub pic3D_KeyDown(KeyCode As Integer, Shift As Integer)
  With Pos3D
   If KeyCode = vbKeyUp Then
     MoveEntityX Pos3D, .LookDegX, 1
   End If
   If KeyCode = vbKeyDown Then
    MoveEntityX Pos3D, .LookDegX + 180, 1
   End If
   
   If KeyCode = vbKeyLeft Then MoveEntityX Pos3D, .LookDegX - 90, 1
   If KeyCode = vbKeyRight Then MoveEntityX Pos3D, .LookDegX + 90, 1
   
   If KeyCode = vbKeyControl Then Pos3D.Z = Pos3D.Z + 1
   If KeyCode = vbKeyShift Then Pos3D.Z = Pos3D.Z - 1
   
   lblPos = "X: " & Int(.X) & " Y: " & Int(.Y) & " Z: " & Int(.Z)
  End With
  
  RefreshBuffersForEditor
  
  Refresh3D
End Sub

Private Sub pic3D_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ZacX = X
  ZacY = Y
  Turning = True
End Sub

Private Sub pic3D_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  With Pos3D
   If Turning = True Then
    TurnEntityX Pos3D, (X - ZacX) / 2
    TurnEntityY Pos3D, (Y - ZacY) / 2
    ZacX = X
    ZacY = Y
    
    EntityPosLook Pos3D
    RefreshBuffersForEditor
    Refresh3D
   End If
  End With
End Sub

Private Sub pic3D_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Turning = False
  stx = Pos3D.LookX
  sty = Pos3D.LookY
End Sub

Private Sub picTextury_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  VybrataTextura = TexturyIndex + Index
  If EdTextury(VybrataTextura) <> "" And VybrateCislo And CoVybrate = "" Then
   Stena(VybrateCislo).Textura = VybrataTextura
   RefreshAll
  End If
  
  NakresliTextury
End Sub
