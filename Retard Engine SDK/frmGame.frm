VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retard Engine (Software developer kit)"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
  EngineSet.With = 800
  EngineSet.Height = 600
  InitEngine frmGame.Hwnd, 1 'Initialize game engine
  LoadMeshFromFile "Data\tiger.x" 'Load bullets mesh
  
  LoadMapFromFile MapFile 'Load map
  RefreshBuffers 'Refresh buffers
  
  PlayerEnt = AddEntity 'Create entity for player (us)
  RetardEntity(PlayerEnt).Type = rePlayer 'Set entity type to player
  AddEntity '(bot which look like tiger)
  RetardEntity(2).Type = rePlayer 'Set entity type to player
  
  frmGame.Show 'Shows game form
  Unload frmSelectMap
  
  Do 'Main game tick
   Game 'All game is here
  Loop

End Sub
