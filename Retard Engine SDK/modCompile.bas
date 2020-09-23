Attribute VB_Name = "modCompile"
Option Explicit


Dim BufNum As Integer
Dim TMPSet As reMapSettings
Dim SizeX As Integer, SizeY As Integer

Public Function CompileMap() As Boolean
  Dim a As Integer, b As Byte
  Dim PouzivaTex(0 To 1000) As Integer, CisloTex As Integer
  Dim StartTime As Long
  Dim TexturyNazov() As String
  Dim AC As D3DVECTOR, AB As D3DVECTOR, DC As D3DVECTOR, DB As D3DVECTOR
  Dim ABC As D3DVECTOR, DBC As D3DVECTOR, TMPW As D3DVECTOR
  Dim d As Single, t As Single
  Dim TMPPocetEntit As Integer

  
  StartTime = GetTickCount
  PridajText "Starting at " & Time
  PridajText ""
  PridajText ""
  TMPSet = EdMapSet
  
  PridajText "Proccesing Buffers"
  
  For a = 1 To PocetStena
   With Stena(a)
    
    If PocetStena >= a Then
     RetardBuffer(a).Vertices(0).position = Vec3(.Pos(1).X, -.Pos(1).Y, -.Pos(1).Z)
     RetardBuffer(a).Vertices(1).position = Vec3(.Pos(2).X, -.Pos(2).Y, -.Pos(2).Z)
     RetardBuffer(a).Vertices(2).position = Vec3(.Pos(0).X, -.Pos(0).Y, -.Pos(0).Z)
     RetardBuffer(a).Vertices(3).position = Vec3(.Pos(3).X, -.Pos(3).Y, -.Pos(3).Z)
     
     For b = 0 To 3
      RetardBuffer(a).Vertices(b).tu = .PocetTexturU(b)
      RetardBuffer(a).Vertices(b).TV = .PocetTexturV(b)
      RetardBuffer(a).Pos(b) = Vec3(-.Pos(b).X, .Pos(b).Z, .Pos(b).Y)
     Next b
     
     PouzivaTex(.Textura) = 1
    End If
   End With
  Next a
  
  PridajText "OK"
  PridajText ""
  PridajText "Proccesing Textures"
  
  ReDim TexturyNazov(MapSet.TextureCount)
  
  For a = 0 To MapSet.TextureCount
   If PouzivaTex(a) = 1 Then
    PouzivaTex(a) = CisloTex
    TexturyNazov(CisloTex) = XorString(EdTextury(a))
    CisloTex = CisloTex + 1
   End If
  Next a
  CisloTex = CisloTex - 1
  TMPSet.TextureCount = CisloTex
  
  ReDim Preserve TexturyNazov(CisloTex)
    
  PridajText CisloTex + 1 & " Textures OK"
  PridajText ""
  PridajText "Finishing Buffers"
  
  ReDim picTMPLight(PocetStena)
  ReDim PicCreated(PocetStena)
  
  For a = 1 To PocetStena
   RetardBuffer(a).Textura = PouzivaTex(Stena(a).Textura)
  Next a
  TMPSet.BufferCount = PocetStena
    
  PridajText PocetStena & " Buffers OK"
  
  PridajText ""
  PridajText "Proccesing Entities"
  
  For a = 1 To PocetEntity
   With EdEntity(a)
    If .Type > 0 Then TMPPocetEntit = TMPPocetEntit + 1
   End With
  Next a
  
  ReDim RetardEntity(TMPPocetEntit)
  TMPSet.EntityCount = TMPPocetEntit
  For a = 1 To TMPPocetEntit
   RetardEntity(a) = EdEntity(a)
  Next a
  
  PridajText ""
  PridajText "Proccesing Lights"
  
  Open SuborComp For Binary Access Read Write As #1
   Put #1, , TMPSet
   Put #1, , RetardBuffer
   Put #1, , RetardEntity
   Put #1, , TexturyNazov
  Close
  
  PridajText ""
  PridajText ""
  PridajText "Finished at " & Time
  PridajText "Compiling Time: " & Int((GetTickCount - StartTime) / 1000) & " s"
  CompileMap = True
  Exit Function
Chyba:

  PridajText "Failed"
End Function

Private Sub PridajText(Text As String)
  frmEditor.picProcess.Print Text
  DoEvents
End Sub
