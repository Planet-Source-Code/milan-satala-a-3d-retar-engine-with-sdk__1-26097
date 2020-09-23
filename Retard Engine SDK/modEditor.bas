Attribute VB_Name = "modEditor"
Option Explicit
Public PocetStena As Integer, PocetSvetlo As Integer
Public SuborSave As String, SuborComp As String
Public EdMapSet As reMapSettings, EditorSet As typESet
Public PocetEntity As Integer

Private Type typESet
  PocetTextur As Integer
  
End Type

Public EdTextury() As String

Public Type typEditorStena
  Pos(3) As D3DVECTOR
  PriemX As Integer
  PriemY As Integer
  PriemZ As Integer
  Zivy As Boolean
  Lock As Boolean
  Textura As Integer
  PocetTexturU(0 To 3) As Single
  PocetTexturV(0 To 3) As Single
End Type

Public Stena(1 To 1000) As typEditorStena
Public EdEntity(1 To 1000) As reRetardEntity
Public TMPStena(3) As D3DVECTOR, TMPEnt As reRetardEntityPos

Public Drzim As Boolean, DrzimCo As Byte, Scroll As Boolean
Public CoVybrate As String, VybrateCislo As Integer
Public VybrataTextura As Integer

Public MysX As Integer, MysY As Integer, MysZ As Integer

Public MapX(2) As Integer, MapY(2) As Integer, MapZoom(2) As Integer, MapVelkost(2) As Single

Public Pos3D As reRetardEntityPos

Public Sub RefreshAll()
  Refresh2D 0
  Refresh2D 1
  Refresh2D 2
  frmEditor.NakresliTextury
  RefreshBuffersForEditor
  Refresh3D
End Sub

Public Sub MoveEntityX(Entity As reRetardEntityPos, Angle As Single, Distance As Single)
  With Entity
   .X = .X - (Distance * Sin(RepairAngle(Angle) * 3.14159 / 180))
   .Y = .Y - (Distance * Sin((90 - RepairAngle(Angle)) * 3.14159 / 180))
  End With
End Sub

Public Sub Refresh2D(Index As Byte)
  Dim a As Integer, b As Byte
  Dim PX(3) As Integer, PY(3) As Integer
  Dim EX As Integer, EY As Integer
  MapZoom(Index) = 10
  MapVelkost(Index) = frmEditor.pic2D(Index).ScaleWidth / (MapZoom(Index) * 2)
  
  frmEditor.pic2D(Index).Line (0, 0)-(1000, 1000), 0, BF
  
  For a = 0 To MapZoom(Index) * 2
   frmEditor.pic2D(Index).Line (a * MapVelkost(Index), 0)-Step(0, 1000), vbBlue
   frmEditor.pic2D(Index).Line (0, a * MapVelkost(Index))-Step(1000, 0), vbBlue
  Next a
  
  For a = 1 To PocetStena
   With Stena(a)
    If .Zivy = True Then
     
     For b = 0 To 3
      If Index = 0 Or Index = 1 Then
       PX(b) = (MapX(Index) + MapZoom(Index) + .Pos(b).X) * MapVelkost(Index)
      Else
       PX(b) = (MapX(Index) + MapZoom(Index) + .Pos(b).Y) * MapVelkost(Index)
      End If
      If Index = 0 Then
       PY(b) = (MapY(Index) + MapZoom(Index) + .Pos(b).Y) * MapVelkost(Index)
      Else
       PY(b) = (MapY(Index) + MapZoom(Index) + .Pos(b).Z) * MapVelkost(Index)
      End If
     Next b
     
     If Not (VybrateCislo = a And CoVybrate = "") Then
      For b = 0 To 2
       frmEditor.pic2D(Index).Line (PX(b), PY(b))-(PX(b + 1), PY(b + 1)), &H80000009
      Next b
      frmEditor.pic2D(Index).Line (PX(0), PY(0))-(PX(3), PY(3)), &H80000009
     End If
    End If
   End With
  Next a
  
  If VybrateCislo > 0 And CoVybrate = "" Then
   With Stena(VybrateCislo)
    For b = 0 To 3
     If Index = 0 Or Index = 1 Then
      PX(b) = (MapX(Index) + MapZoom(Index) + .Pos(b).X) * MapVelkost(Index)
     Else
      PX(b) = (MapX(Index) + MapZoom(Index) + .Pos(b).Y) * MapVelkost(Index)
     End If
     If Index = 0 Then
      PY(b) = (MapY(Index) + MapZoom(Index) + .Pos(b).Y) * MapVelkost(Index)
     Else
      PY(b) = (MapY(Index) + MapZoom(Index) + .Pos(b).Z) * MapVelkost(Index)
      End If
     Next b
     
     For b = 0 To 2
      frmEditor.pic2D(Index).Line (PX(b), PY(b))-(PX(b + 1), PY(b + 1)), vbRed
      frmEditor.pic2D(Index).Circle (PX(b), PY(b)), (2), vbWhite
     Next b
     frmEditor.pic2D(Index).Line (PX(0), PY(0))-(PX(3), PY(3)), vbRed
     frmEditor.pic2D(Index).Circle (PX(3), PY(3)), (2), vbWhite
   
    End With
  End If
  
End Sub

Public Sub Refresh3D()
  Dim a As Integer
  
  D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFF&, 1#, 0
  D3DDevice.BeginScene
  
  SetupMatrices Pos3D
  SetupLights
  D3DDevice.SetRenderState D3DRS_AMBIENT, MapSet.AmbientLight 'set ambient light
  
  For a = 1 To MapSet.BufferCount
   With RetardBuffer(a)
    D3DDevice.SetTexture 0, Textures(.Textura)
    D3DDevice.SetStreamSource 0, VertexBuffers(a), VertexSizeInBytes
    D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
   End With
  Next a
  
  D3DDevice.EndScene
  D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
  
End Sub

Public Function PridajEntitu() As Integer
  Dim a As Integer
  
  For a = 1 To PocetEntity
   If EdEntity(a).Type = 0 Then
    PridajEntitu = a
    Exit Function
   End If
  Next a
  
  PocetEntity = PocetEntity + 1
  ReDim Preserve RetardEntity(PocetEntity)
  PridajEntitu = PocetEntity
End Function

Public Sub VypocitajPriem(Cislo As Integer)
  Dim a As Byte
  
  With Stena(Cislo)
   .PriemX = 0
   .PriemY = 0
   .PriemZ = 0
   For a = 0 To 3
   .PriemX = .PriemX + .Pos(a).X
   .PriemY = .PriemY + .Pos(a).Y
   .PriemZ = .PriemZ + .Pos(a).Z
   Next a
   
   .PriemX = .PriemX / 4
   .PriemY = .PriemY / 4
   .PriemZ = .PriemZ / 4
  End With
End Sub


Public Sub RefreshBuffersForEditor()
  Dim a As Integer, b As Byte
  
  MapSet.BufferCount = PocetStena + PocetSvetlo
  MapSet.AmbientLight = vbWhite
  
  ReDim VertexBuffers(MapSet.BufferCount)
  ReDim RetardBuffer(MapSet.BufferCount)
  For a = 1 To PocetStena
   With Stena(a)
    Vertices(0).position = Vec3(.Pos(1).X, -.Pos(1).Y, -.Pos(1).Z)
    Vertices(1).position = Vec3(.Pos(2).X, -.Pos(2).Y, -.Pos(2).Z)
    Vertices(2).position = Vec3(.Pos(0).X, -.Pos(0).Y, -.Pos(0).Z)
    Vertices(3).position = Vec3(.Pos(3).X, -.Pos(3).Y, -.Pos(3).Z)
    
    If .Lock = False Then
     .PocetTexturU(2) = GetDist3D(.Pos(3), .Pos(2)) / 4
     .PocetTexturU(3) = GetDist3D(.Pos(3), .Pos(2)) / 4
     .PocetTexturV(1) = GetDist3D(.Pos(0), .Pos(3)) / 4
     .PocetTexturV(3) = GetDist3D(.Pos(0), .Pos(3)) / 4
    End If

    Vertices(2).tu = .PocetTexturU(2)
    Vertices(3).tu = .PocetTexturU(3)
    Vertices(1).TV = .PocetTexturV(1)
    Vertices(3).TV = .PocetTexturV(3)
    
    RetardBuffer(a).Textura = .Textura
    For b = 0 To 3
     RetardBuffer(a).Vertices(b) = Vertices(b)
     RetardBuffer(a).Pos(b) = Vec3(-.Pos(b).X, .Pos(b).Z, .Pos(b).Y)
    Next b
    
    Set VertexBuffers(a) = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    D3DVertexBuffer8SetData VertexBuffers(a), 0, VertexSizeInBytes * 4, 0, RetardBuffer(a).Vertices(0)
   End With
  Next a
  
End Sub

Public Function Opr(V3 As D3DVECTOR) As D3DVECTOR
  With V3
   Opr.X = .X
   Opr.Y = .Z
   Opr.Z = .Y
  End With
End Function

Function VypUhol(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Integer
  Dim Cislo1 As Long
  Dim Cislo2 As Long
  Dim Uhol As Double
  Dim Poloha As Integer
  
  If X1 = X2 And Y1 < Y2 Then
   Cislo2 = 0
   Poloha = 180
  
  
  ElseIf X1 = X2 And Y1 > Y2 Then
   Cislo2 = 0
   Poloha = 0
  ElseIf X1 < X2 And Y1 = Y2 Then
   Cislo2 = 0
   Poloha = 90
  ElseIf X1 > X2 And Y1 = Y2 Then
   Cislo2 = 0
   Poloha = 270
  ElseIf X1 < X2 And Y1 > Y2 Then
   Cislo1 = Abs(X2 - X1)
   Cislo2 = Abs(Y2 - Y1)
   Poloha = 0
  ElseIf X1 < X2 And Y1 < Y2 Then
   Cislo1 = Abs(Y1 - Y2)
   Cislo2 = Abs(X2 - X1)
   Poloha = 90
  ElseIf X1 > X2 And Y1 < Y2 Then
   Cislo1 = Abs(X1 - X2)
   Cislo2 = Abs(Y1 - Y2)
   Poloha = 180
  ElseIf X1 > X2 And Y1 > Y2 Then
   Cislo1 = Abs(Y2 - Y1)
   Cislo2 = Abs(X1 - X2)
   Poloha = 270
  End If
  
On Error GoTo Chyba
  Uhol = Atn(Cislo1 / Cislo2) * 57
Chyba:

  VypUhol = Uhol + Poloha
End Function

Function Posun_X(ByVal Uhol As Integer, ByVal Rychlost As Integer) As Double
  Posun_X = Rychlost * Sin(Uhol * 3.14159 / 180)
End Function

Function Posun_Y(ByVal Uhol As Integer, ByVal Rychlost As Integer) As Double
  Posun_Y = -Rychlost * Sin((90 - Uhol) * 3.14159 / 180)
End Function

Public Function VecAngle(V1 As D3DVECTOR, V2 As D3DVECTOR) As D3DVECTOR
  Dim Dist As Single, VOut As D3DVECTOR
  D3DXVec3Subtract VOut, V2, V1
  Dist = GetDist3D(V1, V2)
  VecAngle.X = VOut.X / Dist
  VecAngle.Y = VOut.Y / Dist
  VecAngle.Z = VOut.Z / Dist
End Function

Function Vec3(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
    Vec3.X = X
    Vec3.Y = Z
    Vec3.Z = Y
End Function

Public Function Vec3ToPos(V As D3DVECTOR) As reRetardEntityPos
  Vec3ToPos.X = V.X
  Vec3ToPos.Y = V.Y
  Vec3ToPos.Z = V.Z
End Function

Function PosVec(V As D3DVECTOR) As D3DVECTOR
  PosVec.X = -V.X
  PosVec.Y = -V.Z
  PosVec.Z = -V.Y
End Function

'This sub move entity in z
Public Sub MoveEntityZ(Entity As reRetardEntityPos, Angle As Single, Distance As Single)
  With Entity
   .Z = .Z - (Distance * Sin(RepairAngle(Angle) * 3.14159 / 180)) 'Move Z
  End With
End Sub


Public Sub DrawBoxEntity(Entity As reRetardEntityPos, R As Long, G As Long, b As Long)
  Dim Pos(0 To 7) As D3DVECTOR
  
  With Entity
   Pos(0).X = -(.X - .With / 2)
   Pos(0).Z = -(.Y + .With / 2)
   Pos(0).Y = -(.Z + .Height / 2)
   
   Pos(1).X = -(.X + .With / 2)
   Pos(1).Z = Pos(0).Z
   Pos(1).Y = Pos(0).Y
   
   Pos(2).X = Pos(1).X
   Pos(2).Z = -(.Y - .With / 2)
   Pos(2).Y = Pos(0).Y
   
   Pos(3).X = Pos(0).X
   Pos(3).Z = -(.Y - .With / 2)
   Pos(3).Y = Pos(0).Y
   
   Pos(4).X = -(.X - .With / 2)
   Pos(4).Z = -(.Y + .With / 2)
   Pos(4).Y = -(.Z - .Height / 2)
   
   Pos(5).X = -(.X + .With / 2)
   Pos(5).Z = Pos(0).Z
   Pos(5).Y = Pos(4).Y
   
   Pos(6).X = Pos(1).X
   Pos(6).Z = -(.Y - .With / 2)
   Pos(6).Y = Pos(4).Y
   
   Pos(7).X = Pos(0).X
   Pos(7).Z = -(.Y - .With / 2)
   Pos(7).Y = Pos(4).Y
   
   Dim Mtrl As D3DMATERIAL8
   Mtrl.diffuse = MakeColor(R, G, b)
   Mtrl.Ambient = Mtrl.diffuse
   D3DDevice.SetMaterial Mtrl
   D3DDevice.SetTexture 0, Nothing
   
   Vertices(2).position = Pos(0)
   Vertices(0).position = Pos(1)
   Vertices(1).position = Pos(2)
   Vertices(3).position = Pos(3)
   Set TMPVertexBuffer = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
   D3DVertexBuffer8SetData TMPVertexBuffer, 0, VertexSizeInBytes * 4, 0, Vertices(0)
   D3DDevice.SetStreamSource 0, TMPVertexBuffer, VertexSizeInBytes
   D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
   
   Vertices(2).position = Pos(4)
   Vertices(0).position = Pos(5)
   Vertices(1).position = Pos(6)
   Vertices(3).position = Pos(7)
   Set TMPVertexBuffer = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
   D3DVertexBuffer8SetData TMPVertexBuffer, 0, VertexSizeInBytes * 4, 0, Vertices(0)
   D3DDevice.SetStreamSource 0, TMPVertexBuffer, VertexSizeInBytes
   D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
   
   Vertices(2).position = Pos(0)
   Vertices(0).position = Pos(1)
   Vertices(1).position = Pos(5)
   Vertices(3).position = Pos(4)
   Set TMPVertexBuffer = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
   D3DVertexBuffer8SetData TMPVertexBuffer, 0, VertexSizeInBytes * 4, 0, Vertices(0)
   D3DDevice.SetStreamSource 0, TMPVertexBuffer, VertexSizeInBytes
   D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
      
   Vertices(2).position = Pos(3)
   Vertices(0).position = Pos(2)
   Vertices(1).position = Pos(6)
   Vertices(3).position = Pos(7)
   Set TMPVertexBuffer = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
   D3DVertexBuffer8SetData TMPVertexBuffer, 0, VertexSizeInBytes * 4, 0, Vertices(0)
   D3DDevice.SetStreamSource 0, TMPVertexBuffer, VertexSizeInBytes
   D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
   Vertices(2).position = Pos(3)
   Vertices(0).position = Pos(0)
   Vertices(1).position = Pos(4)
   Vertices(3).position = Pos(7)
   Set TMPVertexBuffer = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
   D3DVertexBuffer8SetData TMPVertexBuffer, 0, VertexSizeInBytes * 4, 0, Vertices(0)
   D3DDevice.SetStreamSource 0, TMPVertexBuffer, VertexSizeInBytes
   D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
   Vertices(2).position = Pos(2)
   Vertices(0).position = Pos(1)
   Vertices(1).position = Pos(5)
   Vertices(3).position = Pos(6)
   Set TMPVertexBuffer = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
   D3DVertexBuffer8SetData TMPVertexBuffer, 0, VertexSizeInBytes * 4, 0, Vertices(0)
   D3DDevice.SetStreamSource 0, TMPVertexBuffer, VertexSizeInBytes
   D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
    
   Mtrl.diffuse = MakeColor(255, 255, 255)
   Mtrl.Ambient = Mtrl.diffuse
   D3DDevice.SetMaterial Mtrl
    
  End With
End Sub

