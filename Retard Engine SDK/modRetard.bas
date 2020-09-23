Attribute VB_Name = "modRetard"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long 'Returns some time in miliseconds

Public DirX8 As New DirectX8
Public D3DX As New D3DX8

Public RetardBuffer() As reRetardBuffer
Public RetardEntity() As reRetardEntity
Public RetardMesh() As clsRetardMesh

Public MapSet As reMapSettings 'Map settings
Public MapFile As String 'Filename of the map

Public GameSpeed As Single 'Speed of the game which is calculated by GetTickCount

Public CollisionPos As D3DVECTOR 'Position of collision on buffer (even if not hit buffer)
Public CollisionDist As Single 'Distance of testing point from buffer

Public Type reRetardBullet
  Speed As Single 'Bullet speed. If 0 bullet is moving "Speed of light".
  Range As Integer 'Range of bullet. It does not work in this version.
End Type

Public Type reRetardPlayer
  HP As Single  'Hit points of player. Dont work here.
  TimeToFire As Single 'Time between firing bullets
  Move As D3DVECTOR 'How much are we going to move each way
  IsHit As Boolean 'If player is hit then he is drawn red
End Type

Public Type reRetardEntity
  Type As Byte  'Type of entity. See modConst
  Pos As reRetardEntityPos 'Position and even more of entity
  Bullet As reRetardBullet 'contains all needed for bullets
  Player As reRetardPlayer 'the same but for player
End Type

Public EngineSet As reEngineSettings 'Settings for engine

Public Type reEngineSettings
  ZBuffer As Byte '1-yes 0-no
  With As Integer 'With and height of screen in fullscreen
  Height As Integer
  Bpp As Byte 'Bits per pixel
End Type

Public Type reBulletResult 'This is result of movebullet function
  Hit As Boolean 'If bullet has hit something
  Pos As D3DVECTOR 'Current position of bullet or where didi bullet hit
  On As Integer 'Number of buffer the bullet has hit
End Type

Public Type reMapSettings
  AmbientLight As Long 'Light that shines on everything
  BufferCount As Integer 'Number of buffers,textures, entities ...
  TextureCount As Integer
  EntityCount As Integer
  MeshCount As Integer
  DecalsCount As Integer
End Type

Dim OldTick As Long, OldFPS As Integer

'This sub initialize engine
Public Sub InitEngine(Hwnd As Long, Windowed As Byte)
  Dim Mode As D3DDISPLAYMODE
  Dim d3dpp As D3DPRESENT_PARAMETERS
  Dim a As Byte
    
  Set D3D = DirX8.Direct3DCreate() 'Create Direct3D object
  Set DInput = DirX8.DirectInputCreate 'Create DirectInput object
  Set DIdevice = DInput.CreateDevice("guid_SysMouse") 'Create DirectInput Mouse device
    
  D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Mode 'Get adapter format

  d3dpp.BackBufferWidth = EngineSet.With 'Screen with (if in fullscreen)
  d3dpp.BackBufferHeight = EngineSet.Height 'Screen height (if in fullscreen)
  d3dpp.BackBufferFormat = Mode.Format 'Format of adapter (maybe)
  d3dpp.hDeviceWindow = Hwnd 'Hwnd of form (or picture)
  d3dpp.BackBufferCount = 1
  d3dpp.EnableAutoDepthStencil = 1
  d3dpp.AutoDepthStencilFormat = D3DFMT_D16
  d3dpp.SwapEffect = 1
  d3dpp.Windowed = Windowed 'if 1 then screen is windowed, 0 is fullscreen
  ShowCursor Windowed 'if fullscreen then hide cursor
    
  'Create Device from d3dpp and form hwnd
  Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)

  D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
  D3DDevice.SetRenderState D3DRS_ZENABLE, 1
  D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
  
  D3DDevice.SetTextureStageState 0, D3DTSS_ADDRESSU, D3DTADDRESS_WRAP
  D3DDevice.SetTextureStageState 0, D3DTSS_ADDRESSV, D3DTADDRESS_WRAP
  
  D3DDevice.SetRenderState D3DRS_LIGHTING, 1 'Turn on lighting
  D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX 'Set vertex shader to our custom type
  
  DIdevice.SetCommonDataFormat DIFORMAT_MOUSE 'Start mouse device
  DIdevice.Acquire
  
  Set TMPVertexBuffer = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    
  'begin scene
  D3DDevice.BeginScene
End Sub

'this sub set up engine
Public Sub SetUpEngine(NewSettings As reEngineSettings)
  EngineSet = NewSettings
End Sub

'this sub set up game speed
Public Sub SetSpeed(NewSpeed As Single)
  GameSpeed = NewSpeed 'Set game speed to new game speed
End Sub

Public Sub LoadMapFromFile(File As String)
  Dim TextureName() As String 'File names of textures needed in this map
  Dim a As Integer, Name As String

  MapFile = File
  Open File For Binary Access Read Write As #1 'Open file
   Get #1, , MapSet 'Get map settings from file (Look at reMapSettings for info.)
  
   With MapSet
    ReDim RetardBuffer(.BufferCount) 'Size Retard Buffer
    ReDim VertexBuffers(.BufferCount) 'Size Retard Buffer
    ReDim RetardEntity(.EntityCount) 'Size Retard Entity
    ReDim Textures(.TextureCount) 'Size Textures
    ReDim TextureName(.TextureCount) 'Size Texture name
   End With
   
   Get #1, , RetardBuffer 'Get Retard buffers
   Get #1, , RetardEntity 'Get Retard Entities
   Get #1, , TextureName 'Get texture names
   
   For a = 0 To MapSet.TextureCount 'Load all textures
    Name = XorString(TextureName(a)) 'Name of texture is ciphertext so it is decoded
    Set Textures(a) = D3DX.CreateTextureFromFile(D3DDevice, App.Path & "\Data\" & Name) 'Load texture from file to memory
   Next a
   
   D3DDevice.SetRenderState D3DRS_AMBIENT, MapSet.AmbientLight 'set ambient light
  Close
End Sub

'[File - filename of the mesh ] Returns number of mesh
Public Function LoadMeshFromFile(File As String) As Integer
  'Create space for mesh
  MapSet.MeshCount = MapSet.MeshCount + 1
  ReDim Preserve RetardMesh(MapSet.MeshCount)
  
  Set RetardMesh(MapSet.MeshCount) = New clsRetardMesh 'Create class
  RetardMesh(MapSet.MeshCount).LoadMeshFromFile File 'Load mesh
  
  LoadMeshFromFile = MapSet.MeshCount 'Give back number of mesh
End Function


'This sub unload engine (maybe I should change sub name)
Public Sub ExitEngine()
    Erase Textures 'Delete textures
    'Erase RetardMesh 'Delete textures
    Erase VertexBuffers 'Delete buffers
    Set D3DDevice = Nothing 'unload device
    Set D3D = Nothing 'unload direct3d
    Set D3DX = Nothing 'unload direct3dx
End Sub

'this function tests collision with reRetardEntityPos
Function Collision(TestPosition As reRetardEntityPos) As Integer
  'turn reRetardEntityPos to d3dvector and try CollisionVec3
  Collision = CollisionVec3(PosToVec3(TestPosition))
End Function

'this function tests collision with d3dvector
Function CollisionVec3(TestPosition As D3DVECTOR) As Integer
  Dim a As Integer, PosVec3 As D3DVECTOR, BestDist As Single, BestPos As D3DVECTOR
  BestDist = 10 'Set best distance to high number
    
  For a = 1 To MapSet.BufferCount 'test all buffers
   'Try collision with buffer
   'and if it is true and distance is better (closer) then best distance
   If TestCollisionWithBuffer(TestPosition, RetardBuffer(a)) = True And CollisionDist < BestDist Then
    CollisionVec3 = a 'Set collision number to number of testin buffer
    BestDist = CollisionDist 'set best dist to current dist
    BestPos = CollisionPos 'set best pos to current pos
   End If
  Next a
  
  CollisionDist = BestDist 'set collision distance to best distance
  CollisionPos = BestPos 'set collision position to best pos
End Function

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'!!!!!!!!!!! SEE HELP FIRST !!!!!!!!!!
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Function TestCollisionWithBuffer(TestPosition As D3DVECTOR, TestBuffer As reRetardBuffer) As Boolean
  Dim AC As D3DVECTOR, AB As D3DVECTOR, DC As D3DVECTOR, DB As D3DVECTOR
  Dim ABC As D3DVECTOR, DBC As D3DVECTOR, TMPW As D3DVECTOR
  Dim d As Single, a As Single, t As Single
  
  With TestBuffer
  
   D3DXVec3Subtract AC, .Pos(0), .Pos(1) 'Create directional vector (0 is A, 1 is C)
   D3DXVec3Subtract AB, .Pos(0), .Pos(3) 'Create directional vector (0 is A, 3 is B)
   D3DXVec3Subtract DC, .Pos(2), .Pos(1) 'Create directional vector (2 is D, 1 is C)
   D3DXVec3Subtract DB, .Pos(2), .Pos(3) 'Create directional vector (2 is D, 3 is B)
   
   D3DXVec3Cross ABC, AC, AB 'Create ABC vector
  
   'Now tests all four sides of buffer. If not inside then don't continue
   If TestPlain(AC, .Pos(0), TestPosition) > 0 Then Exit Function
   If TestPlain(AB, .Pos(0), TestPosition) > 0 Then Exit Function
   If TestPlain(DC, .Pos(2), TestPosition) > 0 Then Exit Function
   If TestPlain(DB, .Pos(2), TestPosition) > 0 Then Exit Function
   
   'calculate t
   d = -(ABC.X * .Pos(0).X) - (ABC.Y * .Pos(0).Y) - (ABC.Z * .Pos(0).Z)
   t = -(((ABC.X * TestPosition.X) + (ABC.Y * TestPosition.Y) + (ABC.Z * TestPosition.Z) + d) / ((ABC.X * ABC.X) + (ABC.Y * ABC.Y) + (ABC.Z * ABC.Z)))
  
   'This calculate TV position on buffer
   CollisionPos.X = TestPosition.X + ABC.X * t
   CollisionPos.Y = TestPosition.Y + ABC.Y * t
   CollisionPos.Z = TestPosition.Z + ABC.Z * t
    
   'And distance between TV and buffer
   CollisionDist = GetDist3D(TestPosition, CollisionPos)
   
   'This is not always true becouse it means it just can be in collision
   'If CollisionDist is 1 and bullet is only 0.5 points big it is not in collision.
   'But if it is human (or tiger :-) which is 1.5 points big it is in collision.
   TestCollisionWithBuffer = True
  End With
End Function

Private Function TestPlain(W As D3DVECTOR, VI As D3DVECTOR, TV As D3DVECTOR) As Single
  Dim d As Single
  
  d = -(W.X * VI.X) - (W.Y * VI.Y) - (W.Z * VI.Z) 'Calculate d
  TestPlain = (W.X * TV.X) + (W.Y * TV.Y) + (W.Z * TV.Z) + d 'and test vector position to plain

End Function

'This function make variabile into variabile with game speed
'Example: Bullet is moving 2 meter per second but time between last tick is 10 ms (0.01 s, Framerate is 100 FPS) so bullet moves only 0.02 m
Public Function GS(ByVal Variabile As Single) As Single
  GS = Variabile * GameSpeed 'Calculate variabile with game speed
End Function

'This translate reRetardEntityPos to D3DVector
Public Function PosToVec3(Pos As reRetardEntityPos) As D3DVECTOR
  PosToVec3.X = Pos.X
  PosToVec3.Y = Pos.Y
  PosToVec3.Z = Pos.Z
End Function

'This function gets distace between two points (2D)
Public Function GetDist(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
  GetDist = Sqr((X1 - X2) * (X1 - X2) + (Y1 - Y2) * (Y1 - Y2))
End Function

'This function gets distace between two vectors
Public Function GetDist3D(Pos1 As D3DVECTOR, Pos2 As D3DVECTOR) As Single
  Dim a As Single, b As Single
  
  a = GetDist(Pos1.X, Pos1.Y, Pos2.X, Pos2.Y)
  b = Abs(Pos1.Z - Pos2.Z)
  
  GetDist3D = Sqr(a * a + b * b)
End Function

'this function repairs angle
Public Function RepairAngle(Angle As Single) As Single
  RepairAngle = Angle
  If Angle > 359 Then RepairAngle = Angle - 359
  If Angle < 0 Then RepairAngle = 360 + Angle
End Function

'This function code or decode any string
'For Example: XorString("Retard Engine")="Sdu`se!Dofhod" and XorString("Sdu`se!Dofhod")="Retard Engine"
Public Function XorString(Text As String) As String
  Dim a As Integer
  For a = 1 To Len(Text) 'code all
   XorString = XorString & Chr(Asc(Mid(Text, a, 1)) Xor 1) 'Code text
  Next a
End Function

Function GetAngle(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Integer
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

  GetAngle = Uhol + Poloha
End Function
