Attribute VB_Name = "modRetard3D"
Option Explicit

Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public D3D As Direct3D8
Public D3DDevice As Direct3DDevice8

Public Type CUSTOMVERTEX
    position As D3DVECTOR 'Position
    Color As Long 'I dont know what is this good for but if I give it away the game crash
    tu As Single 'Texture coordinates
    TV As Single
End Type

'This is Vertex shader. position is D3DFVF_XYZ color is D3DFVF_DIFFUSE and tu,tv is D3DFVF_TEX1
Public Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Public Vertices(0 To 3) As CUSTOMVERTEX
Public VertexBuffers() As Direct3DVertexBuffer8 'Compiled retard buffers for DX

Public Textures() As Direct3DTexture8 'Texture are keeped here

Public TMPVertexBuffer As Direct3DVertexBuffer8 'This is needed in editor

Public Type reRetardEntityPos
  X As Single 'X,Y,Z position in space
  Y As Single
  Z As Single
  LookX As Single 'Where is entity looking in degrees
  LookY As Single
  LookDegX As Single '... and radians (for DX)
  LookDegY As Single
  DirVec As D3DVECTOR 'Where is entity looking in vector. If use getdist3d function the distance will be always 1
  Height As Single 'Size of entity
  With As Single
  StandOn As Integer 'What is entity standing on. 0 if flying ...
End Type

Public Type reRetardBuffer
  Textura As Integer 'Number of texture
  Vertices(0 To 3) As CUSTOMVERTEX 'Things needed for rendering
  Pos(0 To 3) As D3DVECTOR 'Position of each point (I am not using same positions like DX)
End Type

'This function place and turn our viewport
Public Sub SetupMatrices(LookBy As reRetardEntityPos)
    Dim matView As D3DMATRIX
    Dim matRotation As D3DMATRIX
    Dim matPitch As D3DMATRIX
    Dim matLook As D3DMATRIX
    Dim matPos As D3DMATRIX
    Dim matWorld As D3DMATRIX
    Dim matProj As D3DMATRIX
    
    D3DXMatrixIdentity matWorld
    D3DDevice.SetTransform D3DTS_WORLD, matWorld 'set world matrix. Not important in this version
    
    D3DXMatrixRotationY matRotation, -LookBy.LookX 'Add look (x axis)
    D3DXMatrixRotationX matPitch, LookBy.LookY 'Add look (y axis)
    
    D3DXMatrixMultiply matLook, matRotation, matPitch 'Create matrix which contains where are we looking
    D3DXMatrixTranslation matPos, LookBy.X, LookBy.Z, LookBy.Y 'Create matrix with our position
    D3DXMatrixMultiply matView, matPos, matLook 'Make one from look and pos
    
    D3DDevice.SetTransform D3DTS_VIEW, matView 'and set it
    
    'Set fov, you may use it for zooming pi/4 is default and smaller the number is, the bigger the zoom is
    D3DXMatrixPerspectiveFovLH matProj, Pi / 4, 1, 1, 10000
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj ':-(
End Sub

'This sub setups lights
Sub SetupLights()
  Dim Mtrl As D3DMATERIAL8
  'Lower number means darker game
  Mtrl.Ambient = MakeColor(255, 255, 255) 'Create ambient color on material
  D3DDevice.SetMaterial Mtrl 'Set material
End Sub

'This sub refresh all buffers
Public Sub RefreshBuffers()
  Dim a As Integer
  
  For a = 1 To MapSet.BufferCount 'calculate all buffers
   'Create new vertex buffer
   Set VertexBuffers(a) = D3DDevice.CreateVertexBuffer(VertexSizeInBytes * 4, 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
   'Set data to vertex buffer
   D3DVertexBuffer8SetData VertexBuffers(a), 0, VertexSizeInBytes * 4, 0, RetardBuffer(a).Vertices(0)
  Next a
End Sub

'This sub render graphics
Public Sub Render(LookBy As reRetardEntityPos)
  Dim a As Integer
     
  'Maybe it looks strange for someone, why do we begin with end scene
  'It is becouse here in this sub we draw only buffers and all other graphics (Weapons, bullets ...)
  'is drawn during the game (for higher preformacne)
  D3DDevice.EndScene
  
  D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0 'Draw everything on the form
  D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFF&, 1#, 0  'Draw blue color
  D3DDevice.BeginScene 'Begin scene
  
  SetupMatrices LookBy 'Create viewport from pos
  SetupLights 'Set up lights
           
  D3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX 'Set vertext to our custom type
  
  For a = 1 To MapSet.BufferCount 'Draw all buffers
   With RetardBuffer(a)
        
    D3DDevice.SetTexture 0, Textures(.Textura) 'Set texture for this buffer
    
    
    D3DDevice.SetStreamSource 0, VertexBuffers(a), VertexSizeInBytes  'Create buffer in memory
    D3DDevice.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2 'And draw it
   End With
  Next a
End Sub

'This function create color
Public Function MakeColor(R As Long, G As Long, b As Long) As D3DCOLORVALUE
  MakeColor.R = R / 255
  MakeColor.G = G / 255
  MakeColor.b = b / 255
  MakeColor.a = 1
End Function
