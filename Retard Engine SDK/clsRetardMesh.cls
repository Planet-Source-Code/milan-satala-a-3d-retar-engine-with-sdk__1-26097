VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRetardMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Mesh As D3DXMesh 'Mesh
Dim MeshMaterials() As D3DMATERIAL8 'Mesh material (1-8)
Dim MeshTextures() As Direct3DTexture8 'Mesh texture (1-8)
Dim NumMaterials As Long 'Number of materials and textures (between 1 and 8)

'Gues what this function do (I took source from Microsoft help)
Public Function LoadMeshFromFile(File As String) As Integer

  Dim MtrlBuffer As D3DXBuffer 'We get some things from this
  Dim i As Long
    
  'Load mesh and get # of materiasls and material buffer
  Set Mesh = D3DX.LoadMeshFromX(App.Path & "\" & File, D3DXMESH_MANAGED, D3DDevice, Nothing, MtrlBuffer, NumMaterials)

  
  ReDim MeshMaterials(NumMaterials) 'Size materials
  ReDim MeshTextures(NumMaterials) 'Size textures
    
  Dim strTexName As String 'Filename of texture
    
  For i = 0 To NumMaterials - 1 'We want to load all textures (multitexturing)
    
   D3DX.BufferGetMaterial MtrlBuffer, i, MeshMaterials(i) 'We get material from material buffer

   MeshMaterials(i).Ambient = MeshMaterials(i).diffuse 'Bla bla
   strTexName = D3DX.BufferGetTextureName(MtrlBuffer, i) 'We get texture filename from material buffer
   
   'Load texture
   Set MeshTextures(i) = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\Data\" + strTexName)
        
  Next
    
  Set MtrlBuffer = Nothing
End Function

'This sub finaly draws mesh
'[X,Y,Z - Positin in space; LookX - Rotation of mesh or where is mesh looking ...]
Public Sub DrawMesh(X As Single, Y As Single, Z As Single, LookX As Single, LookY As Single, Red As Boolean)
    Dim matDefault As D3DMATRIX 'Standart world matrix
    Dim matLook As D3DMATRIX 'Rotation
    Dim matPos As D3DMATRIX 'Position in space
    Dim matWorld As D3DMATRIX 'World matrix - by setting this we place mesh in space
    Dim i As Integer
            
    D3DXMatrixRotationYawPitchRoll matLook, LookX + Pi, LookY, 0
        
    D3DXMatrixTranslation matPos, -X, -Z, -Y
    D3DXMatrixMultiply matWorld, matLook, matPos

    D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
    For i = 0 To NumMaterials - 1 'All textures (for multitexturing)
     
     'Why tiger is red if you hit it
     If Red = True Then
      MeshMaterials(i).Ambient = MakeColor(255, 0, 0)
     Else
      MeshMaterials(i).Ambient = MakeColor(255, 255, 255)
     End If
     
     'Set texture and material
     D3DDevice.SetMaterial MeshMaterials(i)
     D3DDevice.SetTexture 0, MeshTextures(i)
        
     Mesh.DrawSubset i 'Draw mesh
    Next
    
    D3DXMatrixIdentity matDefault
    D3DDevice.SetTransform D3DTS_WORLD, matDefault 'Set standart world matrix
End Sub
