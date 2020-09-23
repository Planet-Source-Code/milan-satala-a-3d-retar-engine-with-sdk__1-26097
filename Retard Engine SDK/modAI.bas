Attribute VB_Name = "modRetardAI"
Option Explicit

'This function search way to target position
'It is not very good but it can come round simple barrier
Public Function FindWay(Pos As reRetardEntityPos, Target As D3DVECTOR) As D3DVECTOR
  Dim a As Integer, b As Integer
  Dim TMPVec As D3DVECTOR
  Dim Dist As Single, BestDist As Single
  
  BestDist = 100000 'set best distance to big number
  'Search around our position
  For a = -1 To 1
  For b = -1 To 1
   TMPVec.X = Pos.X + a
   TMPVec.Y = Pos.Y + b
   TMPVec.Z = Pos.Z - 2 'for moving on the steps
   
   If CollisionVec3(TMPVec) = 0 Or CollisionDist > 1.5 Then 'tests if can move here
    Dist = GetDist3D(TMPVec, Target) 'distance from here to place we want to go
    'if it is here closer to the target then move here
    If Dist < BestDist Then
     BestDist = Dist
     FindWay.X = a
     FindWay.Y = b
    End If
   End If
  Next b
  Next a
End Function
