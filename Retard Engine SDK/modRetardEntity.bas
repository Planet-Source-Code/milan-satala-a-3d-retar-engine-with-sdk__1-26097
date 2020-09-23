Attribute VB_Name = "modRetardEntity"
Option Explicit

Dim AllowedMove As Boolean

'This sub move entity in x and y
Public Sub MoveEntityXY(Entity As reRetardEntityPos, Angle As Single, Distance As Single)
  With Entity
   .X = .X - (Distance * Sin(RepairAngle(Angle) * 3.14159 / 180))   'Move X
   .Y = .Y - (Distance * Sin((90 - RepairAngle(Angle)) * 3.14159 / 180)) 'Move Y
  End With
End Sub

'This sub moves entity in 3D
Public Sub MoveEntity3D(Entity As reRetardEntityPos, Distance As Single)
  With Entity
   .X = .X + .DirVec.X * Distance 'Move X
   .Y = .Y + .DirVec.Y * Distance 'Move Y
   .Z = .Z + .DirVec.Z * Distance 'Move Z
  End With
End Sub

'this function create entity and return its number (Entity architecture is explained at start of this module)
Public Function AddEntity() As Integer
  Dim a As Integer
  
  For a = 1 To MapSet.EntityCount 'Test all entities
   If RetardEntity(a).Type = 0 Then 'if entity is inactive
    AddEntity = a 'then number of "created" entity is a
    Exit Function 'exit function becouse we have entity
   End If
  Next a
  
  'if all entities are active, then we have to create a new one
  MapSet.EntityCount = MapSet.EntityCount + 1 'Raise entity count
  ReDim Preserve RetardEntity(MapSet.EntityCount) 'size Retard entity
  AddEntity = MapSet.EntityCount 'number of created entity is entity count
End Function

'this function moves bullet
Public Function MoveBullet(Entity As reRetardEntityPos, Speed As Single, Range As Integer) As reBulletResult
  Dim a As Integer, Move As Single, Result As Integer, Dist As Single
  
  If Speed = 0 Then 'if speed is 0 then the bullet hit at the moment of its fire (something like railgun)
   Dist = 200 'Distance to travel is range of fire
  Else 'if bullet is normal bullet that flights (like rocket)
   Dist = GS(Speed) 'Distace to travel is speed
  End If
  
  With Entity
   Do Until Dist <= 0 'if distance is 0 or less then exit
    
    Result = Collision(Entity) 'Result is number of nearest buffer
        
    If Result And CollisionDist < 0.1 Then 'if distance is smaller than 0.1 then bullet has hit it
     MoveBullet.Hit = True 'bullet has hit
     MoveBullet.Pos = CollisionPos 'Position of hit is collision pos
     MoveBullet.On = Result 'number of buffer, which has been hit
     Exit Function
    End If
        
    If Result And CollisionDist < 1 And Dist > 0.1 Then 'if we are close to some buffer then move slowly
     Move = 0.1 'move speed is 0.1
    ElseIf Dist < 1 Then 'if dist is smaller than 1 then move that dist
     Move = Dist 'move speed is distane left
    Else 'if bullet isnt near any buffer and distance is bigger than 1
     Move = 1 'move speed is 1 (default)
    End If
    
    MoveEntity3D Entity, Move 'Move entity in 3D with move speed
    Dist = Dist - Move 'Calculate distance
    
    '
    If GetDist3D(PosToVec3(Entity), PosToVec3(RetardEntity(2).Pos)) < 2 Then RetardEntity(2).Player.IsHit = True
   Loop
   
  End With
End Function

'This sub moves entity !!!explanation of this sub is in html help!!!
'!!!It is highely recomended to first run html help!!!
Public Sub MoveEntity(Move As D3DVECTOR, Entity As reRetardEntityPos)
  Dim a As Integer, Count As Integer, Result As Integer
  Dim Dist As Single, TMPPos As reRetardEntityPos
    
  Dist = GetDist(0, 0, Move.X, Move.Y) 'get distance, we have to travel
  
  'First we move on x and then on y, it make us not to stop, when we come to wall
  'but let walk around the wall
  With Entity
  
   TMPPos = Entity 'Set temp position to our current
   TMPPos.X = TMPPos.X + Move.X 'Move in x direction
   Result = MoveEntityDown(TMPPos, 2, Move.Z) 'Try to move
   If AllowedMove Then 'If the move was allowed
    Entity = TMPPos 'than move our current position to temp pos
    .StandOn = Result 'and make as stand on something (if not then value is 0)
   End If
   
   TMPPos = Entity 'Set temp position to our current
   TMPPos.Y = TMPPos.Y + Move.Y  'Move in x direction
   Result = MoveEntityDown(TMPPos, 2, Move.Z) 'Try to move
   If AllowedMove Then 'If the move was allowed
    Entity = TMPPos 'than move our current position to temp pos
    .StandOn = Result 'and make as stand on something (if not then value is 0)
   End If
  End With
End Sub

'This function moves from up to down (for walking on the stairs),
'and when get in collision with somenthing left us there and make us stand on it
Private Function MoveEntityDown(EntityPos As reRetardEntityPos, Up As Single, MoveZ As Single) As Integer
  Dim a As Single, Count As Single, OldZ As Single
  Dim Now As D3DVECTOR, Result As Integer
  
  With EntityPos
   AllowedMove = False 'We cant move there (for now)
   Count = Abs(Up + MoveZ) 'get how many times we will move down (1 point each time)
   OldZ = .Z
   .Z = .Z - Up 'Moves us up (for stairs steping)
   
   Now = PosToVec3(EntityPos) 'Set our testing pos to our current pos
   For a = 1 To Count
    
    Now.Z = Now.Z + 1 'Move us down
    
    Result = CollisionVec3(Now) 'Test collision, result is number of buffer we are in collision with
    If Result = 0 Or (CollisionDist > 1.5 And Result) Then 'if we are not at any buffer or we are but distance from us to it is bigger than 1.5 (our size)
     EntityPos.Z = Now.Z 'Move us to new position
    Else 'if we are too close to buffer
     If a > 1 Then AllowedMove = True 'if a>1 then let us move there if not then the wall is too high
     MoveEntityDown = Result 'we are standing on the buffer, that stopped us
     Exit Function 'exit function
    End If
   Next a
   
   'Now we do all this same BUT:
   'count is not allways whole number and we have to move even with that rest
   '(for example count = 2.5 we move 2 times, but what with 0.5 ? This resolve it)
   
   Now.Z = Now.Z + (Count - a + 1)  'Move us down, but only with rest of the number
    
   Result = CollisionVec3(Now) 'Test collision, result is number of buffer we are in collision with
   If Result = 0 Or (CollisionDist > 1.5 And Result) Then 'if we are not at any buffer or we are but distance from us to it is bigger than 1.5
    EntityPos.Z = Now.Z 'Move us to new position
   Else 'if we are to close to buffer
    If a > 1 Then AllowedMove = True 'if a>1 than let us move there if not then the wall is too high
    MoveEntityDown = Result 'we are standing on the buffer, that stopped us
    Exit Function 'exit function
   End If
      
   'if we get here (what means we are falling) then move is allowed
   AllowedMove = True
  End With
End Function

'This sub turn entity around X axis
Public Sub TurnEntityX(Entity As reRetardEntityPos, Angle As Single)
  With Entity
   .LookDegX = .LookDegX + Angle 'Add angle to entity angle
   If .LookDegX >= 360 Then .LookDegX = .LookDegX - 360 'repair angle if it is bigger than 360
   If .LookDegX < 0 Then .LookDegX = .LookDegX + 360 'repair angle if it is less than 0
  End With
End Sub

'This sub turn entity around X axis
Public Sub TurnEntityY(Entity As reRetardEntityPos, Angle As Single)
  With Entity
   .LookDegY = .LookDegY + Angle 'Add angle to entity angle
   If .LookDegY > 90 Then .LookDegY = 90 'if angle is bigger than 90 then it is 90
   If .LookDegY < -90 Then .LookDegY = -90 'if angle is less than -90 then it is -90
  End With
End Sub

'This function calculate vector with game speed
Public Function SpeedVector(V As D3DVECTOR) As D3DVECTOR
  SpeedVector.X = GS(V.X)
  SpeedVector.Y = GS(V.Y)
  SpeedVector.Z = GS(V.Z)
End Function

'Some important thing for faster run (I use these things many times and I dont want to calculate them many times so they are calculated only once here)
Public Sub EntityPosLook(EntityPos As reRetardEntityPos)
  Dim MoveZ As Single
  With EntityPos
   'Direct X uses Radians
   .LookX = .LookDegX / DegTrans 'get LookX (in Radians) from LookDegX (in Degrees)
   .LookY = .LookDegY / DegTrans 'get LookY (in Radians) from LookDegY (in Degrees)
     
   'Create DirVec (Look on Html)
   .DirVec.Z = -(Sin(.LookDegY * 3.14159 / 180))
   MoveZ = Sin((90 - .LookDegY) * 3.14159 / 180)
   .DirVec.X = -(MoveZ * Sin(.LookDegX * 3.14159 / 180))
   .DirVec.Y = -(MoveZ * Sin((90 - .LookDegX) * 3.14159 / 180))
  End With
End Sub
