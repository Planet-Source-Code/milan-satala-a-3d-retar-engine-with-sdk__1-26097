Attribute VB_Name = "modGame"
Option Explicit

'This module contains all game source, everything else is retard engne

Dim OldTime As Long, GrafPos As reRetardEntityPos 'Old time and viewport
Public PlayerEnt As Integer 'Player entity number

'This sub contains all calculates in the game
Public Sub Game()
  Dim Time As Long 'Time, in miliseconds betwen last tick
  Dim b As Integer, a As Integer
  Dim MovePos As reRetardEntityPos 'Moveing pos, it is not where does we want to move, but how much we want to move
  Dim FireWeapon As Boolean
  
  DoEvents
  
  Time = GetTickCount - OldTime 'Get time

  If Time > 500 Then 'if time is bigger then 0.5 sec (for example while breaking) than dont do anything
   OldTime = GetTickCount 'Set oldtime
   Exit Sub 'Exit sub
  End If
  If Time = 0 Then Exit Sub 'If time is 0 than dont do anything (game speed will do some erros)
  
  OldTime = GetTickCount 'Set oldtime
  
  SetSpeed Time / 1000 'Set game speed
  
  If KeyState(vbKeyEscape) = 1 Then End 'if escape is pressed than end the game

  
  With RetardEntity(PlayerEnt) 'Player entity
   'Move pos is not where we want to move but how much we want to move if movepos.x=1 then (later) try to move one point to the right
   MovePos.Z = .Player.Move.Z 'X and Y are 0 but Z is is keeped becouse it is needed for jumping
   
   If KeyState(vbKeyW) = 2 Then MoveEntityXY MovePos, .Pos.LookDegX, 12 'Move forward
   If KeyState(vbKeyS) = 2 Then MoveEntityXY MovePos, .Pos.LookDegX + 180, 12 'Move back
   If KeyState(vbKeyA) = 2 Then MoveEntityXY MovePos, .Pos.LookDegX - 90, 7 'Strafe left
   If KeyState(vbKeyD) = 2 Then MoveEntityXY MovePos, .Pos.LookDegX + 90, 7 'Strafe Right
   If KeyState(vbKeyTab) = 1 Then frmGame.WindowState = vbMinimized 'Minimalize the form
   
   If KeyState(vbKeySpace) = 1 And .Pos.StandOn Then MovePos.Z = -7 'if the space key is pressed and we are standing on something then jump
   
   If KeyState(vbKeyShift) = 2 And .Player.TimeToFire <= 0 Then 'Fire weapon
    FireWeapon = True
    .Player.TimeToFire = 0.1 'Time to next fire
   End If
   
   'Set how much to move
   .Player.Move = PosToVec3(MovePos)
   
   GrafPos = .Pos 'Set viewport to our current entity pos
   GrafPos.Z = GrafPos.Z - 2 'Moves viewport 2 points over the ground (but if you want to feel like a dog then delete this line)
  
   Render GrafPos 'Render graphics from our viewport
   DIdevice.GetDeviceStateMouse MouseState 'Get mouse state
   TurnEntityX .Pos, MouseState.lX / 2 'Turn around x axis
   TurnEntityY .Pos, -(MouseState.lY / 2) 'Turn around y axis (delete the minus if you are using invert mouse)
   
  End With
   
  If FireWeapon Then 'if fire weapon is true then why not to fire ?
   FireBulletFromPos GrafPos 'Fire weapon from our current viewport position
  End If
    
  'Entities !!!
  For a = 1 To MapSet.EntityCount 'Lets calculate every entity
   With RetardEntity(a)
    If .Type Then EntityPosLook .Pos 'if entity does exist, than calculate all things, which we will need later
    Select Case .Type
     Case reBullet ' if the type is bullet then ...
     
      Dim Result As reBulletResult 'Declare result of bullet flight
      Result = MoveBullet(.Pos, .Bullet.Speed, 0) 'Calculate bullet flight
      
      If Result.Hit = True Or .Bullet.Speed = 0 Then 'if the bullet has hit something then
       .Type = 0 'Turn entity off
      End If
      
     Case rePlayer ' if the type is player then ...
      If PlayerEnt <> a Then 'If it is not you then draw tiger and make AI calculates
       RetardMesh(1).DrawMesh .Pos.X, .Pos.Y, .Pos.Z + 1, .Pos.LookX, .Pos.LookY, .Player.IsHit
       BotAI a
       .Player.IsHit = False
      End If
      CalculatePlayer a '... Calculate player
    End Select
   End With
  Next a
End Sub

Private Sub BotAI(Num As Integer)
  Dim Result As D3DVECTOR, TargetAngle As Integer, MovePos As reRetardEntityPos
  
  With RetardEntity(Num)
   Result = FindWay(.Pos, PosToVec3(RetardEntity(PlayerEnt).Pos)) 'Find best (ehm) way
   If Result.X Or Result.Y Then ' If it is not in position
    TargetAngle = GetAngle(0, 0, Result.X, Result.Y) 'find angle to turn
    .Pos.LookDegX = -TargetAngle 'turn there
    MoveEntityXY MovePos, .Pos.LookDegX, 1 'Move our move (?)
    .Player.Move.X = MovePos.X
    .Player.Move.Y = MovePos.Y
   End If
  End With
End Sub

'This sub calculate player (in original version it is much more complexive)
Public Sub CalculatePlayer(Num As Integer)
  With RetardEntity(Num)
   'If we are jumping this makes us stop flying up and make us flying down
   If .Player.Move.Z < 10 Then .Player.Move.Z = .Player.Move.Z + GS(20)
   MoveEntity SpeedVector(.Player.Move), .Pos 'Move player
   'Calculate time to fire
   If .Player.TimeToFire > 0 Then .Player.TimeToFire = .Player.TimeToFire - GS(1) Else .Player.TimeToFire = 0
  End With
End Sub

Public Sub FireBulletFromPos(Pos As reRetardEntityPos)
  Dim EntNum As Integer 'Number of bullet entity
  
  EntNum = AddEntity 'Create entity for bullet
  
  With RetardEntity(EntNum) 'Bullet entity
   .Pos = Pos 'set pos to pos where the bullet was fired from
   .Pos.With = 0.1 'With of bullet is 0.1
   .Pos.Height = 0.1 'Height of bullet is 0.1
   .Bullet.Speed = 0 'Bullet speed is 10
   .Bullet.Range = 5 'Range of fire is 5
   .Type = reBullet 'Type of entity is bullet
  End With
End Sub
