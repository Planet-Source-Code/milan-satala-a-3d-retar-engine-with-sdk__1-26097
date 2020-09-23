Attribute VB_Name = "modConst"
Option Explicit

Public Const MoveQuality As Single = 2 'If you have got good computer incerase

Public Const Pi As Single = 3.141592653
Public Const DegTrans As Single = 180 / Pi 'You have to multiply radians with this to get degrees
Public Const Pi2 As Single = 2 * Pi

Public Const VertexSizeInBytes As Long = 24 'Size of custom vertex

'Game constants (retardentity.type)
Public Const reBullet As Byte = 1
Public Const reLight As Byte = 2
Public Const rePlayer As Byte = 5
Public Const reExplosion As Byte = 7
Public Const reHit As Byte = 8
Public Const reWeapon As Byte = 9
Public Const reAmmo As Byte = 10
