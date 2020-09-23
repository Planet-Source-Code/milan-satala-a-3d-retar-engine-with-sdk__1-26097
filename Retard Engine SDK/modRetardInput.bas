Attribute VB_Name = "modRetardInput"
Option Explicit

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Const KEY_TOGGLED As Integer = &H1
Private Const KEY_DOWN As Integer = &H1000

Dim TimeToEnd As Boolean
Dim LastState(0 To 255) As Byte

Public DInput As DirectInput8
Public DIdevice As DirectInputDevice8
Public MouseState As DIMOUSESTATE

'This function was programed by  my friend Tomas Banovec
'It give you a state of key: 1 - KeyDown, 2 - Hold, 3 - KeyUp
Public Function KeyState(ByVal m_Key As Byte) As Byte
 If (GetKeyState(m_Key) And KEY_DOWN) Then KeyState = 1
 If LastState(m_Key) > 0 And KeyState = 1 Then KeyState = 2
 If LastState(m_Key) = 3 Then LastState(m_Key) = 0
 If LastState(m_Key) > 0 And KeyState = 0 Then KeyState = 3
 LastState(m_Key) = KeyState
End Function
