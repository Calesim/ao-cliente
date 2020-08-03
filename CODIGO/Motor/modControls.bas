Attribute VB_Name = "modControls"
Option Explicit

Public Const BUTTON_UP As Long = vbKeyW
Public Const BUTTON_DOWN As Long = vbKeyS
Public Const BUTTON_LEFT As Long = vbKeyA
Public Const BUTTON_RIGHT As Long = vbKeyD

Public Const BUTTON2_UP As Long = vbKeyT
Public Const BUTTON2_DOWN As Long = vbKeyG
Public Const BUTTON2_LEFT As Long = vbKeyF
Public Const BUTTON2_RIGHT As Long = vbKeyH

Public Const BUTTON_NEUTRAL_FLAG As Long = 0
Public Const BUTTON_UP_FLAG As Long = 1
Public Const BUTTON_DOWN_FLAG As Long = 2
Public Const BUTTON_LEFT_FLAG As Long = 4
Public Const BUTTON_RIGHT_FLAG As Long = 8

Public Const BUTTON2_UP_FLAG As Long = 16
Public Const BUTTON2_DOWN_FLAG As Long = 32
Public Const BUTTON2_LEFT_FLAG As Long = 64
Public Const BUTTON2_RIGHT_FLAG As Long = 128

Public Key_State As Long

Public Mouse As Vector
Public Mouse_Held As Boolean
Public Cursor As Vector

Public Function Check_Key(Key_Flag As Long) As Long

    Check_Key = Key_State And Key_Flag

End Function

Public Sub Keyboard_Controls()
    
    If Check_Key(BUTTON_NEUTRAL_FLAG) Then
        Player.Moving = False
    Else
        Player.Moving = True
    End If
    
    If Check_Key(BUTTON_UP_FLAG) Then
        Player.Position.Y = Player.Position.Y - 1
    End If
            
    If Check_Key(BUTTON_DOWN_FLAG) Then
        Player.Position.Y = Player.Position.Y + 1
    End If
            
    If Check_Key(BUTTON_LEFT_FLAG) Then
        Player.Position.X = Player.Position.X - 1
    End If
            
    If Check_Key(BUTTON_RIGHT_FLAG) Then
        Player.Position.X = Player.Position.X + 1
    End If
    
    If Check_Key(BUTTON2_UP_FLAG) Then
        Monster.Position.Y = Monster.Position.Y - 1
    End If
            
    If Check_Key(BUTTON2_DOWN_FLAG) Then
        Monster.Position.Y = Monster.Position.Y + 1
    End If
            
    If Check_Key(BUTTON2_LEFT_FLAG) Then
        Monster.Position.X = Monster.Position.X - 1
    End If
            
    If Check_Key(BUTTON2_RIGHT_FLAG) Then
        Monster.Position.X = Monster.Position.X + 1
    End If

End Sub

Public Sub Mouse_Controls()

    If Mouse_Held = True Then
        'Process the click based on the radio button checked
        If frmMain.radStart.Value Then
            If Map.Collision_Map.Response(Cursor.X, Cursor.Y) <> COLLISION_WALL Then
                Player.Position.X = Cursor.X * TILE_SIZE
                Player.Position.Y = Cursor.Y * TILE_SIZE
                Monster.AStar_Moving = False
            End If
        ElseIf frmMain.radEnd.Value Then
            If Map.Collision_Map.Response(Cursor.X, Cursor.Y) <> COLLISION_WALL Then
                Monster.Position.X = Cursor.X * TILE_SIZE
                Monster.Position.Y = Cursor.Y * TILE_SIZE
                Monster.AStar_Moving = False
            End If
        ElseIf frmMain.radWall.Value = 1 Then
            Map.Tile(Cursor.X, Cursor.Y) = 1
            Map.Collision_Map.Response(Cursor.X, Cursor.Y) = COLLISION_WALL
        ElseIf frmMain.radWall.Value = 0 And frmMain.radStart.Value = 0 And frmMain.radEnd.Value = 0 Then
            frmMain.Caption = Cursor.X & ", " & Cursor.Y
            Map.Tile(Cursor.X, Cursor.Y) = 0
            Map.Collision_Map.Response(Cursor.X, Cursor.Y) = COLLISION_NONE
        End If
        Monster.Compute_AStar_Enabled = True
        Clear_AStar Map
    End If

End Sub


