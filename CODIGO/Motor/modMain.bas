Attribute VB_Name = "modMain"
Option Explicit

Public Running As Boolean
Public Done As Boolean

Public Sub Game_Loop()
    
    Do While Running = True
        Keyboard_Controls
        Mouse_Controls
        Collision_Detection
        Get_Player_Info
        Get_Sprite_Info Monster
        AStar_Find_Path Map, Monster, Player
        Follow_AStar_Path Map, Monster, 1
        frmMain.picMain.Cls
        Render Map
        Draw_Sprite Monster, RGB(255, 0, 0)
        Draw_Player RGB(0, 255, 0)
        Draw_AStar_Path Monster
        'frmMain.Caption = Player.Coordinates.X & ", " & Player.Coordinates.Y & " ----- " & Player.Center_Coordinates.X & ", " & Player.Center_Coordinates.Y
        Lock_Framerate 60
        DoEvents
    Loop

End Sub

Public Sub Main2()

    With frmMain
        .Show
        With .picMain
            .AutoRedraw = True
            .BackColor = RGB(0, 0, 0)
            .ScaleMode = vbPixels
            .ScaleWidth = 375
            .ScaleHeight = 375
        End With
    End With
    
    Hi_Res_Timer_Initialize
    
    Player.Position.X = 1 * TILE_SIZE: Player.Position.Y = 1 * TILE_SIZE
    Monster.Position.X = 22 * TILE_SIZE: Monster.Position.Y = 1 * TILE_SIZE
    Map_Setup
    Running = True
    Game_Loop

End Sub

Public Sub Shutdown()

    Running = False
    Unload frmMain
    
End Sub

