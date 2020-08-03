Attribute VB_Name = "modSprite"
Option Explicit

Public Type Node_Type

    OCList As Long
    G As Long
    H As Long
    F As Long
    X As Long
    Y As Long
    
End Type

Public Type Heap_Type

    Score As Long
    X As Long
    Y As Long
    
End Type

Public Type Sprite_Type

    Position As Vector
    Center_Position As Vector
    Previous_Position As Vector
    Previous_Coordinates As Vector
    Previous_Coordinates_Position As Vector
    Coordinates As Vector
    Coordinates_Position As Vector
    Center_Coordinates As Vector
    Center_Coordinates_Position As Vector
    Previous_Center_Coordinates As Vector
    Previous_Center_Coordinates_Position As Vector
    
    'Collision Stuff
    Collided As Boolean
    NColl As Vector
    DColl As Single
    Moving As Boolean
    
    'AI stuff
    Compute_AStar_Enabled As Boolean
    Length_Of_AStar_Path As Long
    Current_Path As Long
    Nodes() As Node_Type
    Size_Of_Heap As Long 'Size of the heap array
    Heap() As Heap_Type 'Heap Array
    AStar_Path() As Vector
    Path_Found As Boolean
    Path_Hunt As Boolean
    Vec As Vector
    AStar_Moving As Boolean
    
End Type

Public Sub Convert_Position_To_Coordinates(Sprite As Sprite_Type)

    Sprite.Coordinates.X = Int(Sprite.Position.X / TILE_SIZE)
    Sprite.Coordinates.Y = Int(Sprite.Position.Y / TILE_SIZE)

End Sub

Public Sub Get_Player_Info()
    
    'NOTE: FIX CODE CAUSE ASTAR IS DEPENDENT ON PLAYER / MONSTER COORDINATES TO FIRE ASTAR
    
    With Player
        If .Previous_Center_Coordinates.X <> .Center_Coordinates.X Or .Previous_Center_Coordinates.Y <> .Center_Coordinates.Y Then
            .Previous_Coordinates.X = .Coordinates.X
            .Previous_Coordinates.Y = .Coordinates.Y
            .Previous_Center_Coordinates.X = .Center_Coordinates.X
            .Previous_Center_Coordinates.Y = .Center_Coordinates.Y
            .Previous_Coordinates_Position.X = .Coordinates_Position.X
            .Previous_Coordinates_Position.Y = .Coordinates_Position.Y
            .Previous_Center_Coordinates_Position.X = .Center_Coordinates_Position.X
            .Previous_Center_Coordinates_Position.Y = .Center_Coordinates_Position.Y
            .Compute_AStar_Enabled = True
            Clear_AStar Map
        Else
            .Compute_AStar_Enabled = False
        End If
        .Coordinates.X = Int(.Position.X / TILE_SIZE)
        .Coordinates.Y = Int(.Position.Y / TILE_SIZE)
        .Center_Coordinates.X = Int((.Position.X + (TILE_SIZE / 2)) / TILE_SIZE)
        .Center_Coordinates.Y = Int((.Position.Y + (TILE_SIZE / 2)) / TILE_SIZE)
        .Coordinates_Position.X = Int(.Position.X / TILE_SIZE) * TILE_SIZE
        .Coordinates_Position.Y = Int(.Position.Y / TILE_SIZE) * TILE_SIZE
        .Center_Coordinates_Position.X = Int(.Position.X / TILE_SIZE) * TILE_SIZE + (TILE_SIZE / 2)
        .Center_Coordinates_Position.Y = Int(.Position.Y / TILE_SIZE) * TILE_SIZE + (TILE_SIZE / 2)
        .Center_Position.X = .Position.X + (TILE_SIZE / 2)
        .Center_Position.Y = .Position.Y + (TILE_SIZE / 2)
    End With

End Sub

Public Sub Get_Sprite_Info(Sprite As Sprite_Type)

    Static Temp As Vector

    With Sprite
        If .Position.X <= 0 Then .Position.X = 0
        If .Position.Y <= 0 Then .Position.Y = 0
        If .Position.X >= (Map.Width - 1) * TILE_SIZE Then .Position.X = (Map.Width - 1) * TILE_SIZE
        If .Position.Y >= (Map.Height - 1) * TILE_SIZE Then .Position.Y = (Map.Height - 1) * TILE_SIZE
        If Temp.X <> .Center_Coordinates.X Or Temp.Y <> .Center_Coordinates.Y Then
            Temp.X = .Center_Coordinates.X
            Temp.Y = .Center_Coordinates.Y
           ' .Compute_AStar_Enabled = True
           ' Clear_AStar Map
        Else
            .Compute_AStar_Enabled = False
        End If
        .Coordinates.X = Int(.Position.X / TILE_SIZE)
        .Coordinates.Y = Int(.Position.Y / TILE_SIZE)
        .Center_Coordinates.X = Int((.Position.X + (TILE_SIZE / 2)) / TILE_SIZE)
        .Center_Coordinates.Y = Int((.Position.Y + (TILE_SIZE / 2)) / TILE_SIZE)
        .Coordinates_Position.X = Int(.Position.X / TILE_SIZE) * TILE_SIZE
        .Coordinates_Position.Y = Int(.Position.Y / TILE_SIZE) * TILE_SIZE
        .Center_Coordinates_Position.X = Int(.Position.X / TILE_SIZE) * TILE_SIZE + (TILE_SIZE / 2)
        .Center_Coordinates_Position.Y = Int(.Position.Y / TILE_SIZE) * TILE_SIZE + (TILE_SIZE / 2)
        .Center_Position.X = .Position.X + (TILE_SIZE / 2)
        .Center_Position.Y = .Position.Y + (TILE_SIZE / 2)
    End With

End Sub

Public Sub Draw_Player(Color As Long)

    With Player
        Draw_Filled_Rectangle .Position.X, .Position.Y, TILE_SIZE, TILE_SIZE, Color
    End With
        
End Sub

Public Sub Draw_Sprite(Sprite As Sprite_Type, Color As Long)
        
    With Sprite
        Draw_Filled_Rectangle .Position.X, .Position.Y, TILE_SIZE, TILE_SIZE, Color
    End With
        
End Sub

