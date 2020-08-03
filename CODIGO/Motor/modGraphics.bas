Attribute VB_Name = "modGraphics"
Option Explicit

Public Sub Draw_Pixel(ByVal X As Long, ByVal Y As Long, ByVal Color As Long)

    frmMain.picMain.PSet (X, Y), Color

End Sub

Public Sub Draw_Rectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long)

    frmMain.picMain.Line (X, Y)-(X + Width, Y + Height), Color, B

End Sub

Public Sub Draw_Filled_Rectangle(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long)

    frmMain.picMain.Line (X, Y)-(X + Width, Y + Height), Color, BF

End Sub

Public Sub Draw_Circle(ByVal X As Long, ByVal Y As Long, ByVal Radius As Long, ByVal Color As Long)

    frmMain.picMain.Circle (X, Y), Radius, Color

End Sub

Public Sub Render(Map As Map_Type)

    Dim Current As Vector
    
    'Asume proper drawing order....
    
    For Current.Y = 0 To Map.Height - 1
        For Current.X = 0 To Map.Width - 1
            
            'Draw the walls
            If Map.Tile(Current.X, Current.Y) = COLLISION_WALL Then
                Draw_Filled_Rectangle Current.X * TILE_SIZE, Current.Y * TILE_SIZE, TILE_SIZE, TILE_SIZE, RGB(0, 0, 255)
            ElseIf Map.Tile(Current.X, Current.Y) = COLLISION_NONE Then
                Draw_Filled_Rectangle Current.X * TILE_SIZE, Current.Y * TILE_SIZE, TILE_SIZE, TILE_SIZE, RGB(0, 0, 0)
            End If

            'Draw the grid
            Draw_Rectangle Current.X * TILE_SIZE, Current.Y * TILE_SIZE, TILE_SIZE, TILE_SIZE, RGB(255, 255, 255)
            
        Next Current.X
    Next Current.Y

End Sub





