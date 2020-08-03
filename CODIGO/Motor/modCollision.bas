Attribute VB_Name = "modCollision"
Option Explicit

Public Type Collision_Type

    Width As Single 'Same as map width
    Height As Single 'Same as map height
    Map() As String 'Only used for hardcoding maps. For loading maps you would just need Response()
    Response() As Long 'Your collision type. 0 for COLLISION_NONE. 1 for COLLISION_WALL. Other values can be for water, lava, etc.
    Vertex_List() As Vector
    
End Type

'You are welcome to adding more collision types such as COLLISION_WATER, COLLISION_LAVA, etc., to have predators avoid em, but you need to modify the AStar code a notch and
'program it in. To do this just copy and paste where the COLLISION_WALL code is and replace it with water, lava, etc.

Public Const COLLISION_NONE As Long = 0
Public Const COLLISION_WALL As Long = 1

Public Function Collision_Box_To_Box(ByVal B1_X1 As Single, ByVal B1_Y1 As Single, ByVal B1_X2 As Single, ByVal B1_Y2 As Single, _
                                     ByVal B2_X1 As Single, ByVal B2_Y1 As Single, ByVal B2_X2 As Single, ByVal B2_Y2 As Single) As Boolean
                                     
    'Collision_Box_To_Box = ((Abs(B1_X - B2_X) * 2) < (B1_Width + B2_Width)) And ((Abs(B1_Y - B2_Y) * 2) < (B1_Height + B2_Height))
    
    If B1_X1 < B2_X2 And _
       B1_X2 > B2_X1 And _
       B1_Y1 < B2_Y2 And _
       B1_Y2 > B2_Y1 Then
           
        Collision_Box_To_Box = True
        
    End If
    
End Function

Public Function Collision_Box_To_Box2(ByVal B1_X As Single, ByVal B1_Y As Single, B1_Width As Single, B1_Height As Single, ByVal B2_X As Single, ByVal B2_Y As Single, B2_Width As Single, B2_Height As Single) As Long
    
    Const NO_COLLISION As Long = 0
    Const COL_LEFT As Long = 1
    Const COL_RIGHT As Long = 2
    Const COL_UP As Long = 3
    Const COL_DOWN As Long = 4
    
    Dim Side As Long
    Dim Overlap As Long
    
    If Not (B1_X < (B2_X + B2_Width) And _
       (B1_X + B1_Width) > B2_X And _
       B1_Y < (B2_Y + B2_Height) And _
       (B1_Y + B1_Height) > B2_Y) Then
        Collision_Box_To_Box2 = 0
        Exit Function
    End If
    
    Side = COL_LEFT
    Overlap = Abs(B1_X - (B2_X + B2_Width))

    If Abs((B1_X + B1_Width) - B2_X) < Overlap Then
        Side = COL_RIGHT
        Overlap = Abs((B1_X + B1_Width) - B2_X)
    End If
    
    If Abs(B1_Y - (B2_Y + B2_Height)) < Overlap Then
        Side = COL_UP
        Overlap = Abs(B1_Y - (B2_Y + B2_Height))
    End If
    
    If Abs((B1_Y + B1_Height) - B2_Y) < Overlap Then
        Side = COL_DOWN
        Overlap = Abs((B1_Y + B1_Height) - B2_Y)
    End If

    Collision_Box_To_Box2 = Side
    
End Function

Public Function Collide(A() As Vector, B() As Vector, Number_Of_VerticesA As Long, Number_Of_VerticesB As Long, Offset As Vector, N As Vector, T As Single) As Boolean
    
    Dim Axis(64) As Vector
    Dim TAxis(64) As Single
    
    Dim Number_Of_Axes As Long: Number_Of_Axes = 0
    
    Dim I As Long, J As Long
    
    Dim E0 As Vector
    Dim E1 As Vector
    Dim E As Vector
    
    J = Number_Of_VerticesA - 1
    
    For I = 0 To J
        
        E0 = A(J)
        E1 = A(I)
        
        E = Vector_Subtract(E1, E0)
        
        Axis(Number_Of_Axes).X = -E.Y
        Axis(Number_Of_Axes).Y = E.X
        
        If (Interval_Intersect(A(), B(), Number_Of_VerticesA, Number_Of_VerticesB, Axis(Number_Of_Axes), Offset, TAxis(Number_Of_Axes))) = False Then
        
            Collide = False
            Exit Function
        
        End If
        
        Number_Of_Axes = Number_Of_Axes + 1
        
        J = I
        
    Next I
    
    J = Number_Of_VerticesB - 1
    
    For I = 0 To J
        
        E0 = B(J)
        E1 = B(I)
        
        E = Vector_Subtract(E1, E0)

        Axis(Number_Of_Axes).X = -E.Y
        Axis(Number_Of_Axes).Y = E.X
        
        If (Interval_Intersect(A(), B(), Number_Of_VerticesA, Number_Of_VerticesB, Axis(Number_Of_Axes), Offset, TAxis(Number_Of_Axes))) = False Then
        
            Collide = False
            Exit Function
        
        End If
        
        Number_Of_Axes = Number_Of_Axes + 1
        
        J = I
        
    Next I
    
    If (Find_Minimum_Translation_Distance(Axis(), TAxis(), Number_Of_Axes, N, T)) = False Then
    
        Collide = False
        Exit Function
        
        
    End If
    
    If Vector_Multiply(N, Offset) < 0 Then
    
        N.X = -N.X
        N.Y = -N.Y
        
    End If
    
    Collide = True

End Function

Public Sub Get_Interval(Vertex_List() As Vector, Number_Of_Vertices As Long, Axis As Vector, Min As Single, Max As Single)

    Min = Vector_Multiply(Vertex_List(0), Axis)
    Max = Vector_Multiply(Vertex_List(0), Axis)
    
    Dim I As Long
    
    For I = 1 To Number_Of_Vertices - 1
    
        Dim D As Single: D = Vector_Multiply(Vertex_List(I), Axis)
    
        If (D < Min) Then
        
            Min = D
            
        ElseIf (D > Max) Then
        
            Max = D
            
        End If
    
    Next I

End Sub

Public Function Interval_Intersect(A() As Vector, B() As Vector, Number_Of_VerticesA As Long, Number_Of_VerticesB As Long, Axis As Vector, Offset As Vector, TAxis As Single) As Boolean

    Dim Min(1) As Single, Max(1) As Single
    
    Get_Interval A(), Number_Of_VerticesA, Axis, Min(0), Max(0)
    Get_Interval B(), Number_Of_VerticesB, Axis, Min(1), Max(1)
    
    Dim H As Single: H = Vector_Multiply(Offset, Axis)
    
    Min(0) = Min(0) + H
    Max(0) = Max(0) + H
    
    Dim D0 As Single: D0 = Min(0) - Max(1)
    Dim D1 As Single: D1 = Min(1) - Max(0)
    
    If ((D0 > 0) Or (D1 > 0)) Then
    
        Interval_Intersect = False
        Exit Function
        
    Else

        If D0 > D1 Then
        
            TAxis = D0
            
        Else
        
            TAxis = D1
            
        End If
        
        Interval_Intersect = True
        Exit Function
        
    End If

End Function

Public Function Normalize(Vec As Vector) As Single

    Dim Length As Single: Length = Sqr(Vec.X * Vec.X + Vec.Y * Vec.Y)
        
    If (Length = 0) Then
    
        Normalize = 0
        Exit Function
        
    End If
    
    Vec = Vector_Multiply2(Vec, (1 / Length))

    Normalize = Length
    
End Function

Public Function Find_Minimum_Translation_Distance(Axis() As Vector, TAxis() As Single, Number_Of_Axes As Long, N As Vector, T As Single) As Boolean

    Dim Mini As Long: Mini = -1

    T = 0
    N = Vector_New(0, 0)
    
    Dim I As Long
    
    For I = 0 To Number_Of_Axes - 1
    
        Dim N2 As Single: N2 = Normalize(Axis(I))
        
        TAxis(I) = TAxis(I) / N2
        
        If TAxis(I) > T Or Mini = -1 Then
    
            Mini = I
            T = TAxis(I)
            N = Axis(I)

        End If
        
    Next I
    
    Find_Minimum_Translation_Distance = (Mini <> -1)

End Function

Public Function Collision_Detection() As Boolean
    
    Dim Vertex_List(4) As Vector
    Dim Vertex_List2(4) As Vector
    Dim Boundry(8) As Vector
    Dim Position As Vector
    Dim I As Long
    
    Vertex_List(0) = Vector_New(0, 0)
    Vertex_List(1) = Vector_New(TILE_SIZE, 0)
    Vertex_List(2) = Vector_New(TILE_SIZE, TILE_SIZE)
    Vertex_List(3) = Vector_New(0, TILE_SIZE)
    
    Vertex_List2(0) = Vector_New(0, 0)
    Vertex_List2(1) = Vector_New(TILE_SIZE, 0)
    Vertex_List2(2) = Vector_New(TILE_SIZE, TILE_SIZE)
    Vertex_List2(3) = Vector_New(0, TILE_SIZE)
    
    With Player
        If .Moving Then
            Boundry(0).X = .Coordinates.X - 1: Boundry(0).Y = .Coordinates.Y
            Boundry(1).X = .Coordinates.X:     Boundry(1).Y = .Coordinates.Y - 1
            Boundry(2).X = .Coordinates.X + 1: Boundry(2).Y = .Coordinates.Y
            Boundry(3).X = .Coordinates.X:     Boundry(3).Y = .Coordinates.Y + 1
            Boundry(4).X = .Coordinates.X:     Boundry(4).Y = .Coordinates.Y
            Boundry(5).X = .Coordinates.X - 1:     Boundry(5).Y = .Coordinates.Y - 1
            Boundry(6).X = .Coordinates.X + 1:     Boundry(6).Y = .Coordinates.Y - 1
            Boundry(7).X = .Coordinates.X - 1:      Boundry(7).Y = .Coordinates.Y + 1
            Boundry(8).X = .Coordinates.X + 1:     Boundry(8).Y = .Coordinates.Y + 1
        
            For I = 0 To 8
                If Boundry(I).X <= 0 Then Boundry(I).X = 0
                If Boundry(I).Y <= 0 Then Boundry(I).Y = 0
                If Boundry(I).X >= Map.Width - 1 Then Boundry(I).X = Map.Width - 1
                If Boundry(I).Y >= Map.Height - 1 Then Boundry(I).Y = Map.Height - 1
                
                Position.X = Map.Collision_Map.Vertex_List(Boundry(I).X, Boundry(I).Y).X
                Position.Y = Map.Collision_Map.Vertex_List(Boundry(I).X, Boundry(I).Y).Y
                .Collided = Collide(Vertex_List2(), Vertex_List(), 4, 4, Vector_Subtract(.Position, Position), .NColl, .DColl)
    
                If .Collided = True Then
                    If Map.Collision_Map.Response(Boundry(I).X, Boundry(I).Y) = COLLISION_WALL Then
                        Collision_Detection = True
                        .Position = Vector_Subtract(.Position, Vector_Multiply2(.NColl, .DColl))
                    End If
                End If
            Next I
        End If
    End With
    
    With Monster
    
        Boundry(0).X = .Coordinates.X - 1: Boundry(0).Y = .Coordinates.Y
        Boundry(1).X = .Coordinates.X:     Boundry(1).Y = .Coordinates.Y - 1
        Boundry(2).X = .Coordinates.X + 1: Boundry(2).Y = .Coordinates.Y
        Boundry(3).X = .Coordinates.X:     Boundry(3).Y = .Coordinates.Y + 1
        Boundry(4).X = .Coordinates.X:     Boundry(4).Y = .Coordinates.Y
        Boundry(5).X = .Coordinates.X - 1:     Boundry(5).Y = .Coordinates.Y - 1
        Boundry(6).X = .Coordinates.X + 1:     Boundry(6).Y = .Coordinates.Y - 1
        Boundry(7).X = .Coordinates.X - 1:      Boundry(7).Y = .Coordinates.Y + 1
        Boundry(8).X = .Coordinates.X + 1:     Boundry(8).Y = .Coordinates.Y + 1
    
        For I = 0 To 8
            If Boundry(I).X <= 0 Then Boundry(I).X = 0
            If Boundry(I).Y <= 0 Then Boundry(I).Y = 0
            If Boundry(I).X >= Map.Width - 1 Then Boundry(I).X = Map.Width - 1
            If Boundry(I).Y >= Map.Height - 1 Then Boundry(I).Y = Map.Height - 1
        
            Position.X = Map.Collision_Map.Vertex_List(Boundry(I).X, Boundry(I).Y).X
            Position.Y = Map.Collision_Map.Vertex_List(Boundry(I).X, Boundry(I).Y).Y
            .Collided = Collide(Vertex_List2(), Vertex_List(), 4, 4, Vector_Subtract(.Position, Position), .NColl, .DColl)
            
            If .Collided = True Then
                If Map.Collision_Map.Response(Boundry(I).X, Boundry(I).Y) = COLLISION_WALL Then
                    Collision_Detection = True
                    .Position = Vector_Subtract(.Position, Vector_Multiply2(.NColl, (.DColl * 1.01)))
                End If
            End If
        Next I
        
    End With
    
End Function

Public Function Collision_Detection2(ByVal Overlap As Long)

    Const NO_COLLISION As Long = 0
    Const COL_LEFT As Long = 1
    Const COL_RIGHT As Long = 2
    Const COL_UP As Long = 3
    Const COL_DOWN As Long = 4

    Dim Boundry(8) As Vector
    Dim Side As Long
    Dim I As Long

    With Player
    
        Boundry(0).X = .Coordinates.X - 1: Boundry(0).Y = .Coordinates.Y
        Boundry(1).X = .Coordinates.X:     Boundry(1).Y = .Coordinates.Y - 1
        Boundry(2).X = .Coordinates.X + 1: Boundry(2).Y = .Coordinates.Y
        Boundry(3).X = .Coordinates.X:     Boundry(3).Y = .Coordinates.Y + 1
        Boundry(4).X = .Coordinates.X:     Boundry(4).Y = .Coordinates.Y
        Boundry(5).X = .Coordinates.X - 1:     Boundry(5).Y = .Coordinates.Y - 1
        Boundry(6).X = .Coordinates.X + 1:     Boundry(6).Y = .Coordinates.Y - 1
        Boundry(7).X = .Coordinates.X - 1:      Boundry(7).Y = .Coordinates.Y + 1
        Boundry(8).X = .Coordinates.X + 1:     Boundry(8).Y = .Coordinates.Y + 1
        
        For I = 0 To 8
        
            If Boundry(I).X <= 0 Then Boundry(I).X = 0
            If Boundry(I).Y <= 0 Then Boundry(I).Y = 0
            If Boundry(I).X >= Map.Width - 1 Then Boundry(I).X = Map.Width - 1
            If Boundry(I).Y >= Map.Height - 1 Then Boundry(I).Y = Map.Height - 1

            Side = Collision_Box_To_Box2(Player.Position.X, Player.Position.Y, TILE_SIZE, TILE_SIZE, Boundry(I).X * TILE_SIZE, Boundry(I).Y * TILE_SIZE, TILE_SIZE, TILE_SIZE)
            
            If Side <> 0 Then
                If Map.Collision_Map.Response(Boundry(I).X, Boundry(I).Y) = COLLISION_WALL Then
                    Collision_Detection2 = True
                    .Collided = True
                    Select Case Side
                        Case COL_LEFT: .Position.X = .Position.X + Overlap
                        Case COL_RIGHT:: .Position.X = .Position.X - Overlap
                        Case COL_UP:: .Position.Y = .Position.Y + Overlap
                        Case COL_DOWN:: .Position.Y = .Position.Y - Overlap
                    End Select
                End If
            End If
            
        Next I
        
    End With
    
End Function

    

