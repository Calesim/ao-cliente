Attribute VB_Name = "modAStar"
Option Explicit

Public Const Opened As Long = 1
Public Const Closed As Long = 2

Public mu As Vector

Public Sub Clear_AStar(Map As Map_Type)

    Dim Current As Vector
    
    Reset_Heap Monster
    'If IsArrayInitialized(VarPtrArray(Sprite.Nodes)) = True Then
        For Current.Y = 0 To Map.Height - 1
            For Current.X = 0 To Map.Width - 1
                With Monster.Nodes(Current.X, Current.Y)
                    .F = 0
                    .G = 0
                    .H = 0
                    .OCList = 0
                    .X = 0
                    .Y = 0
                End With
            Next Current.X
        Next Current.Y
    'End If

End Sub

    'Reset the heap
Public Sub Reset_Heap(Sprite As Sprite_Type)

    Sprite.Size_Of_Heap = 0
    ReDim Sprite.Heap(0)
    
End Sub

'Remove the Root Object from the heap
Public Sub Remove_Root(Sprite As Sprite_Type)
    
    Dim Parent As Long
    Dim Child_Index As Long
    
    'If only the root exists
    If Sprite.Size_Of_Heap <= 1 Then
        Sprite.Size_Of_Heap = 0
        ReDim Sprite.Heap(0)
        Exit Sub
    End If

    'First copy the very bottom object to the top
    Sprite.Heap(1) = Sprite.Heap(Sprite.Size_Of_Heap)

    'Resize the count
    Sprite.Size_Of_Heap = Sprite.Size_Of_Heap - 1

    'Shrink the array
    ReDim Preserve Sprite.Heap(Sprite.Size_Of_Heap)

    'Sort the top item to it's correct position
    Parent = 1
    Child_Index = 1

    'Sink the item to it's correct location
    Do While True
        Child_Index = Parent
        If 2 * Child_Index + 1 <= Sprite.Size_Of_Heap Then
            'Find the lowest value of the 2 child nodes
            If Sprite.Heap(Child_Index).Score >= Sprite.Heap(2 * Child_Index).Score Then Parent = 2 * Child_Index
            If Sprite.Heap(Parent).Score >= Sprite.Heap(2 * Child_Index + 1).Score Then Parent = 2 * Child_Index + 1
        Else 'Just process the one node
            If 2 * Child_Index <= Sprite.Size_Of_Heap Then
                If Sprite.Heap(Child_Index).Score >= Sprite.Heap(2 * Child_Index).Score Then Parent = 2 * Child_Index
            End If
        End If

        'Swap out the child/parent
        If Parent <> Child_Index Then
            Dim Temp_Heap As Heap_Type
            Temp_Heap = Sprite.Heap(Child_Index)
            Sprite.Heap(Child_Index) = Sprite.Heap(Parent)
            Sprite.Heap(Parent) = Temp_Heap
        Else
            Exit Do
        End If

    Loop

End Sub

'Add the new element to the heap
Public Sub Add(Sprite As Sprite_Type, ByVal Score As Long, ByVal X As Long, ByVal Y As Long)
    
    Dim Position As Long
    
    '**We will be ignoring the (0) place in the heap array because
    '**it's easier to handle the heap with a base of (1..?)

    'Increment the array count
    Sprite.Size_Of_Heap = Sprite.Size_Of_Heap + 1

    'Make room in the array
    ReDim Preserve Sprite.Heap(Sprite.Size_Of_Heap)

    'Store the data
    With Sprite.Heap(Sprite.Size_Of_Heap)
        .Score = Score
        .X = X
        .Y = Y
    End With

    'Bubble the item to its correct location
    
    Position = Sprite.Size_Of_Heap

    Do While Position <> 1
        If Sprite.Heap(Position).Score <= Sprite.Heap(Position / 2).Score Then
            Dim Temp_Heap As Heap_Type
            Temp_Heap = Sprite.Heap(Position / 2)
            Sprite.Heap(Position / 2) = Sprite.Heap(Position)
            Sprite.Heap(Position) = Temp_Heap
            Position = Position / 2
        Else
            Exit Do
        End If
    Loop

End Sub

Public Function AStar_Find_Path(Map As Map_Type, Predator As Sprite_Type, Prey As Sprite_Type) As Boolean
    
    Dim Parent As Vector
    Dim Current As Vector
    Dim tempCost As Long
    Dim Walkable As Boolean
    Dim Temp As Vector, Temp2 As Vector
    Dim Current_Node As Long

    If Predator.Compute_AStar_Enabled = True Or Prey.Compute_AStar_Enabled = True Then
    
        If Prey.Coordinates.X < 0 Or Prey.Coordinates.Y < 0 Or Prey.Coordinates.X > (Map.Width - 1) Or Prey.Coordinates.Y > (Map.Height - 1) Then Exit Function
        If Predator.Coordinates.X < 0 Or Predator.Coordinates.Y < 0 Or Predator.Coordinates.X > (Map.Width - 1) Or Predator.Coordinates.Y > (Map.Height - 1) Then Exit Function
        
        'Make sure the starting point and ending point are not the same
        If (Predator.Coordinates.X = Prey.Coordinates.X) And (Predator.Coordinates.Y = Prey.Coordinates.Y) Then Exit Function
        
        If Map.Collision_Map.Response(Predator.Coordinates.X, Predator.Coordinates.Y) = COLLISION_WALL Then Exit Function
        If Map.Collision_Map.Response(Prey.Coordinates.X, Prey.Coordinates.Y) = COLLISION_WALL Then Exit Function
    
        'Set the flags
        Predator.Path_Found = False
        Predator.Path_Hunt = True
    
        'Put the starting point on the open list
        Predator.Nodes(Predator.Coordinates.X, Predator.Coordinates.Y).OCList = Opened
        Add Predator, 0, Predator.Coordinates.X, Predator.Coordinates.Y
    
        'Find the children
        Do While Predator.Path_Hunt
            If Predator.Size_Of_Heap <> 0 Then
                'Get the parent node
                Parent.X = Predator.Heap(1).X
                Parent.Y = Predator.Heap(1).Y
    
                'Remove the root
                Predator.Nodes(Parent.X, Parent.Y).OCList = Closed
                Remove_Root Predator
    
                'Find the available children to add to the open list
                For Current.Y = (Parent.Y - 1) To (Parent.Y + 1)
                    For Current.X = (Parent.X - 1) To (Parent.X + 1)
    
                        'Make sure we are not out of bounds
                        If Current.X >= 0 And Current.X <= Map.Width - 1 And Current.Y >= 0 And Current.Y <= Map.Height - 1 Then
    
                            'Make sure it's not on the closed list
                            If Predator.Nodes(Current.X, Current.Y).OCList <> Closed Then
    
                                'Make sure no wall
                                If Map.Collision_Map.Response(Current.X, Current.Y) = COLLISION_NONE Then
    
                                    'Don't cut across corners
                                    Walkable = True
                                    
                                    If Current.X = Parent.X - 1 Then
                                        If Current.Y = Parent.Y - 1 Then
                                            If Map.Collision_Map.Response(Parent.X - 1, Parent.Y) = COLLISION_WALL Or Map.Collision_Map.Response(Parent.X, Parent.Y - 1) = COLLISION_WALL Then Walkable = False
                                        ElseIf Current.Y = Parent.Y + 1 Then
                                            If Map.Collision_Map.Response(Parent.X, Parent.Y + 1) = COLLISION_WALL Or Map.Collision_Map.Response(Parent.X - 1, Parent.Y) = COLLISION_WALL Then Walkable = False
                                        End If
                                    ElseIf Current.X = Parent.X + 1 Then
                                        If Current.Y = Parent.Y - 1 Then
                                            If Map.Collision_Map.Response(Parent.X, Parent.Y - 1) = COLLISION_WALL Or Map.Collision_Map.Response(Parent.X + 1, Parent.Y) = COLLISION_WALL Then Walkable = False
                                        ElseIf Current.Y = Parent.Y + 1 Then
                                            If Map.Collision_Map.Response(Parent.X + 1, Parent.Y) = COLLISION_WALL Or Map.Collision_Map.Response(Parent.X, Parent.Y + 1) = COLLISION_WALL Then Walkable = False
                                        End If
                                    End If
    
                                    'If we can move this way
                                    If Walkable = True Then
                                        If Predator.Nodes(Current.X, Current.Y).OCList <> Opened Then
    
                                            'Calculate the G
                                            If Math.Abs(Current.X - Parent.X) = 1 And Math.Abs(Current.Y - Parent.Y) = 1 Then
                                                Predator.Nodes(Current.X, Current.Y).G = Predator.Nodes(Parent.X, Parent.Y).G + 14
                                            Else
                                                Predator.Nodes(Current.X, Current.Y).G = Predator.Nodes(Parent.X, Parent.Y).G + 10
                                            End If
    
                                            'Calculate the H
                                            Predator.Nodes(Current.X, Current.Y).H = 10 * (Math.Abs(Current.X - Prey.Coordinates.X) + Math.Abs(Current.Y - Prey.Coordinates.Y))
                                            Predator.Nodes(Current.X, Current.Y).F = (Predator.Nodes(Current.X, Current.Y).G + Predator.Nodes(Current.X, Current.Y).H)
    
                                            'Add the parent value
                                            Predator.Nodes(Current.X, Current.Y).X = Parent.X
                                            Predator.Nodes(Current.X, Current.Y).Y = Parent.Y
    
                                            'Add the item to the heap
                                            Add Predator, Predator.Nodes(Current.X, Current.Y).F, Current.X, Current.Y
    
                                            'Add the item to the open list
                                            Predator.Nodes(Current.X, Current.Y).OCList = Opened
    
                                        Else
                                            'We will check for better value
                                            Dim AddedG As Long
                                            If Math.Abs(Current.X - Parent.X) = COLLISION_WALL And Math.Abs(Current.Y - Parent.Y) = COLLISION_WALL Then
                                                AddedG = 14
                                            Else
                                                AddedG = 10
                                            End If
                                            
                                            tempCost = Predator.Nodes(Parent.X, Parent.Y).G + AddedG
    
                                            If tempCost < Predator.Nodes(Current.X, Current.Y).G Then
                                                Predator.Nodes(Current.X, Current.Y).G = tempCost
                                                Predator.Nodes(Current.X, Current.Y).X = Parent.X
                                                Predator.Nodes(Current.X, Current.Y).Y = Parent.Y
                                                If Predator.Nodes(Current.X, Current.Y).OCList = Opened Then
                                                    Dim NewCost As Long: NewCost = Predator.Nodes(Current.X, Current.Y).H + Predator.Nodes(Current.X, Current.Y).G
                                                    Add Predator, NewCost, Current.X, Current.Y
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next Current.X
                Next Current.Y
            Else
                Predator.Path_Found = False
                Predator.Path_Hunt = False
                'MsgBox "Path Not Found", vbExclamation
                'Instead of a message box, you could have the  run back where it originated instead or not move at all.
                Exit Function
            End If
    
            'If we find a path
            If Predator.Nodes(Prey.Coordinates.X, Prey.Coordinates.Y).OCList = Opened Then
                Predator.Path_Found = True
                Predator.Path_Hunt = False
                'MsgBox "path found"
            End If
    
        Loop
        Dim V As Vector
        If Predator.Path_Found Then
            Temp.X = Prey.Coordinates.X
            Temp.Y = Prey.Coordinates.Y
            Do While True
                ReDim Preserve Predator.AStar_Path(Current_Node) As Vector
                Temp2.X = Predator.Nodes(Temp.X, Temp.Y).X
                Temp2.Y = Predator.Nodes(Temp.X, Temp.Y).Y
                Predator.AStar_Path(Current_Node).X = Temp.X
                Predator.AStar_Path(Current_Node).Y = Temp.Y
                If Temp.X = Predator.Coordinates.X And Temp.Y = Predator.Coordinates.Y Then Exit Do
                Current_Node = Current_Node + 1
                Temp.X = Temp2.X
                Temp.Y = Temp2.Y
            Loop
            Predator.Length_Of_AStar_Path = Current_Node
            Predator.Current_Path = Current_Node
        End If
    End If

End Function

Public Sub Follow_AStar_Path(Map As Map_Type, Predator As Sprite_Type, ByVal Speed As Single)
        
    Dim AStar_Position As Vector
    Static Temp_AStar_Path As Vector
    Dim Delta As Vector
    Dim Distance As Single
    Dim Move As Vector
    Dim Ratio As Single
    Dim New_Position As Vector
    Dim Angle As Single
    Dim Velocity As Vector
    Dim Vec As Vector
    
    If Predator.Path_Found = True And Predator.Path_Hunt = False Then
        If IsArrayInitialized(VarPtrArray(Predator.AStar_Path)) = True Then
            If Predator.Current_Path >= 0 Then
                If Predator.AStar_Moving = False Then
                    Temp_AStar_Path.X = Predator.AStar_Path(Predator.Current_Path).X
                    Temp_AStar_Path.Y = Predator.AStar_Path(Predator.Current_Path).Y
                End If
                
                AStar_Position.X = (Temp_AStar_Path.X * TILE_SIZE)
                AStar_Position.Y = (Temp_AStar_Path.Y * TILE_SIZE)
                
                Delta.X = AStar_Position.X - Predator.Position.X
                Delta.Y = AStar_Position.Y - Predator.Position.Y
                
                Distance = Sqr(Delta.X * Delta.X + Delta.Y * Delta.Y)
                
                If Distance > Speed Then
                    Ratio = Speed / Distance
                    Move.X = Ratio * Delta.X
                    Move.Y = Ratio * Delta.Y
                    Predator.Position.X = Predator.Position.X + Move.X
                    Predator.Position.Y = Predator.Position.Y + Move.Y
                    Predator.AStar_Moving = True
                Else
                    Predator.Position.X = AStar_Position.X
                    Predator.Position.Y = AStar_Position.Y
                    Predator.Current_Path = Predator.Current_Path - 1
                    If Predator.Current_Path <= 0 Then Predator.Current_Path = 0
                    Predator.AStar_Moving = False
                End If
            End If
        End If
    Else
        Clear_AStar Map
    End If

End Sub

Public Sub Draw_AStar_Path(Sprite As Sprite_Type)
    
    Dim Current_Node As Long
    Dim Position As Vector
    
    If IsArrayInitialized(VarPtrArray(Sprite.AStar_Path)) = False Then
        Exit Sub
    End If
    If Sprite.Length_Of_AStar_Path <= 0 Then
        Exit Sub
    End If
    If Sprite.Path_Found = True And Sprite.Path_Hunt = False Then
        For Current_Node = Sprite.Length_Of_AStar_Path To 0 Step -1
            Position.X = (Sprite.AStar_Path(Current_Node).X * TILE_SIZE) + (TILE_SIZE / 2)
            Position.Y = (Sprite.AStar_Path(Current_Node).Y * TILE_SIZE) + (TILE_SIZE / 2)
            Draw_Pixel Position.X, Position.Y, RGB(255, 255, 0)
        Next Current_Node
    End If

End Sub


