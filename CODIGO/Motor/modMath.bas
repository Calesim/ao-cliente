Attribute VB_Name = "modMath"
Option Explicit

Public Const PI As Single = 3.14159265358979 'Atn(1) * 4

Public Type Vector

    X As Single
    Y As Single
    
End Type

Public Function ACos(ByVal Value As Double) As Double

    On Error GoTo Error_Handler
    ACos = Atn(-Value / Sqr(-Value * Value + 1)) + 2 * Atn(1)
    Exit Function
Error_Handler:
    ACos = 0
    
End Function

Public Function ASin(ByVal Value As Double) As Double

    On Error GoTo Error_Handler
    ASin = Atn(Value / Sqr(-Value * Value + 1))
    Exit Function
Error_Handler:
    ASin = 0

End Function

Public Function Ceil(ByVal Number) As Long

    If Number >= 0 Then
        If Number = Int(Number) Then
            Ceil = Number
        Else
            Ceil = Int(Number) + 1
        End If
    ElseIf Number < 0 Then
        Ceil = Int(Number)
        
    End If
End Function

Public Function Floor(ByVal Number) As Long

    Floor = Fix(Number)
    
End Function

Public Function Linear_Interpolation_1D(ByVal Vertex As Single, ByVal X_Start As Single, ByVal X_End As Single, ByRef mu As Single, ByVal Speed) As Boolean

    If (X_End >= X_Start) Then
        mu = mu + Convert_Speed_To_MU(Speed, X_Start, X_End)
    Else
        mu = mu + Convert_Speed_To_MU(-Speed, X_Start, X_End)
    End If
    
    If mu <= 0 Then mu = 0
    If mu >= 1 Then
        mu = 1
        Linear_Interpolation_1D = True
    End If
    
   Vertex = (X_Start * (1 - mu) + X_End * mu)

End Function

Public Function Linear_Interpolation_2D(ByRef Position As Vector, ByVal X_Start As Single, ByVal Y_Start As Single, ByVal X_End As Single, ByVal Y_End As Single, ByRef mu As Vector, ByVal Speed As Single) As Boolean
    
    Dim Radian As Single
    
    If (X_End - X_Start) = 0 And (Y_End - Y_Start) <> 0 Then
        mu.X = 1
    ElseIf (X_End - X_Start) <> 0 And (Y_End - Y_Start) = 0 Then
        mu.Y = 1
    ElseIf (X_End - X_Start) = 0 And (Y_End - Y_Start) = 0 Then
        Linear_Interpolation_2D = True
        Exit Function
    End If
    
    Radian = Get_Radian(X_Start, Y_Start, X_End, Y_End)
    
    mu.X = mu.X + Convert_Speed_To_MU(Speed, X_Start, X_End) * Cos(Radian)
    mu.Y = mu.Y + Convert_Speed_To_MU(Speed, Y_Start, Y_End) * Sin(Radian)

    If mu.X <= 0 Then mu.X = 0
    If mu.X >= 1 Then mu.X = 1

    If mu.Y <= 0 Then mu.Y = 0
    If mu.Y >= 1 Then mu.Y = 1
    
    Position.X = Ceil(X_Start * (1 - mu.X) + X_End * mu.X)
    Position.Y = Ceil(Y_Start * (1 - mu.Y) + Y_End * mu.Y)
    
    If mu.X = 1 And mu.Y = 1 Then
        Linear_Interpolation_2D = True
        Exit Function
    End If
    
End Function

Public Function Convert_Speed_To_MU(ByVal Speed As Single, ByVal X_Start As Single, ByVal X_End As Single) As Single
    
    If (X_End - X_Start) <> 0 Then Convert_Speed_To_MU = Speed / (X_End - X_Start)

End Function

Public Function Vector_New(ByVal X As Single, ByVal Y As Single) As Vector

    Vector_New.X = X
    Vector_New.Y = Y
    
End Function

Public Function Vector_Subtract(ByRef A As Vector, ByRef B As Vector) As Vector

    Vector_Subtract.X = A.X - B.X
    Vector_Subtract.Y = A.Y - B.Y
    
End Function

Public Function Vector_Multiply(ByRef A As Vector, ByRef B As Vector) As Single

    Vector_Multiply = A.X * B.X + A.Y * B.Y

End Function

Public Function Vector_Multiply2(ByRef A As Vector, ByVal Value As Single) As Vector

    Vector_Multiply2 = Vector_New(A.X * Value, A.Y * Value)

End Function

Public Function Vector_Dot_Product(ByRef A As Vector, ByRef B As Vector) As Single

    Vector_Dot_Product = (A.X * B.X) + (A.Y * B.Y)

End Function

Public Function Get_Radian(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single

    Dim DX As Single, DY As Single
    Dim Angle As Single

        DX = X2 - X1
        DY = Y2 - Y1
        
        Angle = 0

        If DX = 0 Then
            If DY = 0 Then
                Angle = 0
            ElseIf DY > 0 Then
                Angle = PI / 2
            Else
                Angle = PI * 3 / 2
            End If
        ElseIf DY = 0 Then
            If DX > 0 Then
                Angle = 0
            Else
                Angle = PI
            End If
        Else
            If DX < 0 Then
                Angle = Atn(DY / DX) + PI
            ElseIf DY < 0 Then
                Angle = Atn(DY / DX) + (2 * PI)
            Else
                Angle = Atn(DY / DX)
            End If
        End If
        Get_Radian = Angle

End Function

Public Function Get_Degree(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single

    Dim DX As Single, DY As Single
    Dim Angle As Single

        DX = X2 - X1
        DY = Y2 - Y1
        Angle = 0

        If DX = 0 Then
            If DY = 0 Then
                Angle = 0
            ElseIf DY > 0 Then
                Angle = PI / 2
            Else
                Angle = PI * 3 / 2
            End If
        ElseIf DY = 0 Then
            If DX > 0 Then
                Angle = 0
            Else
                Angle = PI
            End If
        Else
            If DX < 0 Then
                Angle = Atn(DY / DX) + PI
            ElseIf DY < 0 Then
                Angle = Atn(DY / DX) + (2 * PI)
            Else
                Angle = Atn(DY / DX)
            End If
        End If
        Angle = Angle * PI / 180
        Get_Degree = Angle

End Function

Public Function Angle_Between(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Single
    
    On Error GoTo Error_Handler
    
    Dim Dot As Single
    Dim V1 As Vector, V2 As Vector
    Dim Theta As Single
    Dim V1_Mag As Single, V2_Mag As Single
    
    V1.X = X1: V1.Y = Y1
    V2.X = X2: V2.Y = Y2
    
    Dot = Vector_Dot_Product(V1, V2)
    V1_Mag = CSng(Sqr((X1 * X1) + (Y1 * Y1)))
    V2_Mag = CSng(Sqr((X2 * X2) + (Y2 * Y2)))
    Theta = CSng(ACos(Dot / (V1_Mag * V2_Mag)))
    
    Angle_Between = Theta
    Exit Function
Error_Handler:
    Angle_Between = 0
    
'static public float angleBetween(PVector v1, PVector v2) {
'float dot = v1.dot(v2);
'float theta = (float) Math.acos(dot / (v1.mag() * v2.mag()));
'return theta;

End Function

Public Function Degree_To_Radian(ByVal Angle As Single) As Single

    Degree_To_Radian = Angle * PI / 180
    
End Function

Public Function Radian_To_Degree(ByVal Angle As Single) As Single

    Radian_To_Degree = Angle * 180 / PI
    
End Function

