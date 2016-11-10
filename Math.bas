Attribute VB_Name = "Math"
Option Explicit

Public Const PI As Single = 3.14159265358979

Public Function Get_Radian(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single

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

Public Function Get_Degree(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single

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

Public Function Check_Triangle(A As TLVERTEX, B As TLVERTEX, C As TLVERTEX, ByVal X As Single, ByVal Y As Single) As Boolean

    If Determinant(A.X, A.Y, B.X, B.Y, X, Y) = True And Determinant(B.X, B.Y, C.X, C.Y, X, Y) = True And Determinant(C.X, C.Y, A.X, A.Y, X, Y) = True Then
        
        Check_Triangle = True
        
    Else
        
        Check_Triangle = False
        
    End If
    
End Function

Public Function Determinant(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single) As Boolean

    Dim determ As Single
    
    determ = (X2 - X1) * (Y3 - Y1) - (X3 - X1) * (Y2 - Y1)

    If determ >= 0 Then
    
        Determinant = True
        
    Else
    
        Determinant = False
        
    End If
    
End Function

Public Function Modulus(ByVal A As Single, ByVal B As Single) As Single

    Modulus = A - Int(A / B) * B
    
End Function

Public Function Degree_To_Radian(Angle As Single) As Single

    Degree_To_Radian = Angle * PI / 180
    
End Function

Public Function Radian_To_Degree(Angle As Single) As Single

    Radian_To_Degree = Angle * 180 / PI
    
End Function

Public Function RPM_To_Degrees_Per_Second(ByVal RPM As Single) As Single
    
    RPM_To_Degrees_Per_Second = (360 * RPM) / 60

End Function

Public Function RPM_To_Radians_Per_Second(ByVal RPM As Single) As Single
    
    RPM_To_Radians_Per_Second = (6.28318531 * RPM) / 60

End Function

Public Sub Swap(ByRef V1 As Integer, ByRef V2 As Integer)

    Dim Temp As Integer
    
    Temp = V1
    V1 = V2
    V2 = Temp

End Sub
