Attribute VB_Name = "Main_Mod"
Option Explicit

Public Type MOUSE_EVENT_TYPE

    Position As D3DVECTOR2
    X_Data As Long
    Y_Data As Long
    X_Movement As Boolean
    Y_Movement As Boolean
    Left_Click As Boolean
    Right_Click As Boolean

End Type

Public Const LOW_HZ As Long = 44100 - 2000
Public Const HI_HZ As Long = 44100 + 2000

Public Const VINYL_WEIGHT As Single = 0.33
Public Const VINYL_MASS_KG As Single = 0.1496854821
Public Const VINYL_DIAMETER_CM As Single = 29.2100000297942
Public Const VINYL_RADIUS_CM As Single = 14.6050000148971
Public Const VINYL_MOMENT_OF_INERTIA_KG_CM_2 As Single = (0.5 * VINYL_MASS_KG * VINYL_RADIUS_CM * VINYL_RADIUS_CM) / 100

Public Const STYLUS_DIAMETER_CM As Single = 0.1133980925
Public Const STYLUS_RADIUS_CM As Single = STYLUS_DIAMETER_CM / 2

Public Const PLATTER_WEIGHT As Single = 1.633
Public Const PLATTER_MASS_KG As Single = 0.74071634021
Public Const PLATTER_ACTUAL_RADIUS_CM As Single = 17.1450000148971
Public Const PLATTER_MOMENT_OF_INERTIA_KG_CM_2 As Single = (0.5 * PLATTER_MASS_KG * PLATTER_ACTUAL_RADIUS_CM * PLATTER_ACTUAL_RADIUS_CM) / 100

Public Const VINYL_RADIUS_PIXELS As Single = 150
Public Const NEUTRAL As Long = 0
Public Const CLOCKWISE As Long = 1
Public Const COUNTER_CLOCKWISE As Long = 2
Public Const MOMENTUM_FRICTION As Single = 0.12
Public Const BRAKE_FRICTION As Single = 0.4

Public Time As Single, Time2 As Single
Public Initial_Time As Long, Initial_Time2 As Long
Public Current_Time As Single
Public New_Time As Single
Public Delta_Time As Single, Delta_Time2 As Single
Public Accumulator As Single
Public Fullscreen_Enabled As Boolean
Public Fullscreen_Width As Long
Public Fullscreen_Height As Long
Public Running As Boolean
Public Vertex_List(18) As TLVERTEX
Public Vertex_Buffer As Direct3DVertexBuffer8
Public Texture(1) As Direct3DTexture8
Public Pitch As Single
Public Center_X As Single, Center_Y As Single
Public Scalar As D3DVECTOR2
Public Mouse_Angle As Single
Public Vinyl_Pos As D3DVECTOR2
Public Turntable_Pos As D3DVECTOR2
Public Distance As Single
Public Obj As PHYSICS2D
Public Time_Step As Single
Public Previous As Single
Public Velocity As Single
Public Flag(4) As Boolean
Public Vector_Angle As Single
Public Old_Angle As Single
Public Direction As Long
Public Turntable_Power As Boolean
Public Motor_Power As Boolean
Public Snapshot_Number As Long
Public Str As String
Public Old_Hz As Long
Public Hz As Long

Public Sub DirectInput_Mouse_Callback()

    Dim Mouse_Device_Data(1 To Mouse_Buffer_Size) As DIDEVICEOBJECTDATA
    
    Dim Current_Event As Long
    
    Static Old_Sequence As Long
    
    Dim intBuffer() As Integer, ReadPos As Long
    
    On Error GoTo Error_Handler
    
    Mouse.X_Movement = False
    Mouse.Y_Movement = False
    
    For Current_Event = 1 To Mouse_Device.GetDeviceData(Mouse_Device_Data, 0)
    
        Select Case Mouse_Device_Data(Current_Event).lOfs
        
            Case DIMOFS_X
                
                Mouse.X_Data = Mouse_Device_Data(Current_Event).lData
                
                If Mouse_Device_Data(Current_Event).lData <> 0 Then
                
                    Mouse.X_Movement = True

                End If
                
                If Mouse.Left_Click = False Or Distance > VINYL_RADIUS_PIXELS Then
                
                    Mouse.Position.X = Mouse.Position.X + Mouse_Device_Data(Current_Event).lData
                    
                    Mouse_Angle = 0
                    
                ElseIf Mouse.Left_Click = True And Distance <= VINYL_RADIUS_PIXELS Then
                
                    Dim T As Single, V As Single
                        
                    Mouse_Angle = Get_Radian(Vinyl_Pos.X, Vinyl_Pos.Y, Mouse.Position.X, Mouse.Position.Y)
                    
                    If Modulus(Mouse_Angle, 360) >= Degree_To_Radian(0) And Modulus(Mouse_Angle, 360) <= Degree_To_Radian(180) Then
                    
                        T = -1
                    
                    Else
                    
                        T = 1
                        
                    End If
    
                    V = (VINYL_RADIUS_PIXELS - Abs(Mouse.Position.X - Vinyl_Pos.X)) / VINYL_RADIUS_PIXELS
                
                    Mouse_Angle = Mouse_Angle + Degree_To_Radian((Mouse_Device_Data(Current_Event).lData * T * 2) * (Distance / VINYL_RADIUS_PIXELS)) * V
                    Obj.Angle = Obj.Angle + Degree_To_Radian((Mouse_Device_Data(Current_Event).lData * T * 2) * (Distance / VINYL_RADIUS_PIXELS)) * V
                
                    Mouse.Position.X = Vinyl_Pos.X + Cos(Mouse_Angle) * Distance
                    Mouse.Position.Y = Vinyl_Pos.Y + sIn(Mouse_Angle) * Distance
                    
                End If
                
            Case DIMOFS_Y
                
                Mouse.Y_Data = Mouse_Device_Data(Current_Event).lData
                
                If Mouse_Device_Data(Current_Event).lData <> 0 Then
                
                    Mouse.Y_Movement = True
                    
                End If
                
                If Mouse.Left_Click = False Or Distance > VINYL_RADIUS_PIXELS Then

                    Mouse.Position.Y = Mouse.Position.Y + Mouse_Device_Data(Current_Event).lData
                    
                    Mouse_Angle = 0
                    Flag(1) = False
                    
                ElseIf Mouse.Left_Click = True And Distance <= VINYL_RADIUS_PIXELS Then
            
                    Dim u As Single, W As Single
                        
                    Mouse_Angle = Get_Radian(Vinyl_Pos.X, Vinyl_Pos.Y, Mouse.Position.X, Mouse.Position.Y)
                    
                    If Mouse_Angle >= Degree_To_Radian(90) And Mouse_Angle <= Degree_To_Radian(270) Then
                    
                       u = -1
                    
                    Else
                    
                        u = 1
                        
                    End If
                    
    
                    W = (VINYL_RADIUS_PIXELS - Abs(Mouse.Position.X - Vinyl_Pos.X)) / VINYL_RADIUS_PIXELS
                
                    Mouse_Angle = Mouse_Angle + Degree_To_Radian((Mouse_Device_Data(Current_Event).lData * u * 2) * (Distance / VINYL_RADIUS_PIXELS)) * W
                    Obj.Angle = Obj.Angle + Degree_To_Radian((Mouse_Device_Data(Current_Event).lData * u * 2) * (Distance / VINYL_RADIUS_PIXELS)) * W
                    
                    Mouse.Position.X = Vinyl_Pos.X + Cos(Mouse_Angle) * Distance
                    Mouse.Position.Y = Vinyl_Pos.Y + sIn(Mouse_Angle) * Distance
                    
                    If Flag(1) = False Then
                    
                        Flag(1) = True
                        
                    End If
                
                End If
                
            Case DIMOFS_BUTTON0
            
                If Mouse_Device_Data(Current_Event).lData > 0 Then
                    
                    Mouse.Left_Click = True
                
                Else
                    
                    Mouse.Left_Click = False
                
                End If
            
            Case DIMOFS_BUTTON1
            
                If Mouse_Device_Data(Current_Event).lData > 0 Then
                    
                    Mouse.Right_Click = True
                
                Else
                    
                    Mouse.Right_Click = False
                
                End If
            
            Case DIMOFS_BUTTON2
           
                If Mouse_Device_Data(Current_Event).lData > 0 Then
               
                Else
                
                End If
            
            Case DIMOFS_BUTTON3
            
                If Mouse_Device_Data(Current_Event).lData > 0 Then
                
                Else
                
                End If

        End Select
    
    Next Current_Event
    
    If Mouse.Position.X <= 0 Then Mouse.Position.X = 0
    If Mouse.Position.Y <= 0 Then Mouse.Position.Y = 0

    If Fullscreen_Enabled Then
    
        If Mouse.Position.X >= Fullscreen_Width Then Mouse.Position.X = Fullscreen_Width
        If Mouse.Position.Y >= Fullscreen_Height Then Mouse.Position.Y = Fullscreen_Height
    
    Else
    
        If Mouse.Position.X >= frmMain.ScaleWidth Then Mouse.Position.X = frmMain.ScaleWidth
        If Mouse.Position.Y >= frmMain.ScaleHeight Then Mouse.Position.Y = frmMain.ScaleHeight
    
    End If
    
Error_Handler:

    If Err.Number = DIERR_INPUTLOST Then
    
        Mouse_Device.Unacquire
    
    End If
End Sub

Public Sub Window_Setup(Window As Form, Optional ByVal X As Long = -1, Optional ByVal Y As Long = -1, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional Caption As String = " ", Optional Auto_Redraw As Boolean = False, Optional ByVal Back_Color As Long = -1)
    
    'Use -1 for default values and "" for default strings.
    
    With Window
    
        If Caption <> " " Then .Caption = Caption 'Else use current setting. Note: Some
                                                  'people may want "" as the caption.
        .AutoRedraw = Auto_Redraw
        .ScaleMode = 3
        If X <> -1 Then .Left = X 'Else use current setting.
        If Y <> -1 Then .Top = Y 'Else use current setting.
        If Width <> -1 Then .Width = Width * Screen.TwipsPerPixelX 'Else use current setting.
        If Height <> -1 Then .Height = Height * Screen.TwipsPerPixelY 'Else use current setting.
        If Back_Color <> -1 Then .BackColor = Back_Color 'Else use current setting.
        .Show
        .Refresh
        .SetFocus
        
    End With
    
End Sub

Public Sub Main()

    Fullscreen_Width = 640
    Fullscreen_Height = 480

    Window_Setup frmMain, -1, -1, Fullscreen_Width, Fullscreen_Height, "DJ Program", , RGB(0, 0, 0)
    
    Center_X = frmMain.ScaleWidth / 2
    Center_Y = frmMain.ScaleHeight / 2
    
    DirectX_Initialize
    
    DirectInput_Initialize_Mouse frmMain
    
    Direct3D_Device.SetRenderState D3DRS_ALPHAREF, 255
    Direct3D_Device.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
    
    Create_Polygon
    
    Load_Textures
    
    Pitch = RPM_To_Radians_Per_Second(33.3333333333333)
    
    Setup_Object
    
    Initial_Time = timeGetTime
    
    Previous = Obj.Angle
    
    Time_Step = 1 / 2000
    
    Velocity = Pitch
    
    DirectSound_Load App.Path & "\Beep Ahhh Fresh.wav"
    
    'DirectSound_Play Buffer
    
    Running = True
    
    Game_Loop

End Sub

Public Sub Setup_Object()

    With Obj
        
        .Mass = VINYL_MASS_KG
        .One_Over_Mass = 1 / .Mass
        .Inertia = VINYL_MOMENT_OF_INERTIA_KG_CM_2 + PLATTER_MOMENT_OF_INERTIA_KG_CM_2
        .One_Over_Inertia = 1 / .Inertia
        
    End With

End Sub

Public Function Interpolate(Previous As Single, Current As Single, Alpha As Single) As Single

    Dim State As Single
    
    State = Current * Alpha + Previous * (1 - Alpha)
    
    Interpolate = State
    
End Function

Public Sub Load_Textures()

    Dim File_Path(1) As String
    Dim Width As Long
    Dim Height As Long
    Dim Transparency_Color(1) As Long
    
    File_Path(0) = App.Path & "\vinyl.bmp"
    File_Path(1) = App.Path & "\Turntable.jpg"

    Width = 1024
    Height = 1024
    
    Transparency_Color(0) = D3DColorRGBA(255, 255, 255, 255)
    Transparency_Color(1) = D3DColorRGBA(0, 0, 0, 255)

    Set Texture(0) = Direct3DX.CreateTextureFromFileEx(Direct3D_Device, _
                                                    File_Path(0), _
                                                    Width, Height, _
                                                    0, _
                                                    0, _
                                                    D3DFMT_A8R8G8B8, _
                                                    D3DPOOL_MANAGED, _
                                                    D3DX_FILTER_POINT, _
                                                    D3DX_FILTER_POINT, _
                                                    Transparency_Color(0), _
                                                    ByVal 0, _
                                                    ByVal 0)
                                                    
    Set Texture(1) = Direct3DX.CreateTextureFromFileEx(Direct3D_Device, _
                                                    File_Path(1), _
                                                    Width, Height, _
                                                    0, _
                                                    0, _
                                                    D3DFMT_A8R8G8B8, _
                                                    D3DPOOL_MANAGED, _
                                                    D3DX_FILTER_POINT, _
                                                    D3DX_FILTER_POINT, _
                                                    Transparency_Color(1), _
                                                    ByVal 0, _
                                                    ByVal 0)

End Sub

Private Sub Create_Polygon()
    
    Dim Color As Long
    
    Dim X As Single, Y As Single
    
    Dim Width As Single, Height As Single
    
    Color = D3DColorRGBA(255, 255, 255, 0)
    
    Width = VINYL_RADIUS_PIXELS * Scalar.X
    Height = VINYL_RADIUS_PIXELS * Scalar.Y
    
    Vertex_List(0) = Create_TLVertex(-Width, -Height, 0, 1, Color, 0, 0, 0)
    Vertex_List(1) = Create_TLVertex(Width, -Height, 0, 1, Color, 0, 1, 0)
    Vertex_List(2) = Create_TLVertex(-Width, Height, 0, 1, Color, 0, 0, 1)
    Vertex_List(3) = Create_TLVertex(Width, -Height, 0, 1, Color, 0, 1, 0)
    Vertex_List(4) = Create_TLVertex(Width, Height, 0, 1, Color, 0, 1, 1)
    Vertex_List(5) = Create_TLVertex(-Width, Height, 0, 1, Color, 0, 0, 1)
    
    Color = D3DColorRGBA(255, 255, 255, 255)
    
    Width = 175 * Scalar.X
    Height = 224 * Scalar.Y
    
    Vertex_List(6) = Create_TLVertex(-Width, -Height, 0, 1, Color, 0, 0, 0)
    Vertex_List(7) = Create_TLVertex(Width, -Height, 0, 1, Color, 0, 1, 0)
    Vertex_List(8) = Create_TLVertex(-Width, Height, 0, 1, Color, 0, 0, 1)
    Vertex_List(9) = Create_TLVertex(Width, -Height, 0, 1, Color, 0, 1, 0)
    Vertex_List(10) = Create_TLVertex(Width, Height, 0, 1, Color, 0, 1, 1)
    Vertex_List(11) = Create_TLVertex(-Width, Height, 0, 1, Color, 0, 0, 1)
    
    Width = 5 * Scalar.X
    Height = 5 * Scalar.Y
    
    Vertex_List(12) = Create_TLVertex(-Width, -Height, 0, 1, Color, 0, 0, 0)
    Vertex_List(13) = Create_TLVertex(Width, -Height, 0, 1, Color, 0, 1, 0)
    Vertex_List(14) = Create_TLVertex(-Width, Height, 0, 1, Color, 0, 0, 1)
    Vertex_List(15) = Create_TLVertex(Width, -Height, 0, 1, Color, 0, 1, 0)
    Vertex_List(16) = Create_TLVertex(Width, Height, 0, 1, Color, 0, 1, 1)
    Vertex_List(17) = Create_TLVertex(-Width, Height, 0, 1, Color, 0, 0, 1)
    
    Vinyl_Pos.X = frmMain.ScaleWidth / 2 + -2
    Vinyl_Pos.Y = frmMain.ScaleHeight / 2 + 47
    
    Turntable_Pos.X = frmMain.ScaleWidth / 2
    Turntable_Pos.Y = frmMain.ScaleHeight / 2

    Set Vertex_Buffer = Direct3D_Device.CreateVertexBuffer(Len(Vertex_List(0)) * UBound(Vertex_List), 0, FVF_TLVERTEX, D3DPOOL_MANAGED)
    D3DVertexBuffer8SetData Vertex_Buffer, 0, Len(Vertex_List(0)) * UBound(Vertex_List), 0, Vertex_List(0)
    Direct3D_Device.SetStreamSource 0, Vertex_Buffer, Len(Vertex_List(0))

End Sub

Public Sub Rotate(ByVal Radian As Single, Starting_Vertex As Long, Ending_Vertex As Single)

    Dim New_Vertex() As TLVERTEX
    Dim I As Long
    
    ReDim New_Vertex(UBound(Vertex_List)) As TLVERTEX
    
    For I = Starting_Vertex To Ending_Vertex

        New_Vertex(I).X = ((Vertex_List(I).X * Cos(Radian) - Vertex_List(I).Y * sIn(Radian)))
        New_Vertex(I).Y = ((Vertex_List(I).X * sIn(Radian) + Vertex_List(I).Y * Cos(Radian)))
        
        Vertex_List(I).X = New_Vertex(I).X
        Vertex_List(I).Y = New_Vertex(I).Y
        
    Next I
    
    Set Vertex_Buffer = Direct3D_Device.CreateVertexBuffer(Len(Vertex_List(0)) * UBound(Vertex_List), 0, FVF_TLVERTEX, D3DPOOL_MANAGED)
    D3DVertexBuffer8SetData Vertex_Buffer, 0, Len(Vertex_List(0)) * UBound(Vertex_List), 0, Vertex_List(0)
    Direct3D_Device.SetStreamSource 0, Vertex_Buffer, Len(Vertex_List(0))
    
End Sub

Public Sub Translate(ByVal X As Single, ByVal Y As Single, Starting_Vertex As Long, Ending_Vertex As Single)

    Dim New_Vertex() As TLVERTEX
    Dim I As Long
    
    ReDim New_Vertex(UBound(Vertex_List)) As TLVERTEX
    
    For I = Starting_Vertex To Ending_Vertex
    
        New_Vertex(I).X = X + Vertex_List(I).X
        New_Vertex(I).Y = Y + Vertex_List(I).Y
        
        Vertex_List(I).X = New_Vertex(I).X
        Vertex_List(I).Y = New_Vertex(I).Y
        
    Next I
    
    Set Vertex_Buffer = Direct3D_Device.CreateVertexBuffer(Len(Vertex_List(0)) * UBound(Vertex_List), 0, FVF_TLVERTEX, D3DPOOL_MANAGED)
    D3DVertexBuffer8SetData Vertex_Buffer, 0, Len(Vertex_List(0)) * UBound(Vertex_List), 0, Vertex_List(0)
    Direct3D_Device.SetStreamSource 0, Vertex_Buffer, Len(Vertex_List(0))
    
End Sub

Public Sub Render()

    Direct3D_Device.BeginScene
    
        Create_Polygon
        
        'Turntables
        
        Rotate 0, 6, 11
        
        Translate Turntable_Pos.X, Turntable_Pos.Y, 6, 11
    
        Direct3D_Device.SetRenderState D3DRS_ALPHATESTENABLE, False
        
        Direct3D_Device.SetTexture 0, Texture(1)
        
        Direct3D_Device.DrawPrimitive D3DPT_TRIANGLELIST, 6, 2
        
        'Vinyls
        
        Rotate Obj.Angle, 0, 5
        
        Translate Vinyl_Pos.X, Vinyl_Pos.Y, 0, 5
    
        Direct3D_Device.SetRenderState D3DRS_ALPHATESTENABLE, True
        
        Direct3D_Device.SetTexture 0, Texture(0)
    
        Direct3D_Device.DrawPrimitive D3DPT_TRIANGLELIST, 0, 2
        
        'Mouse Pointer
        
        Rotate 0, 12, 17
        
        Translate Mouse.Position.X, Mouse.Position.Y, 12, 17
    
        Direct3D_Device.SetRenderState D3DRS_ALPHATESTENABLE, False
        
        Direct3D_Device.SetTexture 0, Nothing
        
        Direct3D_Device.DrawPrimitive D3DPT_TRIANGLELIST, 12, 2
        
        DirectX8_Initialize_Font "Arial", 14
        
        Str = "Turntable Power: " & Turntable_Power & "     " & "Motor Power: " & _
              Motor_Power & vbCrLf & Format(Hz, "###,##0.00") & "Hz"
        
        DirectX8_Draw_Text Str, 40, 40, 500, 100, &HFF000000 + RGB(255, 255, 255)
    
    Direct3D_Device.EndScene

End Sub

Public Sub Prevent_LockUp()

    Mouse_Device.GetDeviceStateMouse Mouse_State
    
    If Mouse_State.Buttons(0) = 0 Then Mouse.Left_Click = False

End Sub

Public Sub Turntable_Physics()
    
    Calculate_Torque_2Da Obj, (Vector_Angle * 180 / PI), (Distance / VINYL_RADIUS_PIXELS) * VINYL_RADIUS_CM
    
    If Mouse.Left_Click = False Then
        
        If Turntable_Power = True And Motor_Power = True Then
        
            Obj.Torque.Net = ONE_KGF_CM_TO_NEWTONS_CM
            
            If Direction = CLOCKWISE Then
            
                If Obj.Angular_Velocity > Pitch Then
                
                    Obj.Torque.Net = -ONE_KGF_CM_TO_NEWTONS_CM
                    
                End If
            
            Else
           
                If Obj.Angular_Velocity < Pitch Then
                
                    Obj.Torque.Net = ONE_KGF_CM_TO_NEWTONS_CM
                    
                End If
    
            End If
            
        ElseIf Turntable_Power = True And Motor_Power = False Then
            
            If Flag(4) = False Then
           
                If Direction = CLOCKWISE Then
                
                    Obj.Torque.Net = -Calculate_Frictional_Force(MOMENTUM_FRICTION, Calculate_Normal_Force(VINYL_MOMENT_OF_INERTIA_KG_CM_2 + PLATTER_MOMENT_OF_INERTIA_KG_CM_2, EARTH_GRAVITY))
                
                ElseIf Direction = COUNTER_CLOCKWISE Then
                
                    Obj.Torque.Net = Calculate_Frictional_Force(MOMENTUM_FRICTION, Calculate_Normal_Force(VINYL_MOMENT_OF_INERTIA_KG_CM_2 + PLATTER_MOMENT_OF_INERTIA_KG_CM_2, EARTH_GRAVITY))
                
                End If
                
            Else
            
                If Direction = CLOCKWISE Then
                
                    Obj.Torque.Net = -Calculate_Frictional_Force(BRAKE_FRICTION, Calculate_Normal_Force(Obj.Inertia, EARTH_GRAVITY))
                
                ElseIf Direction = COUNTER_CLOCKWISE Then
                
                    Obj.Torque.Net = Calculate_Frictional_Force(BRAKE_FRICTION, Calculate_Normal_Force(Obj.Inertia, EARTH_GRAVITY))
                
                End If
            
            End If
            
        ElseIf Turntable_Power = False And Motor_Power = False Then
            
            If Direction = CLOCKWISE Then
            
                Obj.Torque.Net = -Calculate_Frictional_Force(MOMENTUM_FRICTION, Calculate_Normal_Force(Obj.Inertia, EARTH_GRAVITY))
            
            ElseIf Direction = COUNTER_CLOCKWISE Then
            
                Obj.Torque.Net = Calculate_Frictional_Force(MOMENTUM_FRICTION, Calculate_Normal_Force(Obj.Inertia, EARTH_GRAVITY))
            
            End If
            
        ElseIf Turntable_Power = False And Motor_Power = True Then
            
            Motor_Power = False
            
        End If
        
    ElseIf Mouse.Left_Click = True And Distance <= VINYL_RADIUS_PIXELS Then
    
        Obj.Angular_Velocity = 0
        
    ElseIf Mouse.Left_Click = True And Distance > VINYL_RADIUS_PIXELS Then
        
        Obj.Torque.Net = 0
        
    End If
    
    Const Velocity_Resistance As Long = 15
    
    If Obj.Angular_Velocity >= Velocity_Resistance Then Obj.Angular_Velocity = Velocity_Resistance
    If Obj.Angular_Velocity <= -Velocity_Resistance Then Obj.Angular_Velocity = -Velocity_Resistance
    
    If Direction = CLOCKWISE Then
    
        If Obj.Angular_Velocity < 0 Then
                    
           Obj.Torque.Net = 0
           Obj.Angular_Velocity = 0
    
        End If
        
    ElseIf Direction = COUNTER_CLOCKWISE Then
       
        If Obj.Angular_Velocity > 0 Then
    
            Obj.Torque.Net = 0
            Obj.Angular_Velocity = 0
                        
        End If
        
    End If
    
    'Timestep

        New_Time = timeGetTime
        Delta_Time = (New_Time - Current_Time) / 1000
        Delta_Time2 = Delta_Time
        Current_Time = New_Time
        
        If Delta_Time > 0.25 Then Delta_Time = 0.25
        
        Accumulator = Accumulator + Delta_Time
        
        While (Accumulator >= Time_Step)
        
            'DoEvents
            
            Accumulator = Accumulator - Time_Step
             
            Integrate2D Obj, Time_Step, FORTH_ORDER_RUNGE_KUTTA
            
            Time = Time + Time_Step
            
        Wend

    Interpolate Previous, Obj.Angle, Accumulator / Time_Step
    
    Vector_Angle = Obj.Angle - Old_Angle
    
    If Vector_Angle > 0 Then
        
        Direction = CLOCKWISE
            
    ElseIf Vector_Angle < 0 Then
        
        Direction = COUNTER_CLOCKWISE
        
    Else
    
        Direction = NEUTRAL
            
    End If
    
    Old_Angle = Obj.Angle

End Sub

Public Sub Game_Loop()
    
    Do While Running = True
        
        DoEvents
        
        Dim Vector As D3DVECTOR2
        
        Vector.X = Mouse.Position.X - Vinyl_Pos.X
        Vector.Y = Mouse.Position.Y - Vinyl_Pos.Y
    
        Distance = Sqr(Vector.X * Vector.X + Vector.Y * Vector.Y)
        
        Turntable_Physics
        
        Prevent_LockUp
        
        If Not Buffer Is Nothing Then
                
            If Direction = CLOCKWISE Then
            
                Reverse_Sound = False
            
                If Obj.Angular_Velocity <= 0 Then
                
                    DirectSound_Pause Buffer
                    
                Else
                
                    DirectSound_Play Buffer
            
                End If
            
            ElseIf Direction = COUNTER_CLOCKWISE Then
            
                Reverse_Sound = True
            
                If Obj.Angular_Velocity >= 0 Then
                
                    DirectSound_Pause Buffer
                    
                Else
                
                    DirectSound_Play Buffer
            
                End If
                
            Else
            
                DirectSound_Pause Buffer
            
            End If
                
            Hz = Abs(44100 * (Obj.Angular_Velocity / Pitch))
                
            If Obj.Angular_Velocity > 0 And Hz >= LOW_HZ And Hz <= HI_HZ And Motor_Power = True Then
                
                Obj.Angular_Velocity = RPM_To_Radians_Per_Second(33.3333333333333)
                Hz = 44100
                
            End If
                
            If Hz >= 176400 Then Hz = 176400
                
            Old_Hz = Hz
            
            DirectSound_Set_Speed Buffer, Hz
            
        End If
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
            
            Render
        
        Direct3D_Device.Present ByVal 0, ByVal 0, 0, ByVal 0
        
    Loop

End Sub

Public Sub Close_Program()

    Running = False 'This helps the program bail out of the game loop.
    
    'Unload all of the DirectX objects.
    
    Set Texture(0) = Nothing
    Set Direct3DX = Nothing
    Set Direct3D_Device = Nothing
    Set Direct3D = Nothing
    Set DirectX8 = Nothing
    
    DirectSound_Pause Buffer

    If Wave_File_Number <> 0 Then
    
        Close Wave_File_Number
        Wave_File_Number = 0
        
    End If
    
    DirectSound_Uninitialize
    
    Unload frmMain 'Unload the form.
    
    End 'Ends the program.
    
    'Although the Unload statement located above exits the program, you
    'will end up with an Automation error after doing so. The End statement
    'will help prevent that, and end the app completely.

End Sub
