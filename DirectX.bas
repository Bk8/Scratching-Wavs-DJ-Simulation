Attribute VB_Name = "DirectX"
Option Explicit

Public Type TLVERTEX

    X As Single
    Y As Single
    Z As Single
    RHW As Single
    Color As Long
    Specular As Long
    TU As Single
    TV As Single
    
End Type

Public Const COLOR_DEPTH_16_BIT As Long = D3DFMT_R5G6B5
Public Const COLOR_DEPTH_24_BIT As Long = D3DFMT_A8R8G8B8
Public Const COLOR_DEPTH_32_BIT As Long = D3DFMT_X8R8G8B8

Public Const FVF_TLVERTEX As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Public DirectX8 As DirectX8
Public Direct3D As Direct3D8
Public Direct3D_Device As Direct3DDevice8
Public Direct3DX As D3DX8

Public Function DirectX_Initialize() As Boolean

    On Error GoTo Error_Handler
    
    Dim Display_Mode As D3DDISPLAYMODE
    Dim Direct3D_Window As D3DPRESENT_PARAMETERS
    
    Set DirectX8 = New DirectX8
    Set Direct3D = DirectX8.Direct3DCreate()
    Set Direct3DX = New D3DX8
    
    If Fullscreen_Enabled = True Then
    
        Display_Mode.Width = Fullscreen_Width
        Display_Mode.Height = Fullscreen_Height
        Display_Mode.Format = COLOR_DEPTH_16_BIT
    
        Direct3D_Window.Windowed = False
        Direct3D_Window.BackBufferCount = 1
        Direct3D_Window.BackBufferWidth = Display_Mode.Width
        Direct3D_Window.BackBufferHeight = Display_Mode.Height 'Match the backbuffer height with the display height
        Direct3D_Window.hDeviceWindow = frmMain.hWnd 'Use frmMain as the device window.
        
        Scalar.X = Display_Mode.Width / frmMain.ScaleWidth
        Scalar.Y = Display_Mode.Height / frmMain.ScaleHeight * 0.95
        
    Else
    
        Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                        'are already on. Incase you are confused, I'm
                                                                        'talking about your current screen resolution. ;)
        
        Direct3D_Window.Windowed = True 'The app will be in windowed mode.
    
        Scalar.X = 1
        Scalar.Y = 1
    
    End If
    
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    
    'Creates the rendering device with some useful info, along with the info
    'we've already setup for Direct3D_Window.
    Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Direct3D_Window)
    
    Direct3D_Device.SetVertexShader FVF_TLVERTEX 'Set the type of vertex shading. (Required)
    
    'These lines are not needed, but it's nice to be able to filter the
    'textures to make them look nicer.
    
    Direct3D_Device.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
    Direct3D_Device.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
    
    Exit Function
    
Error_Handler:
    
    MsgBox "An error occured while initializing DirectX", vbCritical
    
    Close_Program
    
    DirectX_Initialize = False

End Function

Public Function Create_TLVertex(X As Single, Y As Single, Z As Single, RHW As Single, Color As Long, Specular As Long, TU As Single, TV As Single) As TLVERTEX

    Create_TLVertex.X = X
    Create_TLVertex.Y = Y
    Create_TLVertex.Z = Z
    Create_TLVertex.RHW = RHW
    Create_TLVertex.Color = Color
    Create_TLVertex.Specular = Specular
    Create_TLVertex.TU = TU
    Create_TLVertex.TV = TV
    
End Function

Public Sub DirectX_Snapshot(ByVal File_Path As String)

    Dim Surface As Direct3DSurface8
    Dim SrcPalette As PALETTEENTRY
    Dim SrcRect As RECT
    Dim Direct3D_Display_Mode As D3DDISPLAYMODE

    'get display dimensions
    Direct3D_Device.GetDisplayMode Direct3D_Display_Mode

    'create a surface to put front buffer on,
    'GetFrontBuffer always returns D3DFMT_A8R8G8B8
    Set Surface = Direct3D_Device.CreateImageSurface(Direct3D_Display_Mode.Width, Direct3D_Display_Mode.Height, D3DFMT_A8R8G8B8)

    'get data from front buffer
    Direct3D_Device.GetFrontBuffer Surface

    'we are saving entire area of this surface
    
    With SrcRect
    
        .Left = 0
        .Right = Direct3D_Display_Mode.Width
        .Top = 0
        .bottom = Direct3D_Display_Mode.Height
        
    End With

    'save this surface to a BMP file
    Direct3DX.SaveSurfaceToFile File_Path, D3DXIFF_BMP, Surface, SrcPalette, SrcRect

End Sub
