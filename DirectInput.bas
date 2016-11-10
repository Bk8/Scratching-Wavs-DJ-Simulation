Attribute VB_Name = "DirectInput"
Option Explicit

Public Const Mouse_Buffer_Size As Long = 10

Public Direct_Input As DirectInput8

Public Mouse_Device As DirectInputDevice8
Public Mouse_State As DIMOUSESTATE
Public Mouse_Properties As DIPROPLONG
Public Mouse_Event_Handle As Long
Public Mouse As MOUSE_EVENT_TYPE
Public Mouse_Cursor_Position As Point_API

Public Sub DirectInput_Initialize_Mouse(Window As Form)
    
    'Initialize mouse input.
    
    Set Direct_Input = DirectX8.DirectInputCreate
    Set Mouse_Device = Direct_Input.CreateDevice("GUID_SYSMOUSE")
    Mouse_Device.SetCommonDataFormat DIFORMAT_MOUSE
    Mouse_Device.SetCooperativeLevel Window.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
    Mouse_Properties.lHow = DIPH_DEVICE
    Mouse_Properties.lObj = 0
    Mouse_Properties.lData = Mouse_Buffer_Size
    Mouse_Device.SetProperty "DIPROP_BUFFERSIZE", Mouse_Properties
    Mouse_Event_Handle = DirectX8.CreateEvent(Window)
    Mouse_Device.SetEventNotification Mouse_Event_Handle
    Mouse_Device.Acquire
    
    'Set cursor in the middle of the screen.
    
    GetCursorPos Mouse_Cursor_Position
    ScreenToClient Window.hWnd, Mouse_Cursor_Position
    Mouse.Position.X = CSng(Mouse_Cursor_Position.X)
    Mouse.Position.Y = CSng(Mouse_Cursor_Position.Y)
    
End Sub
