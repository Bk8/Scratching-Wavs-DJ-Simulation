Attribute VB_Name = "DirectXGraphics"
Option Explicit

Public Main_Font As D3DXFont
Public Main_Font_Description As IFont
Public Text_Rect As RECT
Public Font As New StdFont

Public Sub DirectX8_Initialize_Font(Font_Name As String, ByVal Font_Size As Long)

    Font.Name = Font_Name
    Font.Size = Font_Size
    Set Main_Font_Description = Font
    Set Main_Font = Direct3DX.CreateFont(Direct3D_Device, Main_Font_Description.hFont)

End Sub

Public Sub DirectX8_Draw_Text(Text As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long)
    
    Text_Rect.Left = X
    Text_Rect.Top = Y
    Text_Rect.Right = X + Width
    Text_Rect.bottom = Y + Height
    Direct3DX.DrawText Main_Font, Color, Text, Text_Rect, DT_TOP Or DT_LEFT

End Sub
