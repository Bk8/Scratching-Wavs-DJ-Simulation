Attribute VB_Name = "DirectSound"
Option Explicit

Public Enum IfStringNotFound

    ReturnOriginalStr = 0
    ReturnEmptyStr = 1
    
End Enum

Public Type FileHeader

    lRiff As Long
    lFileSize As Long
    lWave As Long
    lFormat As Long
    lFormatLength As Long
    
End Type

Public Type WaveFormat

    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    
End Type

Public Type ChunkHeader

    lType As Long
    lLen As Long
    
End Type

Public Sound_Enum As DirectSoundEnum8
Public DirectSound8 As DirectSound8

Public Buffer As DirectSoundSecondaryBuffer8
Public Buffer_Wave_Format As WAVEFORMATEX
Public Buffer_Description As DSBUFFERDESC

Public Wave_Format As WAVEFORMATEX
Public Wave_File_Number As Long
Public lDataLength As Long
Public lDataStartPos As Long

Public BuffLen As Long, HalfBuffLen As Long

Public EventsNotify() As DSBPOSITIONNOTIFY
Public EndEvent As Long, MidEvent As Long, StartEvent As Long, LastEvent As Long

Public Sound_Driver_List() As String

Public Reverse_Sound As Boolean

Public Sound_Flag(1) As Boolean

Public intBuffer() As Integer, ReadPos As Long

Public Function DirectSound_Initialize(Window As Form, Optional ByVal SamplesPerSec As Long = 44100, _
                            Optional ByVal BitsPerSample As Integer = 16, _
                            Optional ByVal Channels As Integer = 2, Optional GUID As String) As String
    On Error GoTo ReturnError
    
    Set Sound_Enum = DirectX8.GetDSEnum
    
    If Len(GUID) > 0 Then
        Set DirectSound8 = DirectX8.DirectSoundCreate(GUID)
    Else
        Set DirectSound8 = DirectX8.DirectSoundCreate(Sound_Enum.GetGuid(1))
    End If
    
    Buffer_Wave_Format.nFormatTag = WAVE_FORMAT_PCM
    Buffer_Wave_Format.nChannels = Channels
    Buffer_Wave_Format.nBitsPerSample = BitsPerSample
    Buffer_Wave_Format.lSamplesPerSec = SamplesPerSec
    
    Buffer_Wave_Format.nBlockAlign = (Buffer_Wave_Format.nBitsPerSample * Buffer_Wave_Format.nChannels) \ 8
    Buffer_Wave_Format.lAvgBytesPerSec = Buffer_Wave_Format.lSamplesPerSec * Buffer_Wave_Format.nBlockAlign
    
    HalfBuffLen = Buffer_Wave_Format.lAvgBytesPerSec / 15 ' this number determines the length of the buffer
    HalfBuffLen = HalfBuffLen + (HalfBuffLen Mod Buffer_Wave_Format.nBlockAlign)
    
    BuffLen = HalfBuffLen * 2
    
    Buffer_Description.fxFormat = Buffer_Wave_Format
    Buffer_Description.lBufferBytes = BuffLen
    Buffer_Description.lFlags = DSBCAPS_CTRLPOSITIONNOTIFY Or DSBCAPS_STICKYFOCUS Or DSBCAPS_CTRLFREQUENCY
    
    DirectSound8.SetCooperativeLevel Window.hWnd, DSSCL_NORMAL
    Set Buffer = DirectSound8.CreateSoundBuffer(Buffer_Description)
    
    ReDim EventsNotify(0 To 2) As DSBPOSITIONNOTIFY
    
    StartEvent = DirectX8.CreateEvent(Window)
    EventsNotify(0).hEventNotify = StartEvent
    EventsNotify(0).lOffset = 1
    
    MidEvent = DirectX8.CreateEvent(Window)
    EventsNotify(1).hEventNotify = MidEvent
    EventsNotify(1).lOffset = HalfBuffLen
    
    EndEvent = DirectX8.CreateEvent(Window)
    EventsNotify(2).hEventNotify = EndEvent
    EventsNotify(2).lOffset = DSBPN_OFFSETSTOP
    
    Buffer.SetNotificationPositions 3, EventsNotify()
    
    DirectSound_Initialize = ""
    Debug.Print "Direct Sound Initialized"
    Exit Function
ReturnError:
    DirectSound_Initialize = "Error: " & Err.Number & vbNewLine & _
        "Desription: " & Err.Description & vbNewLine & _
        "Source: " & Err.Source
    Debug.Print DirectSound_Initialize
    
    Err.Clear
    DirectSound_Uninitialize
    Exit Function
End Function

Public Sub DirectSound_Get_Sound_Driver_List(Sound_Driver_List() As String)
    
    Dim Current_Sound_Driver As Long
    
    Dim Available As Long
    
    For Current_Sound_Driver = 1 To DirectX8.GetDSEnum.GetCount
    
        If DirectX8.GetDSEnum.GetGuid(Current_Sound_Driver) <> "{00000000-0000-0000-0000-000000000000}" Then
            
            ReDim Sound_Driver_List(Available) As String
            
            Sound_Driver_List(Available) = DirectX8.GetDSEnum.GetDescription(Current_Sound_Driver)
            
            Available = Available + 1
            
        End If
        
    Next Current_Sound_Driver

End Sub

Public Function DirectSound_Play(Buffer As DirectSoundSecondaryBuffer8) As Boolean

    On Error GoTo Error_Handler
    
    If Sound_Flag(0) = False Then
        
        If Not Buffer Is Nothing Then Buffer.Play DSBPLAY_LOOPING
    
        Sound_Flag(0) = True
        Sound_Flag(1) = False
    
        DirectSound_Play = True
        
    End If
    
    Exit Function
    
Error_Handler:

    DirectSound_Play = False

End Function

Public Sub DirectSound_Set_Position(Wave_File_Number As Long, New_Position As Long)

    If Not Buffer Is Nothing Then
    
        Seek Wave_File_Number, lDataStartPos + New_Position
        
    End If

End Sub

Public Function DirectSound_Get_Length(Wave_File_Number As Long) As Long

    DirectSound_Get_Length = LOF(Wave_File_Number)

End Function

Public Function DirectSound_Get_Current_Position(Wave_File_Number As Long) As Long

    DirectSound_Get_Current_Position = Loc(Wave_File_Number)

End Function

Public Function DirectSound_Pause(Buffer As DirectSoundSecondaryBuffer8) As Boolean

    On Error GoTo Error_Handler
        
    If Sound_Flag(1) = False Then
        
        If Not Buffer Is Nothing Then Buffer.Stop
        
        Sound_Flag(0) = False
        Sound_Flag(1) = True
        
        DirectSound_Pause = True
        
    End If
        
    Exit Function
    
Error_Handler:

    DirectSound_Pause = False
    
End Function

Public Function DirectSound_Stop(Buffer As DirectSoundSecondaryBuffer8) As Boolean

    On Error GoTo Error_Handler

    If Sound_Flag(1) = False Then

        If Not Buffer Is Nothing Then Buffer.Stop
        
        DirectSound_Stop = True
        
        Seek Wave_File_Number, lDataStartPos
        
        Sound_Flag(0) = False
        Sound_Flag(1) = True
        
    End If
        
    Exit Function
    
Error_Handler:

    DirectSound_Stop = False
    
End Function

Public Sub DirectSound_Set_Speed(Buffer As DirectSoundSecondaryBuffer8, ByVal Hz As Long)
    
    On Error Resume Next
    
    Buffer.SetFrequency Hz

End Sub

Public Function DirectSound_Wave_Read_Format(ByVal InFileNum As Integer) As WAVEFORMATEX

    Dim Header As FileHeader
    Dim HdrFormat As WaveFormat
    Dim chunk As ChunkHeader
    Dim by As Byte
    Dim I As Long
    
    Get #InFileNum, 1, Header
    
    If Header.lRiff <> &H46464952 Then Exit Function   ' Check for "RIFF" tag and exit if not found.
    If Header.lWave <> &H45564157 Then Exit Function   ' Check for "WAVE" tag and exit if not found.
    If Header.lFormat <> &H20746D66 Then Exit Function ' Check for "fmt " tag and exit if not found.
    
    ' Check format chunk length; if less than 16, it's not PCM data so we can't use it.
    If Header.lFormatLength < 16 Then Exit Function
    
    Get #InFileNum, , HdrFormat ' Retrieve format.
    
    ' Seek to next chunk by discarding any format bytes.
    For I = 1 To Header.lFormatLength - 16
        Get #InFileNum, , by
    Next
    
    ' Ignore chunks until we get to the "data" chunk.
    Get #InFileNum, , chunk
    Do While chunk.lType <> &H61746164
        For I = 1 To chunk.lLen
            Get #InFileNum, , by
        Next
        Get #InFileNum, , chunk
    Loop
    
    lDataLength = chunk.lLen ' Retrieve the size of the data.
    lDataStartPos = Loc(InFileNum) + 1
    
    ' Fill the returned type with the format information.
    With DirectSound_Wave_Read_Format
        .lAvgBytesPerSec = HdrFormat.nAvgBytesPerSec
        .lSamplesPerSec = HdrFormat.nSamplesPerSec
        .nBitsPerSample = HdrFormat.wBitsPerSample
        .nBlockAlign = HdrFormat.nBlockAlign
        .nChannels = HdrFormat.nChannels
        .nFormatTag = HdrFormat.wFormatTag
        .nSize = 0
    End With
    
End Function

Public Sub DirectSound_Uninitialize()

    On Error Resume Next
    
    DirectSound_Pause Buffer
    
    DoEvents
    
    HalfBuffLen = 0
    
    DirectX8.DestroyEvent EventsNotify(0).hEventNotify
    DirectX8.DestroyEvent EventsNotify(1).hEventNotify
    DirectX8.DestroyEvent EventsNotify(2).hEventNotify
    
    Erase EventsNotify
    
    Set Buffer = Nothing
    Set DirectSound8 = Nothing
    Set Sound_Enum = Nothing
    
End Sub

Public Sub DirectSound_Load(File_Path As String)

    Dim StrRet As String
    
    On Error GoTo Error_Handler
    
    If Wave_File_Number <> 0 Then Close Wave_File_Number
    
    Wave_File_Number = FreeFile
    
    Open File_Path For Binary As Wave_File_Number
    
    Wave_Format = DirectSound_Wave_Read_Format(Wave_File_Number)
    
    DirectSound_Uninitialize
    
    StrRet = DirectSound_Initialize(frmMain, Wave_Format.lSamplesPerSec, Wave_Format.nBitsPerSample, Wave_Format.nChannels, DirectX8.GetDSEnum.GetGuid(2))
    
    If StrRet <> "" Then
    
        MsgBox StrRet
        Close Wave_File_Number
        Wave_File_Number = 0
        
    End If
    
    DoEvents
    
    Exit Sub
    
Error_Handler:

    Exit Sub
    
End Sub

Public Sub DirectSound_Reverse_Buffer(Buffer() As Integer)

    Dim K As Long, HHBuff As Long
    
    HHBuff = UBound(Buffer) \ 2
    
    For K = 0 To HHBuff
    
        Swap Buffer(K), Buffer(UBound(Buffer) - K)
        
    Next K
    
End Sub

Private Sub DirectSound_Control(ByVal EventID As Long)
    
    On Error Resume Next
    
    If Wave_File_Number = 0 Then Exit Sub
    
    LastEvent = EventID
    
    If Reverse_Sound = False Then ' Forward Play
    
        If Loc(Wave_File_Number) >= LOF(Wave_File_Number) Then
        
            DirectSound_Stop Buffer
            'Exit Sub
            
        End If
        
        If Loc(Wave_File_Number) + HalfBuffLen > LOF(Wave_File_Number) Then
        
            ReDim intBuffer((LOF(Wave_File_Number) - Loc(Wave_File_Number)) \ 2) ' read only remaining bytes
            
        Else
        
            ReDim intBuffer(HalfBuffLen \ 2 - 1)
            
        End If
        
        Get Wave_File_Number, , intBuffer
        
    Else
    
        ' Reverse Play
        
        If Loc(Wave_File_Number) <= lDataStartPos Then
            
            DirectSound_Set_Position 1, DirectSound_Get_Length(1)
            
        End If
        
        If Loc(Wave_File_Number) - HalfBuffLen < lDataStartPos Then
        
            ReDim intBuffer((Loc(Wave_File_Number) - lDataStartPos) \ 2)
            
        Else
        
            ReDim intBuffer(HalfBuffLen \ 2 - 1)
            
        End If
        
        ReadPos = Loc(Wave_File_Number) - (UBound(intBuffer) * 2 + 1)
        
        Get Wave_File_Number, , intBuffer
        Seek Wave_File_Number, ReadPos
        
        DirectSound_Reverse_Buffer intBuffer
        
    End If

End Sub

Public Sub DirectSound_Callback(ByVal EventID As Long)
    
    Select Case EventID
    
        Case StartEvent
        
            DirectSound_Control EventID
        
            Buffer.WriteBuffer HalfBuffLen, HalfBuffLen, intBuffer(0), DSCBLOCK_DEFAULT
            
        Case MidEvent
        
            DirectSound_Control EventID
        
            Buffer.WriteBuffer 0, HalfBuffLen, intBuffer(0), DSCBLOCK_DEFAULT
        
    End Select
    
End Sub


