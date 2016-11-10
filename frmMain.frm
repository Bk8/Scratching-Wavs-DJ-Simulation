VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements DirectXEvent8

Private Sub DirectXEvent8_DXCallback(ByVal EventID As Long)

    DirectInput_Mouse_Callback
    DirectSound_Callback EventID

End Sub

Private Sub Form_Activate()

    Main

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then 'If the user presses the Esc key...
    
        Close_Program
    
    End If

    If KeyCode = vbKeyReturn Then
        
        If Flag(3) = False Then
             
            Flag(3) = True
            
            Turntable_Power = (Turntable_Power + 1) Mod 2
                
        End If
        
    End If

    If KeyCode = vbKeySpace Then
        
        If Turntable_Power = True Then
        
            If Flag(2) = False Then
                 
                Flag(2) = True
                
                Motor_Power = (Motor_Power + 1) Mod 2
                
                If Motor_Power = False Then
                
                    Flag(4) = True
                    
                Else
                
                    Flag(4) = False
                    
                End If
                    
            End If
            
        Else
        
            Flag(4) = False
        
        End If
        
    End If


    'If KeyCode = vbKeyF12 Then
    
    '    If Dir$(App.Path & "\Snapshots\", vbDirectory) = "" Then
            
    '        MkDir App.Path & "\Snapshots\"
            
    '    End If
        
        'If Dir$(App.Path & "\Snapshots\SNAP" & Format(Snapshot_Number, "####") & ".bmp") = "" Then
        
        '    DirectX_Snapshot App.Path & "\Snapshots\SNAP" & Format(Snapshot_Number, "####") & ".bmp"
        
        'End If
        
        'While Dir$(App.Path & "\Snapshots\SNAP" & Format(Snapshot_Number, "####") & ".bmp") <> ""
            
        '    DoEvents
            
        '    Snapshot_Number = Snapshot_Number + 1
            
        'Wend
    
    'End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Flag(2) = True Then
        
        Flag(2) = False
        
    End If
    
    If Flag(3) = True Then
        
        Flag(3) = False
        
    End If

End Sub

Private Sub Form_Load()

    'This event will fire before the form has completely loaded.
    
    MsgBox "Instructions" & vbCrLf & "--------------------------------" & vbCrLf & _
           "Return - Turntable Power On/Off" & vbCrLf & _
           "Spacebar - Motor Power On/Off (Only turns on if turntable is on)" & vbCrLf & _
           vbCrLf & "Note: Both powers must be on to get the vinyl to spin automatically." & vbCrLf & vbCrLf & _
           "Use the mouse to spin the vinyl as well.", vbInformation
           

    If MsgBox("Click Yes to go to full screen (Recommended)", vbQuestion Or vbYesNo, "Options") = vbYes Then Fullscreen_Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Close_Program

End Sub

Private Sub Timer1_Timer()

End Sub
