VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Booster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("HiAddIns")
Option Explicit

Enum EventOnOff
    EventOn = -1
    EventOff = 0
End Enum
    
Public Sub TemporarySpeedUp(Optional OnOff As EventOnOff = EventOn)
    With Application
        .EnableEvents = OnOff
        .ScreenUpdating = False
        .DisplayStatusBar = False
        .PrintCommunication = False
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub Class_Terminate()
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayStatusBar = True
        .PrintCommunication = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub
