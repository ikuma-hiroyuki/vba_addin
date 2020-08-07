Attribute VB_Name = "OneTime"
Option Explicit

Sub SFE()
    Dim s As Booster:    Set s = New Booster
    s.TemporarySpeedUp
    Dim Rg As Range, i As Long: i = 1
    For Each Rg In Selection
        Rg = i
        i = i + 1
    Next Rg
End Sub

