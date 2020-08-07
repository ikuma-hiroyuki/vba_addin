Attribute VB_Name = "Coding"
Option Explicit

Sub FNR() 'For Next 作成(最終行)
    Dim Str As String
    Str = vbTab & "Dim LastRow As Long: LastRow = Cells(Rows.Count, 1).End(xlUp).Row" & vbNewLine & _
          vbTab & "dim r as long" & vbNewLine & _
          vbTab & "for r = 2 to LastRow" & _
          vbNewLine & vbNewLine & _
          vbTab & "next r"
    Debug.Print Str
End Sub

Sub FNC() 'For Next 作成(最終列)
    Dim Str As String
    Str = vbTab & "Dim LastCol As Long: LastCol = Cells(1, Columns.Count).End(xltoleft).Column" & vbNewLine & _
          vbTab & "dim c as long" & vbNewLine & _
          vbTab & "for c = 2 to LastCol" & _
          vbNewLine & vbNewLine & _
          vbTab & "next c"
    Debug.Print Str
End Sub

Sub TE() 'タイトル行Enum化
    Debug.Print "private enum"
    Dim Rg As Range
    For Each Rg In Selection
        Debug.Print vbTab & Rg & "=" & Rg.Column
    Next
    Debug.Print "end enum"
End Sub

Sub SRP(FindStr As String, ReplaceStr As String) '選択範囲リプレイス
    Dim Rg As Range
    For Each Rg In Selection
        Rg = Replace(Rg, FindStr, ReplaceStr)
    Next Rg
End Sub

Sub SFE() '選択範囲For Each
    Debug.Print _
        "sub SFE" & vbNewLine & vbTab & _
            "Dim Rg As Range" & vbNewLine & vbTab & _
            "For Each Rg In Selection" & vbNewLine & vbTab & _
                 vbNewLine & vbTab & _
            "Next Rg" & vbNewLine & _
        "end sub"
End Sub
