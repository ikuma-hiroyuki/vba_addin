Attribute VB_Name = "Tools"
Option Explicit

Sub WindowSize()
    With Application
        .Height = 650: .Width = 1155
        .Left = 220: .Top = 104
    End With
End Sub

Private Sub 列幅調整()
    Dim FirstRow As Long: FirstRow = Selection.Row
    Dim Rg As Range
    For Each Rg In Selection
        If Rg.Row <> FirstRow Then Exit For
        Rg.EntireColumn.AutoFit
    Next
End Sub
Private Sub 行幅調整()
    Dim Rg As Range
    For Each Rg In Selection
        Rg.EntireRow.AutoFit
    Next
End Sub

Private Sub タイトル行書式設定()
    Dim Rg As Range
    For Each Rg In Selection
        With Rg
            '.HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(31, 111, 67)
        End With
        With Rg.Font
            .FontStyle = "太字"
            .ThemeColor = xlThemeColorDark1
        End With
    Next
End Sub

Private Sub セルの幅高さ変更()
    With Selection
        .ColumnWidth = Application.CentimetersToPoints(3) * .ColumnWidth / .Width
        .RowHeight = Application.CentimetersToPoints(3)
    End With
End Sub

Sub IndentPlus1()
    Selection.InsertIndent 1
End Sub

Sub IndentMinus1()
    Selection.InsertIndent -1
End Sub

Sub PasteText()
    ActiveSheet.PasteSpecial _
        Format:="HTML", _
        Link:=False, _
        DisplayAsIcon:=False, _
        NoHTMLFormatting:=True
End Sub
