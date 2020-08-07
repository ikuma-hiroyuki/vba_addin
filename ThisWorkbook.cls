VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Application.OnKey "{F1}", "列幅調整"
    Application.OnKey "+{F1}", "行幅調整"
    Application.OnKey "{F3}", "タイトル行書式設定"
    Application.OnKey "+{F4}", "OpenWebSearch"
    Application.OnKey "{F7}", "IndentPlus1"
    Application.OnKey "+{F7}", "IndentMinus1"
    Application.OnKey "{F9}", "PasteText"
End Sub

