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
    Application.OnKey "{F1}", "�񕝒���"
    Application.OnKey "+{F1}", "�s������"
    Application.OnKey "{F3}", "�^�C�g���s�����ݒ�"
    Application.OnKey "+{F4}", "OpenWebSearch"
    Application.OnKey "{F7}", "IndentPlus1"
    Application.OnKey "+{F7}", "IndentMinus1"
    Application.OnKey "{F9}", "PasteText"
End Sub
