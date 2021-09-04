Attribute VB_Name = "Constants"
Public Function fontName()
    fontName = ThisWorkbook.Worksheets("Sheet1").Range("FONT_NAME").Value    '// フォント名
End Function

Public Function targetDir()
    targetDir = ThisWorkbook.Worksheets("Sheet1").Range("targetDir").Value
End Function

Public Function dialogTitle()
    dialogTitle = "フォルダ内のエクセルファイル提出状態化"
End Function
