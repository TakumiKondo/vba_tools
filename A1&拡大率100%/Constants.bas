Attribute VB_Name = "Constants"
Public Function fontName()
    fontName = ThisWorkbook.Worksheets("Sheet1").Range("FONT_NAME").Value    '// �t�H���g��
End Function

Public Function targetDir()
    targetDir = ThisWorkbook.Worksheets("Sheet1").Range("targetDir").Value
End Function

Public Function dialogTitle()
    dialogTitle = "�t�H���_���̃G�N�Z���t�@�C����o��ԉ�"
End Function
