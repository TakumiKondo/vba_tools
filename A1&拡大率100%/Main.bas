Attribute VB_Name = "Main"
Sub Main()
    Dim Path As String: Path = ThisWorkbook.Worksheets("Sheet1").Range("targetDir").Value   '// �Ώۃt�H���_�p�X
    Const cnsTitle = "�t�H���_���̃G�N�Z���t�@�C����o��ԉ�" '�_�C�A���O�̃^�C�g��
    
    Dim Workbook As Workbook
    Dim sheet As Object
    
    ' �t�H���_�̑��݊m�F
    If Dir(Path, vbDirectory) = "" Then
        MsgBox "�w��̃t�H���_�͑��݂��܂���B", vbExclamation, cnsTitle
        Exit Sub
    End If
    
    
    Call beforeExecution
    '// A1���g�嗦100%��ݒ肷��
    Call setA1And100Per(Path)
    Call afterExecution
    
    MsgBox "�������܂����B"

End Sub

'// A1���g�嗦100%��ݒ肷��
Private Sub setA1And100Per(Path)
    '// A1���g�嗦100����ݒ肷��i���[�g�t�H���_�j
    Call executeA1And100Per(Path)
    
    '// �T�u�t�H���_���ċA����
    If ActiveSheet.CheckBoxes("CheckBox").Value = 1 Then
        recursion (Path)
    End If

End Sub


'// Excel�n�t�@�C���̂݁AA1���g�嗦100���ɂ��āA�ŏ��̃V�[�g���w�肵�ĕۑ�����
Private Sub executeA1And100Per(Path)
    Const cnsDIR = "\*.*"
    Dim strFileName As String '�������̃t�@�C�������i�[����ϐ�
    Dim fileAndPath As String '�������̃t�@�C�����i�p�X�܂ށj���i�[����ϐ�
    Dim pos As Long

    ' �擪�̃t�@�C�����̎擾
    strFileName = Dir(Path & cnsDIR, vbNormal)
    ' �t�@�C����������Ȃ��Ȃ�܂ŌJ��Ԃ�
    Do While strFileName <> ""
    
        ' �G�N�Z���t�@�C���݂̂������ΏۂƂ���
        pos = InStrRev(strFileName, ".")
        If Not LCase(Mid(strFileName, pos + 1)) Like "xls*" Then
            ' ���̃t�@�C�������擾
            GoTo Continue
        End If
        
        ' ���t�@�C���iA1&�g�嗦100��.xlsm�j�͏���
        If strFileName = ThisWorkbook.Name Then
            GoTo Continue
        End If
    
        ' �G�N�Z���t�@�C�����J��
        fileAndPath = Path + "\" + strFileName
        Set Workbook = Workbooks.Open(fileAndPath)
    
        '��Ԑ擪�̃V�[�g���珇�Ƀ��[�v�������s��
        For Each sheet In ActiveWorkbook.Sheets
            sheet.Activate                 '�Ώۂ̃V�[�g���A�N�e�B�u�ɂ���
            ActiveSheet.Range("A1").Select '�V�[�g��A1��I������
            ActiveWindow.Zoom = 100        '�g��{����100�ɐݒ肷��
        Next sheet
        Sheets(1).Select
    
        ' �G�N�Z���t�@�C����ۑ����ĕ���
        Workbook.Save
        Workbook.Close
    
Continue:
    
        ' ���̃t�@�C�������擾
        strFileName = Dir()
    Loop
End Sub

' /**
' * A1�Z���ւ̃J�[�\���ړ��Ɗg�嗦100���̃T�u���[�`����
' * �T�u�t�H���_�ɑ΂��čċA�I�ɌĂяo���B
' *
' * @Param Path �����Ώۂ̃t�H���_�p�X
' *
' **/
Private Sub recursion(Path)
    Dim f As Object
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(Path).SubFolders
            Call setA1And100Per(f.Path)
        Next f
    End With
End Sub
