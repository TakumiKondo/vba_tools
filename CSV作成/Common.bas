Attribute VB_Name = "Common"
' �o�͐ݒ�t�H�[���Ăяo��
Sub showUserForm()
    UserForm1.Show
End Sub

' �o�͑ΏۃV�[�g�̎擾
Function getTargetSheets() As Object
    ' ���O����V�[�g
    Dim exclusionSheets As Object
    Set exclusionSheets = CreateObject("Scripting.Dictionary")
    exclusionSheets.Add "�v���E�v��", 1
    exclusionSheets.Add "�݌v", 1
    exclusionSheets.Add "�t�@�C���o��", 1
    
    Dim targetSheets As Object
    Set targetSheets = CreateObject("Scripting.Dictionary")
    
    ' �o�͑ΏۃV�[�g = �S�V�[�g���珜�O����V�[�g������������
    For Each st In ThisWorkbook.Worksheets
        If Not exclusionSheets.Exists(st.Name) Then
            targetSheets.Add st.Name, 1
        End If
    Next st
    
    Set getTargetSheets = targetSheets
End Function

' �o�͐�p�X�̎擾
Function outputPath() As String
    outputPath = ThisWorkbook.Path
End Function

